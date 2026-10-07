import { FieldType, prisma } from "@onlyoffice/db";

import { parseTemplateFields } from "../documents/fields";
import { nativeFlattenAlternateFieldsXml } from "../documents/native-alternate";
import {
  templatePartsHaveUnsupportedDateSettings,
  templatePartsHaveNestedControls,
} from "../documents/native-eligibility";
import {
  safeTemplateArchive,
  reachableTemplateControlParts,
  templateArchiveText,
} from "../documents/package";
import { responsePicturePresence } from "../documents/pictures";
import type { ResponsePictureManifestField } from "../documents/pictures";
import { fail } from "../http/errors";
import { publicIdPattern } from "../http/input";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { FormWithDocuments, JsonRecord } from "../model-types";
import { readObject } from "../storage";

interface FormCounts {
  activeDraftCount: number;
  submissionCount: number;
}

export function formDto(
  form: FormWithDocuments,
  { activeDraftCount, submissionCount }: FormCounts
): JsonRecord {
  return {
    activeDraftCount,
    createdAt: form.createdAt,
    description: form.description ?? "",
    fillMethod: form.fillMethod,
    hasTemplateDraft: Boolean(form.templateDraft),
    publicId: form.publicId,
    status: form.status,
    submissionCount,
    title: form.title,
    updatedAt: form.updatedAt,
    version: form.version,
  };
}
function supportsNativeFill(fields: readonly { type: FieldType }[]): boolean {
  return fields.length > 0;
}

interface ManifestDisplayField {
  label: string;
  placeholder: string | null;
  position: number;
  tag: string;
}

export async function manifestFieldsWithDocumentMetadata<
  T extends ManifestDisplayField,
>(
  objectKey: string,
  displayMetadataVersion: number,
  fields: T[],
  documentBytes?: Uint8Array
): Promise<T[]> {
  if (displayMetadataVersion !== 0) {
    return fields;
  }
  const documentFields = parseTemplateFields(
    documentBytes ?? (await readObject(objectKey))
  );
  if (documentFields.length !== fields.length) {
    fail(
      500,
      "internal_error",
      "The published Field Manifest does not match its document"
    );
  }
  const fieldsByTag = new Map(fields.map((field) => [field.tag, field]));
  return documentFields.map((documentField, position) => {
    const field = fieldsByTag.get(documentField.tag);
    if (!field) {
      fail(
        500,
        "internal_error",
        "The published Field Manifest does not match its document"
      );
    }
    return {
      ...field,
      label: documentField.label,
      placeholder: documentField.placeholder,
      position,
    };
  });
}

export async function supportsNativeTemplate(
  objectKey: string | null | undefined,
  fields: readonly { tag: string; type: FieldType }[],
  documentBytes?: Uint8Array
): Promise<boolean> {
  if (!objectKey || !supportsNativeFill(fields)) {
    return false;
  }
  const { archive, xmlPaths } = safeTemplateArchive(
    documentBytes ?? (await readObject(objectKey))
  );
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  if (
    templatePartsHaveUnsupportedDateSettings(archive, controlParts) ||
    templatePartsHaveNestedControls(archive, controlParts)
  ) {
    return false;
  }
  const fieldTags = new Set(fields.map(({ tag }) => tag));
  for (const archivePath of controlParts) {
    if (
      !nativeFlattenAlternateFieldsXml(
        templateArchiveText(archive, archivePath),
        fieldTags,
        true
      ).supported
    ) {
      return false;
    }
  }
  return true;
}
export async function nativeFields(
  publishedTemplateId: string,
  responseDocumentBytes: Uint8Array
): Promise<{
  fields: {
    label: string;
    options: { displayText: string; value: string }[];
    pictureMaxBytes: number | null;
    pictureMaxHeight: number | null;
    pictureMaxWidth: number | null;
    placeholder: string | null;
    position: number;
    required: boolean;
    tag: string;
    type: FieldType;
  }[];
  pictures: Record<string, boolean>;
}> {
  const [manifest, publishedTemplate] = await Promise.all([
    prisma.fieldManifest.findUnique({
      include: { fields: { orderBy: { position: "asc" } } },
      where: { publishedTemplateId },
    }),
    prisma.publishedTemplate.findUnique({
      select: { objectKey: true },
      where: { id: publishedTemplateId },
    }),
  ]);
  if (!manifest || !publishedTemplate?.objectKey) {
    fail(
      409,
      "native_fill_unsupported",
      "Native filling requires supported fields"
    );
  }
  const documentBytes = await readObject(publishedTemplate.objectKey);
  if (
    !(await supportsNativeTemplate(
      publishedTemplate.objectKey,
      manifest.fields,
      documentBytes
    ))
  ) {
    fail(
      409,
      "native_fill_unsupported",
      "Native filling requires supported fields"
    );
  }
  const fields = await manifestFieldsWithDocumentMetadata(
    publishedTemplate.objectKey,
    manifest.displayMetadataVersion,
    manifest.fields,
    documentBytes
  );
  return {
    fields: fields.map((field) => ({
      ...receiptField(field),
      pictureMaxBytes: field.pictureMaxBytes,
      pictureMaxHeight: field.pictureMaxHeight,
      pictureMaxWidth: field.pictureMaxWidth,
      required: field.required,
    })),
    pictures: responsePicturePresence(
      responseDocumentBytes,
      manifest.fields.filter(({ type }) => type === FieldType.picture)
    ),
  };
}

export async function nativePictureManifestFields(
  publishedTemplateId: string
): Promise<ResponsePictureManifestField[]> {
  const manifest = await prisma.fieldManifest.findUnique({
    select: {
      fields: {
        select: {
          pictureMaxBytes: true,
          pictureMaxHeight: true,
          pictureMaxWidth: true,
          required: true,
          tag: true,
          type: true,
        },
      },
    },
    where: { publishedTemplateId },
  });
  if (!manifest) {
    fail(500, "internal_error", "The published Field Manifest is unavailable");
  }
  return manifest.fields.filter(({ type }) => type === FieldType.picture);
}
export async function findFormByPublicId(
  publicId: string
): Promise<FormWithDocuments> {
  if (!publicIdPattern.test(publicId)) {
    fail(404, "not_found", "Form was not found");
  }
  const form = await prisma.form.findUnique({
    include: { publishedTemplate: true, templateDraft: true },
    where: { publicId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  return form;
}
export function receiptField(
  field: ManifestDisplayField & { options: unknown; type: FieldType }
) {
  return {
    label: field.label,
    options: Array.isArray(field.options)
      ? field.options.flatMap((option) => {
          if (!option || typeof option !== "object" || Array.isArray(option)) {
            return [];
          }
          const optionRecord = option as Record<string, unknown>;
          const displayText = optionRecord.displayText;
          const value = optionRecord.value;
          return typeof displayText === "string" && typeof value === "string"
            ? [{ displayText, value }]
            : [];
        })
      : [],
    placeholder: field.placeholder,
    position: field.position,
    tag: field.tag,
    type: field.type,
  };
}
