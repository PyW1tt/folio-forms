import { prisma, FieldType } from "@onlyoffice/db";
import { strToU8, zipSync } from "fflate";

import {
  nativePdfScalarControlXml,
  overlayNativeResponseDocument,
} from "../documents/native-document";
import {
  safeTemplateArchive,
  reachableTemplateControlParts,
  templateArchiveText,
} from "../documents/package";
import { overlayNativeResponsePictures } from "../documents/picture-overlay";
import { fail } from "../http/errors";
import type { JsonRecord } from "../model-types";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { operationDocumentKey } from "../operations/lifecycle";
import { objectExists, readObject } from "../storage";
import type { NativePictureUpload } from "./native-request";

export async function nativePdfDocument(
  documentKey: string
): Promise<Uint8Array> {
  const sourceObjectKey = await operationDocumentKey(documentKey);
  if (!sourceObjectKey || !(await objectExists(sourceObjectKey))) {
    fail(404, "not_found", "Document was not found");
  }
  const { archive, xmlPaths } = safeTemplateArchive(
    await readObject(sourceObjectKey)
  );
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  for (const archivePath of controlParts) {
    const originalXml = templateArchiveText(archive, archivePath);
    const transformedXml = nativePdfScalarControlXml(originalXml);
    if (transformedXml !== originalXml) {
      archive[archivePath] = strToU8(
        transformedXml.replace(
          /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
          'encoding="UTF-8"'
        )
      );
    }
  }
  return zipSync(archive);
}
export async function nativeResponseDocument(
  publishedTemplateId: string,
  data: JsonRecord,
  baseDocumentBytes?: Uint8Array,
  pictures: ReadonlyMap<string, NativePictureUpload> = new Map()
): Promise<Uint8Array> {
  const [publishedTemplate, manifest] = await Promise.all([
    prisma.publishedTemplate.findUnique({
      select: { objectKey: true },
      where: { id: publishedTemplateId },
    }),
    prisma.fieldManifest.findUnique({
      include: {
        fields: {
          select: {
            options: true,
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
    }),
  ]);
  if (!publishedTemplate?.objectKey || !manifest) {
    fail(409, "document_unavailable", "The published document is unavailable");
  }
  const scalarTags = new Set(
    manifest.fields
      .filter(({ type }) => type !== FieldType.picture)
      .map(({ tag }) => tag)
  );
  const baseBytes =
    baseDocumentBytes ?? (await readObject(publishedTemplate.objectKey));
  return overlayNativeResponsePictures(
    overlayNativeResponseDocument(baseBytes, manifest.fields, data, scalarTags),
    manifest.fields,
    pictures
  );
}
