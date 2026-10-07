import { isDeepStrictEqual } from "node:util";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Prisma } from "@onlyoffice/db";
import { FieldType, prisma, FillMethod } from "@onlyoffice/db";

import type { OnlyOfficeFieldDisplayMetadata } from "../documents/office-values";
import {
  onlyOfficeFieldValueNeedsMetadata,
  onlyOfficeFieldDisplayMetadata,
  normalizeOnlyOfficeDisplayValue,
} from "../documents/office-values";
import { fail } from "../http/errors";
import { jsonRecord } from "../http/input";
import type {
  JsonRecord,
  FormWithDocuments,
  ResponseWithSnapshot,
} from "../model-types";
import { handoffUnavailable } from "../prefill/input";
import { readObject } from "../storage";
import { isValidDateFieldValue } from "./date-value";

export const maxResponseDataBytes = 256 * 1024;
const maxResponseTextLength = 10_000;

export function jsonValue(value: unknown): Prisma.InputJsonValue {
  return value as Prisma.InputJsonValue;
}

export function manifestFieldValueMatches(
  field: { options: unknown; type: FieldType },
  value: unknown
): boolean {
  if (field.type === FieldType.checkbox) {
    return typeof value === "boolean";
  }
  if (field.type === FieldType.dropdown) {
    const optionValues = Array.isArray(field.options)
      ? field.options.flatMap((option) => {
          if (!option || typeof option !== "object" || Array.isArray(option)) {
            return [];
          }
          const optionValue = (option as Record<string, unknown>).value;
          return typeof optionValue === "string" ? [optionValue] : [];
        })
      : [];
    return typeof value === "string" && optionValues.includes(value);
  }
  if (field.type === FieldType.date) {
    return typeof value === "string" && isValidDateFieldValue(value);
  }
  if (field.type === FieldType.combo || field.type === FieldType.text) {
    return typeof value === "string" && value.length <= maxResponseTextLength;
  }
  return false;
}

export function validatePrefillValuesAgainstManifest(
  values: JsonRecord,
  fields: readonly { options: unknown; tag: string; type: FieldType }[]
): void {
  const serialized = JSON.stringify(values);
  if (new TextEncoder().encode(serialized).byteLength > maxResponseDataBytes) {
    handoffUnavailable();
  }
  const fieldsByTag = new Map(fields.map((field) => [field.tag, field]));
  for (const [tag, value] of Object.entries(values)) {
    const field = fieldsByTag.get(tag);
    if (!field || !manifestFieldValueMatches(field, value)) {
      handoffUnavailable();
    }
  }
}

export async function normalizeResponseData(
  form: FormWithDocuments,
  response: ResponseWithSnapshot,
  inputData: unknown,
  requireRequired = false,
  onlyOfficeCompatibility = false,
  canonicalDateFieldsInput: unknown = undefined
): Promise<JsonRecord> {
  const suppliedData = { ...jsonRecord(inputData) };
  const canonicalDateFields = new Set<string>();
  if (canonicalDateFieldsInput !== undefined) {
    if (
      !Array.isArray(canonicalDateFieldsInput) ||
      !canonicalDateFieldsInput.every(
        (field): field is string => typeof field === "string"
      )
    ) {
      fail(
        400,
        "invalid_request",
        "canonicalDateFields must be an array of field tags"
      );
    }
    for (const field of canonicalDateFieldsInput as string[]) {
      if (canonicalDateFields.has(field)) {
        fail(
          400,
          "invalid_request",
          "canonicalDateFields must not contain duplicates"
        );
      }
      canonicalDateFields.add(field);
    }
  }
  const publishedTemplate = form.publishedTemplate;
  if (!publishedTemplate?.objectKey) {
    fail(409, "not_published", "This form has not been published");
  }
  const manifest = await prisma.fieldManifest.findUnique({
    include: { fields: { orderBy: { tag: "asc" } } },
    where: { publishedTemplateId: publishedTemplate.id },
  });
  if (!manifest) {
    fail(500, "internal_error", "The published Field Manifest is unavailable");
  }
  const fieldsByTag = new Map(
    manifest.fields.map((field) => [field.tag, field])
  );
  for (const field of canonicalDateFields) {
    const definition = fieldsByTag.get(field);
    if (
      !definition ||
      definition.type !== FieldType.date ||
      !Object.hasOwn(suppliedData, field) ||
      typeof suppliedData[field] !== "string" ||
      !isValidDateFieldValue(suppliedData[field] as string)
    ) {
      fail(
        422,
        "invalid_response_data",
        `${field} is not a supplied canonical date field`
      );
    }
  }
  const unknownFields = Object.keys(suppliedData).filter(
    (field) => !fieldsByTag.has(field)
  );
  if (unknownFields.length > 0) {
    fail(
      422,
      "invalid_response_data",
      `Unknown form field(s): ${unknownFields.join(", ")}`
    );
  }
  let onlyOfficeMetadataByTag = new Map<
    string,
    OnlyOfficeFieldDisplayMetadata
  >();
  if (onlyOfficeCompatibility) {
    const fieldsNeedingMetadata = manifest.fields.filter(
      (field) =>
        !canonicalDateFields.has(field.tag) &&
        onlyOfficeFieldValueNeedsMetadata(field, suppliedData[field.tag])
    );
    if (fieldsNeedingMetadata.length > 0) {
      onlyOfficeMetadataByTag = onlyOfficeFieldDisplayMetadata(
        await readObject(publishedTemplate.objectKey)
      );
      for (const field of manifest.fields) {
        if (
          Object.hasOwn(suppliedData, field.tag) &&
          !canonicalDateFields.has(field.tag)
        ) {
          suppliedData[field.tag] = normalizeOnlyOfficeDisplayValue(
            field,
            suppliedData[field.tag],
            onlyOfficeMetadataByTag.get(field.tag)
          );
        }
      }
    }
  }
  let data = onlyOfficeCompatibility
    ? { ...jsonRecord(response.draftData ?? {}), ...suppliedData }
    : { ...suppliedData };
  const snapshot = response.prefillSnapshot;
  if (snapshot) {
    const snapshotData = jsonRecord(snapshot.values);
    const lockedFields = jsonRecord(snapshot.lockedFields);
    for (const [field, value] of Object.entries(snapshotData)) {
      if (lockedFields[field] === true) {
        data[field] = value;
      }
    }
  }
  const pictureTags = new Set(
    manifest.fields
      .filter((field) => field.type === FieldType.picture)
      .map((field) => field.tag)
  );
  data = Object.fromEntries(
    Object.entries(data).filter(([fieldTag]) => !pictureTags.has(fieldTag))
  );
  const serialized = JSON.stringify(data);
  if (new TextEncoder().encode(serialized).byteLength > maxResponseDataBytes) {
    fail(413, "response_too_large", "Response data exceeds the size limit");
  }

  for (const [fieldTag, value] of Object.entries(data)) {
    const field = fieldsByTag.get(fieldTag);
    if (field && value !== null && !manifestFieldValueMatches(field, value)) {
      fail(
        422,
        "invalid_response_data",
        `${fieldTag} does not match the published Field Manifest`
      );
    }
  }
  if (requireRequired) {
    for (const field of manifest.fields) {
      if (field.type === FieldType.picture) {
        continue;
      }
      const value = data[field.tag];
      const blankOptionSelected =
        typeof value === "string" &&
        value.trim().length === 0 &&
        (field.type === FieldType.dropdown || field.type === FieldType.combo) &&
        Array.isArray(field.options) &&
        field.options.some(
          (option) =>
            option !== null &&
            typeof option === "object" &&
            !Array.isArray(option) &&
            (option as Record<string, unknown>).value === value
        );
      if (
        field.required &&
        (!Object.hasOwn(data, field.tag) ||
          value === null ||
          (typeof value === "string" &&
            value.trim().length === 0 &&
            !blankOptionSelected) ||
          (field.type === FieldType.checkbox && value !== true))
      ) {
        fail(
          422,
          "invalid_response_data",
          `${field.tag} is required by the published Field Manifest`
        );
      }
    }
  }
  if (snapshot) {
    const snapshotData = jsonRecord(snapshot.values);
    const lockedFields = jsonRecord(snapshot.lockedFields);
    for (const [field, value] of Object.entries(snapshotData)) {
      if (lockedFields[field] === true && Object.hasOwn(suppliedData, field)) {
        const suppliedValue = suppliedData[field];
        if (suppliedValue !== value) {
          fail(422, "invalid_response_data", `${field} is locked by Prefill`);
        }
      }
    }
  }
  return data;
}
export function requireCurrentFillMethod(
  form: FormWithDocuments,
  input: JsonRecord
): void {
  const requestedFillMethod = input.fillMethod ?? FillMethod.onlyoffice;
  if (
    requestedFillMethod !== FillMethod.native &&
    requestedFillMethod !== FillMethod.onlyoffice
  ) {
    fail(400, "invalid_request", "fillMethod must be native or onlyoffice");
  }
  if (requestedFillMethod !== form.fillMethod) {
    fail(409, "fill_method_changed", "The form Fill Method changed");
  }
}
export function changedResponseData(
  previous: JsonRecord,
  next: JsonRecord
): JsonRecord {
  const keys = new Set([...Object.keys(previous), ...Object.keys(next)]);
  return Object.fromEntries(
    [...keys]
      .filter((key) => !isDeepStrictEqual(previous[key], next[key]))
      .map((key) => [key, next[key] ?? null])
  );
}
