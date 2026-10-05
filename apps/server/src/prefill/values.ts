import type { PrefillPolicy } from "@onlyoffice/db";
import { FormStatus } from "@onlyoffice/db";

import {
  externalSchemaItems,
  externalSchemaLeafType,
} from "../forms/field-schema";
import { jsonRecord } from "../http/input";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { PrefillSnapshot, JsonRecord } from "../model-types";
import { handoffUnavailable } from "./input";

export function editableFieldsForSnapshot(
  snapshot: PrefillSnapshot
): JsonRecord {
  const lockedFields = jsonRecord(snapshot.lockedFields);
  return Object.fromEntries(
    Object.entries(lockedFields).map(([field, locked]) => [
      field,
      locked !== true,
    ])
  );
}

function externalPointerSegments(pointer: string): string[] | null {
  if (!pointer.startsWith("/")) {
    return null;
  }
  if (pointer === "/") {
    return [""];
  }
  return pointer
    .slice(1)
    .split("/")
    .map((segment) => {
      let decoded = "";
      for (let index = 0; index < segment.length; index += 1) {
        const character = segment[index];
        if (character !== "~") {
          decoded += character;
          continue;
        }
        const escape = segment[index + 1];
        if (escape === "0") {
          decoded += "~";
        } else if (escape === "1") {
          decoded += "/";
        } else {
          return "";
        }
        index += 1;
      }
      return decoded;
    });
}

function externalValueAtPointer(values: JsonRecord, pointer: string): unknown {
  if (Object.hasOwn(values, pointer)) {
    return values[pointer];
  }
  const segments = externalPointerSegments(pointer);
  if (!segments) {
    return undefined;
  }
  let current: unknown = values;
  for (const segment of segments) {
    if (
      !current ||
      typeof current !== "object" ||
      Array.isArray(current) ||
      !Object.hasOwn(current, segment)
    ) {
      return undefined;
    }
    current = (current as JsonRecord)[segment];
  }
  return current;
}

export function externalValueMatchesSchema(
  pointer: string,
  value: unknown
): boolean {
  const schemaItem = externalSchemaItems.find(
    (candidate) => candidate.pointer === pointer
  );
  return (
    schemaItem !== undefined &&
    externalSchemaLeafType(value) === schemaItem.type
  );
}
export function filteredPrefillValues(
  values: JsonRecord,
  fields: readonly { pointer: string; tag: string }[]
): JsonRecord {
  const filtered: JsonRecord = {};
  for (const field of fields) {
    const value = externalValueAtPointer(values, field.pointer);
    if (value === undefined) {
      continue;
    }
    if (!externalValueMatchesSchema(field.pointer, value)) {
      handoffUnavailable();
    }
    filtered[field.tag] = value;
  }
  return filtered;
}

export function publishedPrefillConfiguration(
  form: {
    publishedTemplate: {
      contentHash: string;
      id: string;
      prefillConfiguration: {
        configurationHash: string;
        fields: { pointer: string; policy: PrefillPolicy; tag: string }[];
        publishedTemplateId: string | null;
      } | null;
    } | null;
    status: FormStatus;
  } | null
): {
  configurationHash: string;
  fields: { pointer: string; policy: PrefillPolicy; tag: string }[];
  publishedTemplateId: string;
  templateId: string;
} | null {
  const template = form?.publishedTemplate;
  const configuration = template?.prefillConfiguration;
  if (
    !template ||
    !configuration ||
    configuration.fields.length === 0 ||
    configuration.publishedTemplateId !== template.id ||
    configuration.configurationHash !== template.contentHash ||
    form.status !== FormStatus.published
  ) {
    return null;
  }
  return {
    configurationHash: configuration.configurationHash,
    fields: configuration.fields,
    publishedTemplateId: configuration.publishedTemplateId,
    templateId: template.id,
  };
}
export function lockedPrefillForSnapshot(
  snapshot: PrefillSnapshot | null
): { data: JsonRecord; editableFields: JsonRecord } | undefined {
  if (!snapshot) {
    return;
  }
  const values = jsonRecord(snapshot.values);
  const lockedFields = jsonRecord(snapshot.lockedFields);
  const data = Object.fromEntries(
    Object.entries(values).filter(([tag]) => lockedFields[tag] === true)
  );
  return Object.keys(data).length > 0
    ? { data, editableFields: editableFieldsForSnapshot(snapshot) }
    : undefined;
}
