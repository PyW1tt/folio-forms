// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { createHmac } from "node:crypto";

import { env } from "@onlyoffice/env/server";

import { fail } from "../http/errors";

const schemaPageSize = 5;
const schemaQueryMaximumLength = 200;
type ExternalSchemaType = "string" | "number" | "boolean" | "null";
export interface ExternalSchemaItem {
  pointer: string;
  type: ExternalSchemaType;
}

const externalMockSchema = {
  account: {
    active: true,
    address: {
      city: "",
      country: "",
      postalCode: "",
    },
    consent: {
      privacy: true,
      terms: true,
    },
    contact: {
      email: "",
      phone: "",
    },
    contacts: [{ name: "" }],
    "display/name": "",
    id: "",
    loginCount: 0,
    score: 0,
    settings: {
      language: "",
      timezone: "",
    },
    "tilde~key": "",
  },
  ignoredObject: {
    nested: {
      value: null,
    },
  },
  person: {
    birthDate: "",
    "contact/details": {
      "line~1": "",
    },
    name: "",
  },
} as const;

function externalSchemaPointerSegment(value: string): string {
  return value.replaceAll("~", "~0").replaceAll("/", "~1");
}

export function externalSchemaLeafType(
  value: unknown
): ExternalSchemaType | null {
  if (value === null) {
    return "null";
  }
  if (typeof value === "string") {
    return "string";
  }
  if (typeof value === "number") {
    return "number";
  }
  if (typeof value === "boolean") {
    return "boolean";
  }
  return null;
}

function flattenExternalSchema(
  value: unknown,
  parentPointer = "",
  items: ExternalSchemaItem[] = []
): ExternalSchemaItem[] {
  const type = externalSchemaLeafType(value);
  if (type) {
    items.push({ pointer: parentPointer, type });
    return items;
  }
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return items;
  }
  for (const key of Object.keys(value).toSorted()) {
    const pointer = `${parentPointer}/${externalSchemaPointerSegment(key)}`;
    flattenExternalSchema(
      (value as Record<string, unknown>)[key],
      pointer,
      items
    );
  }
  return items;
}

export const externalSchemaItems = flattenExternalSchema(externalMockSchema);

function schemaCursor(query: string, offset: number): string {
  const payload = Buffer.from(JSON.stringify({ offset, query })).toString(
    "base64url"
  );
  const unsigned = `schema-v1.${payload}`;
  const signature = createHmac("sha256", env.EDITOR_CAPABILITY_SECRET)
    .update(unsigned)
    .digest("base64url");
  return `${unsigned}.${signature}`;
}

function schemaCursorOffset(cursor: string, query: string): number {
  const [version, payload, signature, extra] = cursor.split(".");
  if (!version || !payload || !signature || extra || version !== "schema-v1") {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  const unsigned = `${version}.${payload}`;
  const expected = createHmac("sha256", env.EDITOR_CAPABILITY_SECRET)
    .update(unsigned)
    .digest("base64url");
  if (signature !== expected) {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  let decoded: unknown;
  try {
    decoded = JSON.parse(Buffer.from(payload, "base64url").toString("utf-8"));
  } catch {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  if (
    !decoded ||
    typeof decoded !== "object" ||
    Array.isArray(decoded) ||
    !("offset" in decoded) ||
    !("query" in decoded) ||
    typeof decoded.offset !== "number" ||
    !Number.isInteger(decoded.offset) ||
    decoded.offset < 0 ||
    typeof decoded.query !== "string" ||
    decoded.query !== query ||
    decoded.offset > externalSchemaItems.length
  ) {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  return decoded.offset;
}

function schemaQueryValue(value: unknown): string {
  if (value === undefined) {
    return "";
  }
  if (typeof value !== "string") {
    fail(400, "invalid_schema_query", "Schema query must be a string");
  }
  const query = value.trim().toLowerCase();
  if (query.length > schemaQueryMaximumLength) {
    fail(400, "invalid_schema_query", "Schema query is too long");
  }
  return query;
}

export function schemaPage(
  queryValue: unknown,
  cursorValue: unknown
): { items: ExternalSchemaItem[]; nextCursor: string | null } {
  const query = schemaQueryValue(queryValue);
  const matching = externalSchemaItems.filter((item) =>
    item.pointer.toLowerCase().includes(query)
  );
  const offset =
    cursorValue === undefined
      ? 0
      : typeof cursorValue === "string"
        ? schemaCursorOffset(cursorValue, query)
        : fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  if (offset > matching.length) {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  const items = matching.slice(offset, offset + schemaPageSize);
  const nextOffset = offset + items.length;
  return {
    items,
    nextCursor:
      nextOffset < matching.length ? schemaCursor(query, nextOffset) : null,
  };
}
