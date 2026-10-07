import { AuditOutcome } from "@onlyoffice/db";

import { fail } from "../http/errors";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { idPattern } from "../http/input";

export const auditEventPageSize = 50;
export interface AuditEventCursor {
  createdAt: Date;
  id: string;
}

export function auditEventCursor(
  value: string | undefined
): AuditEventCursor | null {
  if (value === undefined) {
    return null;
  }
  try {
    const decoded = JSON.parse(
      Buffer.from(value, "base64url").toString("utf-8")
    ) as { createdAt?: unknown; id?: unknown };
    if (
      typeof decoded.createdAt !== "string" ||
      typeof decoded.id !== "string" ||
      !idPattern.test(decoded.id)
    ) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    const createdAt = new Date(decoded.createdAt);
    if (!Number.isFinite(createdAt.getTime())) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    return { createdAt, id: decoded.id };
  } catch {
    fail(400, "invalid_request", "cursor is invalid");
  }
}

export function auditEventCursorValue(event: AuditEventCursor): string {
  return Buffer.from(
    JSON.stringify({
      createdAt: event.createdAt.toISOString(),
      id: event.id,
    })
  ).toString("base64url");
}

export function auditFilterValue(
  value: string | undefined,
  key: string,
  maximumLength: number
): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  const trimmed = value.trim();
  if (!trimmed || trimmed.length > maximumLength) {
    fail(400, "invalid_request", `${key} is invalid`);
  }
  return trimmed;
}

export function auditOutcomeValue(
  value: string | undefined
): AuditOutcome | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (value !== AuditOutcome.failure && value !== AuditOutcome.success) {
    fail(400, "invalid_request", "outcome is invalid");
  }
  return value;
}
