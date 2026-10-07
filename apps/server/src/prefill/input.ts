// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { createHash, timingSafeEqual } from "node:crypto";

import { normalizeEmail } from "../auth/identity";
import { fail } from "../http/errors";
import {
  accountEmailMaximumLength,
  accountEmailPattern,
  asRecord,
  readRequestBytes,
} from "../http/input";
import type { JsonRecord } from "../model-types";
import type { PrefillHandoffCreateInput } from "./create-launch";

export function configuredPrefillReturnUrl(value: string): string {
  const url = new URL(value);
  if (
    !["http:", "https:"].includes(url.protocol) ||
    url.username ||
    url.password
  ) {
    throw new Error(
      "PREFILL_RETURN_URL must be an HTTP(S) URL without credentials"
    );
  }
  return url.toString();
}
export const handoffCodeLifetimeMs = 120_000;
export const pendingClaimLifetimeSeconds = 10 * 60;
export const handoffCodeMaximumLength = 256;
const handoffExternalReferenceMaximumLength = 512;
const prefillHandoffBodyMaximumBytes = 512 * 1024;
const pendingClaimCookieName = "__Host-folio-pending-claim";

export function prefillHandoffSecretMatches(
  expectedSecret: string,
  receivedSecret: string | null
): boolean {
  if (!receivedSecret) {
    return false;
  }
  const expectedDigest = createHash("sha256").update(expectedSecret).digest();
  const receivedDigest = createHash("sha256").update(receivedSecret).digest();
  return timingSafeEqual(expectedDigest, receivedDigest);
}

export function pendingClaimCookie(value: string, maxAge: number): string {
  return `${pendingClaimCookieName}=${value}; Path=/; Max-Age=${maxAge}; HttpOnly; Secure; SameSite=Lax`;
}

export function pendingClaimFor(request: Request): string | undefined {
  const cookieHeader = request.headers.get("cookie");
  if (!cookieHeader) {
    return undefined;
  }
  for (const part of cookieHeader.split(";")) {
    const separator = part.indexOf("=");
    if (separator === -1) {
      continue;
    }
    const name = part.slice(0, separator).trim();
    if (name !== pendingClaimCookieName) {
      continue;
    }
    const value = part.slice(separator + 1).trim();
    return value.length > 0 && value.length <= handoffCodeMaximumLength
      ? value
      : undefined;
  }
  return undefined;
}

export function handoffUnavailable(): never {
  fail(409, "handoff_unavailable", "The prefill handoff is unavailable");
}

export function prefillRequired(): never {
  fail(409, "prefill_required", "A prefill handoff is required");
}

function handoffEmail(value: unknown): string {
  if (typeof value !== "string") {
    fail(400, "invalid_request", "email must be a valid email address");
  }
  const email = normalizeEmail(value);
  if (
    email.length > accountEmailMaximumLength ||
    !accountEmailPattern.test(email)
  ) {
    fail(400, "invalid_request", "email must be a valid email address");
  }
  return email;
}

function handoffExternalReference(value: unknown): string {
  if (
    typeof value !== "string" ||
    value.trim().length === 0 ||
    value.trim().length > handoffExternalReferenceMaximumLength
  ) {
    fail(
      400,
      "invalid_request",
      "externalReference must be a non-empty bounded string"
    );
  }
  return value.trim();
}

function handoffCode(value: unknown): string {
  if (
    typeof value !== "string" ||
    value.trim().length === 0 ||
    value.trim().length > handoffCodeMaximumLength
  ) {
    fail(400, "invalid_request", "code is required");
  }
  return value.trim();
}

export function handoffCreateInput(
  input: JsonRecord
): PrefillHandoffCreateInput {
  const keys = Object.keys(input);
  const expectedKeys = new Set([
    "email",
    "externalReference",
    "publicId",
    "values",
  ]);
  if (
    keys.length !== expectedKeys.size ||
    keys.some((key) => !expectedKeys.has(key))
  ) {
    fail(
      400,
      "invalid_request",
      "publicId, email, externalReference, and values are required"
    );
  }
  if (typeof input.publicId !== "string") {
    fail(404, "not_found", "Form was not found");
  }
  const values = asRecord(input.values, "values must be a JSON object");
  return {
    email: handoffEmail(input.email),
    externalReference: handoffExternalReference(input.externalReference),
    publicId: input.publicId,
    values,
  };
}
export function handoffStatusInput(input: JsonRecord): {
  externalReference: string;
} {
  if (
    Object.keys(input).length !== 1 ||
    Object.keys(input)[0] !== "externalReference"
  ) {
    fail(400, "invalid_request", "externalReference is required");
  }
  return {
    externalReference: handoffExternalReference(input.externalReference),
  };
}

export function requireTopLevelNavigation(request: Request): void {
  const fetchMode = request.headers.get("sec-fetch-mode");
  const fetchDestination = request.headers.get("sec-fetch-dest");
  if (
    (fetchMode !== null && fetchMode !== "navigate") ||
    (fetchDestination !== null && fetchDestination !== "document")
  ) {
    fail(400, "invalid_request", "The handoff must be a top-level navigation");
  }
}

export async function readPrefillHandoffCode(
  request: Request
): Promise<string> {
  const contentType =
    request.headers
      .get("content-type")
      ?.split(";", 1)[0]
      ?.trim()
      .toLowerCase() ?? "";
  if (
    contentType !== "application/x-www-form-urlencoded" &&
    contentType !== "multipart/form-data"
  ) {
    fail(415, "invalid_request", "The handoff code body format is unsupported");
  }
  const bytes = await readRequestBytes(
    request,
    prefillHandoffBodyMaximumBytes,
    "The handoff code is required"
  );
  if (contentType === "application/x-www-form-urlencoded") {
    const entries = new Map<string, string>();
    for (const [key, value] of new URLSearchParams(
      new TextDecoder().decode(bytes)
    )) {
      if (key !== "code" || entries.has(key)) {
        fail(400, "invalid_request", "Only code is accepted");
      }
      entries.set(key, value);
    }
    return handoffCode(entries.get("code"));
  }
  if (contentType === "multipart/form-data") {
    const multipartRequest = new Request(request.url, {
      body: bytes,
      headers: {
        "content-type": request.headers.get("content-type") as string,
      },
      method: "POST",
    });
    let formData: Awaited<ReturnType<typeof multipartRequest.formData>>;
    try {
      formData = await multipartRequest.formData();
    } catch {
      fail(400, "invalid_request", "Multipart form data is invalid");
    }
    let code: string | undefined;
    for (const [key, value] of formData.entries()) {
      if (key !== "code" || typeof value !== "string" || code !== undefined) {
        fail(400, "invalid_request", "Only code is accepted");
      }
      code = value;
    }
    return handoffCode(code);
  }
  fail(415, "invalid_request", "The handoff code body format is unsupported");
}
