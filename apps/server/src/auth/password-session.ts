// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { createHmac } from "node:crypto";

import { auth } from "@onlyoffice/auth";
import { prisma } from "@onlyoffice/db";
import { env } from "@onlyoffice/env/server";

import type { AiAuthoringSessions } from "../ai-authoring";
import type { JsonRecord } from "../model-types";
import type { Identity } from "./identity";
import { identityFor, normalizeEmail } from "./identity";

const loginFailureLimit = 5;
const loginFailureWindowMs = 15 * 60_000;
export const passwordMinimumLength = 12;
export const passwordMaximumLength = 128;

function loginDigest(kind: "email" | "ip", value: string): string {
  return createHmac("sha256", env.BETTER_AUTH_SECRET)
    .update(`${kind}:${value}`)
    .digest("hex");
}

function loginFailureKey(email: string, sourceIp: string) {
  return {
    emailDigest: loginDigest("email", email),
    ipDigest: loginDigest("ip", sourceIp.trim().toLowerCase() || "unknown"),
  };
}

async function reserveLoginAttempt(
  emailDigest: string,
  ipDigest: string
): Promise<number> {
  const now = new Date();
  await prisma.loginFailure.deleteMany({
    where: { expiresAt: { lte: now } },
  });
  const expiresAt = new Date(now.getTime() + loginFailureWindowMs);
  const [failure] = await prisma.$queryRaw<{ attempts: number }[]>`
    INSERT INTO "login_failures" AS "login_failure" (
      "id",
      "email_digest",
      "ip_digest",
      "attempts",
      "window_started_at",
      "expires_at",
      "updated_at"
    )
    VALUES (
      ${crypto.randomUUID()}::uuid,
      ${emailDigest},
      ${ipDigest},
      1,
      ${now},
      ${expiresAt},
      ${now}
    )
    ON CONFLICT ("email_digest", "ip_digest") DO UPDATE SET
      "attempts" = CASE
        WHEN "login_failure"."expires_at" <= ${now} THEN 1
        ELSE "login_failure"."attempts" + 1
      END,
      "window_started_at" = CASE
        WHEN "login_failure"."expires_at" <= ${now} THEN ${now}
        ELSE "login_failure"."window_started_at"
      END,
      "expires_at" = CASE
        WHEN "login_failure"."expires_at" <= ${now} THEN ${expiresAt}
        ELSE "login_failure"."expires_at"
      END,
      "updated_at" = ${now}
    RETURNING "attempts"
  `;
  return failure?.attempts ?? loginFailureLimit + 1;
}

async function clearLoginFailures(
  emailDigest: string,
  ipDigest: string
): Promise<void> {
  await prisma.loginFailure.deleteMany({ where: { emailDigest, ipDigest } });
}

function loginError(
  status: 401 | 429,
  error: "invalid_credentials" | "login_throttled"
): globalThis.Response {
  return Response.json(
    { error, message: "Email or password is invalid" },
    {
      headers: status === 429 ? { "Retry-After": "900" } : undefined,
      status,
    }
  );
}

async function withAiAuthoringSessionsEnding<T>(
  aiAuthoring: AiAuthoringSessions,
  ownerSessionIds: readonly string[],
  revokeOwnerSessions: () => Promise<T>
): Promise<T> {
  const uniqueSessionIds = [...new Set(ownerSessionIds)];
  if (uniqueSessionIds.length === 0) {
    return await revokeOwnerSessions();
  }
  let revocation: Promise<T> | undefined;
  const revokeDeferred = async (): Promise<T> => {
    await Promise.resolve();
    return await revokeOwnerSessions();
  };
  const revokeOnce = (): Promise<T> => (revocation ??= revokeDeferred());
  const endings = await Promise.allSettled(
    uniqueSessionIds.map((sessionId) =>
      aiAuthoring.endForSession(sessionId, revokeOnce)
    )
  );
  const errors = [
    ...new Set(
      endings.flatMap((ending) =>
        ending.status === "rejected" ? [ending.reason] : []
      )
    ),
  ];
  if (errors.length === 1) {
    throw errors[0];
  }
  if (errors.length > 1) {
    throw new AggregateError(errors, "AI authoring session cleanup failed");
  }
  return await (revocation ?? revokeOnce());
}

export async function endAiAuthoringSessions(
  aiAuthoring: AiAuthoringSessions,
  ownerSessionIds: readonly string[]
): Promise<void> {
  await withAiAuthoringSessionsEnding(aiAuthoring, ownerSessionIds, () =>
    Promise.resolve()
  );
}

export async function isAuthoringOwnerSessionCurrent(
  request: Request,
  owner: Identity
): Promise<boolean> {
  const currentIdentity = await identityFor(request);
  return (
    currentIdentity !== null &&
    currentIdentity.id === owner.id &&
    currentIdentity.sessionId === owner.sessionId &&
    currentIdentity.role === "admin" &&
    (!currentIdentity.mustChangePassword || currentIdentity.isSso)
  );
}

export async function handleEmailSignIn(
  request: Request,
  body: unknown,
  sourceIp: string,
  aiAuthoring: AiAuthoringSessions
): Promise<globalThis.Response> {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    return loginError(401, "invalid_credentials");
  }
  const input = body as JsonRecord;
  if (typeof input.email !== "string" || typeof input.password !== "string") {
    return loginError(401, "invalid_credentials");
  }
  const email = normalizeEmail(input.email);
  const { emailDigest, ipDigest } = loginFailureKey(email, sourceIp);
  if ((await reserveLoginAttempt(emailDigest, ipDigest)) > loginFailureLimit) {
    return loginError(429, "login_throttled");
  }
  if (
    input.password.length < passwordMinimumLength ||
    input.password.length > passwordMaximumLength
  ) {
    return loginError(401, "invalid_credentials");
  }

  const headers = new Headers(request.headers);
  headers.delete("content-length");
  const authResponse = await auth.handler(
    new Request(request.url, {
      body: JSON.stringify({ email, password: input.password }),
      headers,
      method: "POST",
    })
  );
  if (!authResponse.ok) {
    return loginError(401, "invalid_credentials");
  }

  const user = await prisma.user.findUnique({ where: { email } });
  if (!user?.enabled) {
    if (user) {
      const sessions = await prisma.session.findMany({
        select: { id: true },
        where: { userId: user.id },
      });
      await withAiAuthoringSessionsEnding(
        aiAuthoring,
        sessions.map(({ id }) => id),
        () => prisma.session.deleteMany({ where: { userId: user.id } })
      );
    }
    return loginError(401, "invalid_credentials");
  }

  await clearLoginFailures(emailDigest, ipDigest);
  return authResponse;
}
