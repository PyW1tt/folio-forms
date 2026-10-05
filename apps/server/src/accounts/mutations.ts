import { randomBytes } from "node:crypto";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { Prisma, prisma } from "@onlyoffice/db";

import type { AccountAuditAction } from "../audit/events";
import { createAccountFailureAudit } from "../audit/events";
import { normalizeEmail, identityFor, requireAdmin } from "../auth/identity";
import type { Identity } from "../auth/identity";
import { databaseErrorCode } from "../db-errors";
import { fail, HttpError } from "../http/errors";
import { accountEmailMaximumLength, accountEmailPattern } from "../http/input";

export const accountUserPageSize = 20;
export const accountNameMaximumLength = 120;
// ponytail: one global account lock caps mutation throughput; shard locks only if needed.
export const accountMutationLockId = 1_604_619_418;
export const adminUserSelect = {
  createdAt: true,
  email: true,
  enabled: true,
  id: true,
  mustChangePassword: true,
  name: true,
  role: true,
  updatedAt: true,
} as const;
type AdminUser = Prisma.UserGetPayload<{ select: typeof adminUserSelect }>;

export function normalizedAccountEmail(value: string): string {
  const email = normalizeEmail(value);
  if (
    email.length > accountEmailMaximumLength ||
    !accountEmailPattern.test(email)
  ) {
    fail(400, "invalid_request", "email must be a valid email address");
  }
  return email;
}

export function accountUserSummary(user: AdminUser): AdminUser {
  return {
    createdAt: user.createdAt,
    email: user.email,
    enabled: user.enabled,
    id: user.id,
    mustChangePassword: user.mustChangePassword,
    name: user.name,
    role: user.role,
    updatedAt: user.updatedAt,
  };
}

export function generateTemporaryPassword(): string {
  return randomBytes(24).toString("base64url");
}

async function lockAccountMutationActor(
  tx: Prisma.TransactionClient,
  identity: Identity
): Promise<void> {
  const [actor] = await tx.$queryRaw<{ id: string }[]>(
    Prisma.sql`
      SELECT "account_actor"."id"
      FROM "user" AS "account_actor"
      INNER JOIN "session" AS "account_session"
        ON "account_session"."user_id" = "account_actor"."id"
      WHERE
        "account_actor"."id" = ${identity.id}
        AND "account_actor"."role" = CAST('admin' AS "UserRole")
        AND "account_actor"."enabled" = TRUE
        AND "account_actor"."must_change_password" = FALSE
        AND "account_session"."id" = ${identity.sessionId}
        AND "account_session"."expires_at" > NOW()
      FOR UPDATE OF "account_actor", "account_session"
    `
  );
  if (!actor) {
    fail(403, "forbidden", "Administrator authorization changed");
  }
}

export function accountTransaction<T>(
  identity: Identity,
  operation: (tx: Prisma.TransactionClient) => Promise<T>
): Promise<T> {
  return prisma.$transaction(async (tx) => {
    await tx.$executeRaw`SELECT pg_advisory_xact_lock(${accountMutationLockId})`;
    await lockAccountMutationActor(tx, identity);
    return operation(tx);
  });
}

export async function lockAccountUser(
  tx: Prisma.TransactionClient,
  userId: string
): Promise<AdminUser> {
  const [locked] = await tx.$queryRaw<{ id: string }[]>(
    Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${userId} FOR UPDATE`
  );
  if (!locked) {
    fail(404, "not_found", "User was not found");
  }
  const user = await tx.user.findUnique({
    select: adminUserSelect,
    where: { id: userId },
  });
  if (!user) {
    fail(404, "not_found", "User was not found");
  }
  return user;
}

export async function withAdminMutation<T>(
  request: Request,
  action: AccountAuditAction,
  targetId: string | null | (() => Promise<string | null>),
  operation: (
    identity: Identity,
    setAction: (action: AccountAuditAction) => void
  ) => Promise<T>
): Promise<T> {
  const identity = await identityFor(request);
  if (!identity) {
    fail(401, "unauthorized", "Authentication is required");
  }
  requireAdmin(identity);
  let auditAction = action;
  let auditTargetId = typeof targetId === "function" ? null : targetId;
  try {
    if (identity.mustChangePassword) {
      fail(403, "password_change_required", "Password replacement is required");
    }
    if (typeof targetId === "function") {
      auditTargetId = await targetId();
    }
    return await operation(identity, (nextAction) => {
      auditAction = nextAction;
    });
  } catch (error) {
    const normalizedError =
      databaseErrorCode(error) === "P2002"
        ? new HttpError(409, "email_in_use", "Email is already in use")
        : error;
    try {
      await createAccountFailureAudit({
        action: auditAction,
        actorId: identity.id,
        error: normalizedError,
        targetId: auditTargetId,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw normalizedError;
  }
}
