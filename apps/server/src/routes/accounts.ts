import { auth } from "@onlyoffice/auth";
import {
  prisma,
  Prisma,
  LegacyAccountLinkStatus,
  AuditOutcome,
  HandoffStatus,
} from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import {
  adminUserSelect,
  accountUserPageSize,
  accountUserSummary,
  withAdminMutation,
  accountNameMaximumLength,
  normalizedAccountEmail,
  generateTemporaryPassword,
  accountTransaction,
  lockAccountUser,
} from "../accounts/mutations";
import {
  reviewLegacyAccountLink,
  readLegacyAccountLinkReviewGeneration,
} from "../accounts/review";
import { accountAuditTargetId, createAccountAudit } from "../audit/events";
import type { AccountAuditAction } from "../audit/events";
import {
  requireIdentity,
  requireAdmin,
  normalizeEmail,
} from "../auth/identity";
import type { UserRole } from "../auth/identity";
import { endAiAuthoringSessions } from "../auth/password-session";
import { fail } from "../http/errors";
import {
  validateId,
  queryString,
  idPattern,
  accountEmailMaximumLength,
  readJsonRecord,
  accountBodyMaximumBytes,
  requiredString,
} from "../http/input";
import { invalidateLegacyAccountLinkReview } from "../legacy-sso/accounts";
import type { JsonRecord } from "../model-types";
import type { RouteDependencies } from "./dependencies";

export function registerAccountRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "aiAuthoring">
): void {
  const { aiAuthoring } = dependencies;
  app
    .get("/api/admin/account-links", async ({ request, query }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const queryRecord = query as unknown as JsonRecord;
      const cursor = queryString(queryRecord, "cursor");
      if (cursor !== undefined && !idPattern.test(cursor)) {
        fail(400, "invalid_request", "cursor is invalid");
      }
      const where: Prisma.LegacyAccountLinkRequestWhereInput = {
        status: LegacyAccountLinkStatus.pending,
      };
      if (cursor) {
        where.id = { gt: cursor };
      }
      const rows = await prisma.legacyAccountLinkRequest.findMany({
        orderBy: { id: "asc" },
        select: {
          createdAt: true,
          email: true,
          id: true,
          providerId: true,
          reviewedGeneration: true,
          status: true,
          subject: true,
          user: { select: adminUserSelect },
        },
        take: accountUserPageSize + 1,
        where,
      });
      const page = rows.slice(0, accountUserPageSize);
      return {
        nextCursor:
          rows.length > accountUserPageSize ? (page.at(-1)?.id ?? null) : null,
        requests: page.map(({ user, ...linkRequest }) => ({
          ...linkRequest,
          reviewedGeneration:
            linkRequest.reviewedGeneration?.toString() ?? null,
          user: accountUserSummary(user),
        })),
      };
    })
    .post(
      "/api/admin/account-links/:id/approve",
      ({ request, params }) => {
        const requestId = accountAuditTargetId(params.id);
        if (!requestId) {
          fail(400, "invalid_request", "Account link request id is invalid");
        }
        return withAdminMutation(
          request,
          "approve_legacy_account_link",
          async () => {
            const linkRequest =
              await prisma.legacyAccountLinkRequest.findUnique({
                select: { userId: true },
                where: { id: requestId },
              });
            if (!linkRequest) {
              fail(404, "not_found", "Account link request was not found");
            }
            return linkRequest.userId;
          },
          async (identity) =>
            reviewLegacyAccountLink(
              identity,
              requestId,
              "approved",
              await readLegacyAccountLinkReviewGeneration(request)
            )
        );
      },
      { parse: "none" }
    )
    .post(
      "/api/admin/account-links/:id/reject",
      ({ request, params }) => {
        const requestId = accountAuditTargetId(params.id);
        if (!requestId) {
          fail(400, "invalid_request", "Account link request id is invalid");
        }
        return withAdminMutation(
          request,
          "reject_legacy_account_link",
          async () => {
            const linkRequest =
              await prisma.legacyAccountLinkRequest.findUnique({
                select: { userId: true },
                where: { id: requestId },
              });
            if (!linkRequest) {
              fail(404, "not_found", "Account link request was not found");
            }
            return linkRequest.userId;
          },
          async (identity) =>
            reviewLegacyAccountLink(
              identity,
              requestId,
              "rejected",
              await readLegacyAccountLinkReviewGeneration(request)
            )
        );
      },
      { parse: "none" }
    )
    .get("/api/admin/users", async ({ request, query }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const queryRecord = query as unknown as JsonRecord;
      const cursor = queryString(queryRecord, "cursor");
      const emailQuery = queryString(queryRecord, "email");
      const roleQuery = queryString(queryRecord, "role");
      const enabledQuery = queryString(queryRecord, "enabled");
      if (cursor !== undefined && !idPattern.test(cursor)) {
        fail(400, "invalid_request", "cursor is invalid");
      }
      const email = emailQuery ? normalizeEmail(emailQuery) : undefined;
      if (email && email.length > accountEmailMaximumLength) {
        fail(400, "invalid_request", "email filter is too long");
      }
      let role: UserRole | undefined;
      if (roleQuery !== undefined) {
        if (roleQuery !== "admin" && roleQuery !== "user") {
          fail(400, "invalid_request", "role must be admin or user");
        }
        role = roleQuery;
      }
      let enabled: boolean | undefined;
      if (enabledQuery !== undefined) {
        if (enabledQuery !== "true" && enabledQuery !== "false") {
          fail(400, "invalid_request", "enabled must be true or false");
        }
        enabled = enabledQuery === "true";
      }
      const where: Prisma.UserWhereInput = {};
      if (cursor) {
        where.id = { gt: cursor };
      }
      if (email) {
        where.email = { contains: email };
      }
      if (role) {
        where.role = role;
      }
      if (enabled !== undefined) {
        where.enabled = enabled;
      }
      const users = await prisma.user.findMany({
        orderBy: { id: "asc" },
        select: adminUserSelect,
        take: accountUserPageSize + 1,
        where,
      });
      const page = users.slice(0, accountUserPageSize);
      return {
        nextCursor:
          users.length > accountUserPageSize ? (page.at(-1)?.id ?? null) : null,
        users: page.map(accountUserSummary),
      };
    })
    .post(
      "/api/admin/users",
      ({ request }) =>
        withAdminMutation(request, "create_user", null, async (identity) => {
          const input = await readJsonRecord(request, accountBodyMaximumBytes);
          const keys = Object.keys(input);
          if (
            keys.length !== 3 ||
            keys.some(
              (key) => key !== "name" && key !== "email" && key !== "role"
            )
          ) {
            fail(400, "invalid_request", "name, email, and role are required");
          }
          const name = requiredString(input, "name");
          if (name.length > accountNameMaximumLength) {
            fail(400, "invalid_request", "name is too long");
          }
          const email = normalizedAccountEmail(requiredString(input, "email"));
          const roleValue = requiredString(input, "role");
          if (roleValue !== "admin" && roleValue !== "user") {
            fail(400, "invalid_request", "role must be admin or user");
          }
          const temporaryPassword = generateTemporaryPassword();
          const user = await accountTransaction(identity, async (tx) => {
            const existing = await tx.user.findUnique({
              select: { id: true },
              where: { email },
            });
            if (existing) {
              fail(409, "email_in_use", "Email is already in use");
            }
            const authContext = await auth.$context;
            const passwordHash =
              await authContext.password.hash(temporaryPassword);
            const userId = crypto.randomUUID();
            const created = await tx.user.create({
              data: {
                accounts: {
                  create: {
                    accountId: userId,
                    id: crypto.randomUUID(),
                    issuer: "local:credential",
                    password: passwordHash,
                    providerId: "credential",
                  },
                },
                email,
                emailVerified: true,
                enabled: true,
                id: userId,
                mustChangePassword: true,
                name,
                role: roleValue,
              },
              select: adminUserSelect,
            });
            await createAccountAudit(tx, {
              action: "create_user",
              actorId: identity.id,
              outcome: AuditOutcome.success,
              safeMetadata: { change: "created" },
              targetId: created.id,
            });
            return created;
          });
          return {
            temporaryPassword,
            user: accountUserSummary(user),
          };
        }),
      { parse: "none" }
    )
    .patch(
      "/api/admin/users/:id",
      ({ request, params }) =>
        withAdminMutation(
          request,
          "update_user",
          accountAuditTargetId(params.id),
          async (identity, setAction) => {
            const userId = validateId(params.id, "User");
            const input = await readJsonRecord(
              request,
              accountBodyMaximumBytes
            );
            const keys = Object.keys(input);
            const key = keys[0];
            if (
              keys.length !== 1 ||
              (key !== "enabled" && key !== "email" && key !== "role")
            ) {
              fail(
                400,
                "invalid_request",
                "Exactly one of enabled, email, or role is required"
              );
            }
            let action: AccountAuditAction = "update_user";
            let change = "updated";
            let updateData: Prisma.UserUpdateInput;
            let newEmail: string | undefined;
            let revokeSessions = false;
            if (key === "enabled") {
              if (typeof input.enabled !== "boolean") {
                fail(400, "invalid_request", "enabled must be a boolean");
              }
              action = input.enabled ? "enable_user" : "disable_user";
              change = input.enabled ? "enabled" : "disabled";
              revokeSessions = !input.enabled;
              updateData = { enabled: input.enabled };
            } else if (key === "email") {
              if (typeof input.email !== "string") {
                fail(400, "invalid_request", "email is required");
              }
              newEmail = normalizedAccountEmail(input.email);
              action = "change_user_email";
              change = "email_changed";
              revokeSessions = true;
              updateData = { email: newEmail };
            } else {
              if (input.role !== "admin" && input.role !== "user") {
                fail(400, "invalid_request", "role must be admin or user");
              }
              action = input.role === "admin" ? "promote_user" : "demote_user";
              change = input.role === "admin" ? "promoted" : "demoted";
              revokeSessions = true;
              updateData = { role: input.role };
            }
            setAction(action);
            const result = await accountTransaction(identity, async (tx) => {
              const target = await lockAccountUser(tx, userId);
              const removesFinalAdmin =
                target.role === "admin" &&
                target.enabled &&
                (updateData.enabled === false || updateData.role === "user");
              if (removesFinalAdmin) {
                const enabledAdminCount = await tx.user.count({
                  where: { enabled: true, role: "admin" },
                });
                if (enabledAdminCount <= 1) {
                  fail(
                    409,
                    "final_admin_required",
                    "At least one enabled Admin is required"
                  );
                }
              }
              if (newEmail !== undefined) {
                const existing = await tx.user.findUnique({
                  select: { id: true },
                  where: { email: newEmail },
                });
                if (existing && existing.id !== target.id) {
                  fail(409, "email_in_use", "Email is already in use");
                }
              }
              if (newEmail !== undefined && newEmail !== target.email) {
                const staleHandoffs = await tx.handoff.findMany({
                  select: { id: true },
                  where: {
                    normalizedEmail: target.email,
                    responseId: null,
                  },
                });
                if (staleHandoffs.length > 0) {
                  const staleHandoffIds = staleHandoffs.map(
                    (handoff) => handoff.id
                  );
                  await tx.pendingClaim.deleteMany({
                    where: { handoffId: { in: staleHandoffIds } },
                  });
                  await tx.handoff.deleteMany({
                    where: { id: { in: staleHandoffIds } },
                  });
                }
                const incompleteRequests =
                  await tx.legacyAccountLinkRequest.findMany({
                    orderBy: { id: "asc" },
                    select: {
                      email: true,
                      id: true,
                      providerId: true,
                      userId: true,
                    },
                    where: {
                      status: {
                        in: [
                          LegacyAccountLinkStatus.pending,
                          LegacyAccountLinkStatus.approved,
                        ],
                      },
                      userId: target.id,
                    },
                  });
                for (const linkRequest of incompleteRequests) {
                  await tx.$queryRaw(
                    Prisma.sql`SELECT "id" FROM "legacy_account_link_requests" WHERE "id" = ${linkRequest.id}::uuid FOR UPDATE`
                  );
                }
                for (const linkRequest of incompleteRequests) {
                  await invalidateLegacyAccountLinkReview(
                    tx,
                    linkRequest.id,
                    linkRequest.providerId,
                    linkRequest.email,
                    linkRequest.userId
                  );
                }
              }
              const updated = await tx.user.update({
                data: updateData,
                select: adminUserSelect,
                where: { id: target.id },
              });
              const ownerSessions = revokeSessions
                ? await tx.session.findMany({
                    select: { id: true },
                    where: { userId: target.id },
                  })
                : [];
              const ownerSessionIds = ownerSessions.map(({ id }) => id);
              if (revokeSessions) {
                await tx.session.deleteMany({ where: { userId: target.id } });
              }
              await createAccountAudit(tx, {
                action,
                actorId: identity.id,
                outcome: AuditOutcome.success,
                safeMetadata: { change },
                targetId: target.id,
              });
              return { ownerSessionIds, user: updated };
            });
            await endAiAuthoringSessions(aiAuthoring, result.ownerSessionIds);
            return { user: accountUserSummary(result.user) };
          }
        ),
      { parse: "none" }
    )
    .delete(
      "/api/admin/users/:id",
      ({ request, params }) =>
        withAdminMutation(
          request,
          "delete_user",
          accountAuditTargetId(params.id),
          async (identity) => {
            const userId = validateId(params.id, "User");
            const input = await readJsonRecord(
              request,
              accountBodyMaximumBytes
            );
            if (Object.keys(input).length !== 1 || input.confirm !== true) {
              fail(400, "invalid_request", "confirm must be true");
            }
            const revokedOwnerSessionIds = await accountTransaction(
              identity,
              async (tx) => {
                const target = await lockAccountUser(tx, userId);
                if (target.role === "admin" && target.enabled) {
                  const enabledAdminCount = await tx.user.count({
                    where: { enabled: true, role: "admin" },
                  });
                  if (enabledAdminCount <= 1) {
                    fail(
                      409,
                      "final_admin_required",
                      "At least one enabled Admin is required"
                    );
                  }
                }
                const responseCount = await tx.response.count({
                  where: { userId: target.id },
                });
                if (responseCount > 0) {
                  fail(
                    409,
                    "personal_data_remains",
                    "Personal Responses must be deleted first"
                  );
                }
                const cleanupIntentCount = await tx.objectCleanupIntent.count({
                  where: { deletionOwnerUserId: target.id },
                });
                if (cleanupIntentCount > 0) {
                  fail(
                    409,
                    "personal_data_remains",
                    "Response object cleanup is still pending"
                  );
                }
                const handoffs = await tx.handoff.findMany({
                  select: { id: true },
                  where: { normalizedEmail: target.email },
                });
                if (handoffs.length > 0) {
                  await tx.pendingClaim.deleteMany({
                    where: {
                      handoffId: {
                        in: handoffs.map((handoff) => handoff.id),
                      },
                    },
                  });
                  await tx.handoff.updateMany({
                    data: {
                      codeDigest: null,
                      configurationHash: null,
                      consumedAt: null,
                      filteredValues: Prisma.JsonNull,
                      formId: null,
                      normalizedEmail: null,
                      reservedAt: null,
                      responseId: null,
                      status: HandoffStatus.deleted,
                    },
                    where: {
                      id: { in: handoffs.map((handoff) => handoff.id) },
                    },
                  });
                }
                await createAccountAudit(tx, {
                  action: "delete_user",
                  actorId: identity.id,
                  outcome: AuditOutcome.success,
                  safeMetadata: { change: "deleted" },
                  targetId: target.id,
                });
                const sessions = await tx.session.findMany({
                  select: { id: true },
                  where: { userId: target.id },
                });
                await tx.user.delete({ where: { id: target.id } });
                return sessions.map(({ id }) => id);
              }
            );
            await endAiAuthoringSessions(aiAuthoring, revokedOwnerSessionIds);
            return { deleted: true };
          }
        ),
      { parse: "none" }
    )
    .post(
      "/api/admin/users/:id/password-reset",
      ({ request, params }) =>
        withAdminMutation(
          request,
          "reset_user_password",
          accountAuditTargetId(params.id),
          async (identity) => {
            const userId = validateId(params.id, "User");
            const temporaryPassword = generateTemporaryPassword();
            const result = await accountTransaction(identity, async (tx) => {
              const target = await lockAccountUser(tx, userId);
              const account = await tx.account.findFirst({
                where: { providerId: "credential", userId: target.id },
              });
              if (!account?.password) {
                fail(500, "credential_missing", "Credential is unavailable");
              }
              const authContext = await auth.$context;
              const passwordHash =
                await authContext.password.hash(temporaryPassword);
              await tx.account.update({
                data: { password: passwordHash },
                where: { id: account.id },
              });
              const updated = await tx.user.update({
                data: { mustChangePassword: true },
                select: adminUserSelect,
                where: { id: target.id },
              });
              const ownerSessions = await tx.session.findMany({
                select: { id: true },
                where: { userId: target.id },
              });
              const ownerSessionIds = ownerSessions.map(({ id }) => id);
              await tx.session.deleteMany({ where: { userId: target.id } });
              await createAccountAudit(tx, {
                action: "reset_user_password",
                actorId: identity.id,
                outcome: AuditOutcome.success,
                safeMetadata: { change: "password_reset" },
                targetId: target.id,
              });
              return { ownerSessionIds, user: updated };
            });
            await endAiAuthoringSessions(aiAuthoring, result.ownerSessionIds);
            return {
              temporaryPassword,
              user: accountUserSummary(result.user),
            };
          }
        ),
      { parse: "none" }
    );
}
