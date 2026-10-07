import { Prisma, LegacyAccountLinkStatus, AuditOutcome } from "@onlyoffice/db";

import { createAccountAudit } from "../audit/events";
import type { Identity } from "../auth/identity";
import { HttpError, fail } from "../http/errors";
import { readJsonRecord } from "../http/input";
import { nextLegacySsoGeneration } from "../legacy-sso/accounts";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { JsonRecord } from "../model-types";
import { accountTransaction, lockAccountUser } from "./mutations";

const legacyReviewedGenerationPattern = /^[1-9][0-9]{0,18}$/u;

export async function readLegacyAccountLinkReviewGeneration(
  request: Request
): Promise<bigint | null> {
  let input: JsonRecord;
  try {
    input = await readJsonRecord(request, 1024);
  } catch (error) {
    if (error instanceof HttpError && error.httpStatus === 413) {
      fail(400, "invalid_request", "Account link review body is too large");
    }
    throw error;
  }
  if (
    Object.keys(input).length !== 1 ||
    !Object.hasOwn(input, "reviewedGeneration")
  ) {
    fail(400, "invalid_request", "Only reviewedGeneration is accepted");
  }
  if (input.reviewedGeneration === null) {
    return null;
  }
  if (
    typeof input.reviewedGeneration !== "string" ||
    !legacyReviewedGenerationPattern.test(input.reviewedGeneration)
  ) {
    fail(
      400,
      "invalid_request",
      "reviewedGeneration must be a positive bigint decimal string or null"
    );
  }
  const generation = BigInt(input.reviewedGeneration);
  if (generation > 9_223_372_036_854_775_807n) {
    fail(400, "invalid_request", "reviewedGeneration is out of range");
  }
  return generation;
}

export async function reviewLegacyAccountLink(
  identity: Identity,
  requestId: string,
  decision: "approved" | "rejected",
  expectedReviewedGeneration: bigint | null
): Promise<{ ok: true }> {
  const action =
    decision === "approved"
      ? "approve_legacy_account_link"
      : "reject_legacy_account_link";
  await accountTransaction(identity, async (tx) => {
    const requestSnapshot = await tx.legacyAccountLinkRequest.findUnique({
      select: { userId: true },
      where: { id: requestId },
    });
    if (!requestSnapshot) {
      fail(404, "not_found", "Account link request was not found");
    }
    const user = await lockAccountUser(tx, requestSnapshot.userId);
    const [lockedRequest] = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`SELECT "id" FROM "legacy_account_link_requests" WHERE "id" = ${requestId}::uuid FOR UPDATE`
    );
    if (!lockedRequest) {
      fail(404, "not_found", "Account link request was not found");
    }
    const linkRequest = await tx.legacyAccountLinkRequest.findUnique({
      where: { id: requestId },
    });
    if (!linkRequest) {
      fail(404, "not_found", "Account link request was not found");
    }
    if (linkRequest.status !== LegacyAccountLinkStatus.pending) {
      fail(
        409,
        "account_link_not_pending",
        "Account link request is not pending"
      );
    }
    if (
      linkRequest.reviewedGeneration !== expectedReviewedGeneration ||
      linkRequest.userId !== user.id
    ) {
      fail(
        409,
        "account_link_changed",
        "Account link request changed; reload and review it again"
      );
    }
    if (decision === "approved") {
      if (
        !user.enabled ||
        user.role !== "user" ||
        user.email !== linkRequest.email
      ) {
        fail(
          409,
          "account_link_not_eligible",
          "Candidate account is not eligible for linking"
        );
      }
      const account = await tx.account.findUnique({
        select: { userId: true },
        where: {
          providerId_accountId: {
            accountId: linkRequest.subject,
            providerId: linkRequest.providerId,
          },
        },
      });
      if (account) {
        fail(
          409,
          "identity_already_linked",
          "Legacy identity is already linked"
        );
      }
    }
    const [databaseTime] = await tx.$queryRaw<{ currentTime: Date }[]>(
      Prisma.sql`SELECT CURRENT_TIMESTAMP AS "currentTime"`
    );
    if (!databaseTime) {
      throw new Error("Database did not return current time");
    }
    const generation =
      decision === "approved"
        ? await nextLegacySsoGeneration(tx, linkRequest.providerId)
        : linkRequest.reviewedGeneration;
    const updated = await tx.legacyAccountLinkRequest.updateMany({
      data: {
        reviewedAt: databaseTime.currentTime,
        reviewedById: identity.id,
        reviewedGeneration: generation,
        status:
          decision === "approved"
            ? LegacyAccountLinkStatus.approved
            : LegacyAccountLinkStatus.rejected,
        updatedAt: databaseTime.currentTime,
      },
      where: { id: requestId, status: LegacyAccountLinkStatus.pending },
    });
    if (updated.count !== 1) {
      fail(
        409,
        "account_link_not_pending",
        "Account link request is not pending"
      );
    }
    await createAccountAudit(tx, {
      action,
      actorId: identity.id,
      outcome: AuditOutcome.success,
      safeMetadata: { change: decision },
      targetId: user.id,
    });
  });
  return { ok: true };
}
