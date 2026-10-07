// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { HandoffStatus, ResponseStatus, prisma, Prisma } from "@onlyoffice/db";

import { tokenDigest } from "../digests";
import { handoffUnavailable } from "./input";

type ExternalPrefillStatus =
  | "deleted"
  | "draft"
  | "expired"
  | "pending"
  | "submitted";

function externalPrefillStatus(handoff: {
  consumedAt: Date | null;
  createdAt: Date;
  expiresAt: Date;
  reservedAt: Date | null;
  status: HandoffStatus;
  updatedAt: Date;
  response?: {
    corrections: { revision: number }[];
    status: ResponseStatus;
    submission: { createdAt: Date } | null;
  } | null;
}): {
  consumedAt: string | null;
  createdAt: string;
  expiresAt: string;
  latestCorrectionNumber: number | null;
  reservedAt: string | null;
  status: ExternalPrefillStatus;
  submittedAt: string | null;
  updatedAt: string;
} {
  const status =
    handoff.status === HandoffStatus.deleted
      ? "deleted"
      : handoff.status === HandoffStatus.expired
        ? "expired"
        : handoff.response?.status === ResponseStatus.submitted
          ? "submitted"
          : handoff.response
            ? "draft"
            : "pending";
  return {
    consumedAt: handoff.consumedAt?.toISOString() ?? null,
    createdAt: handoff.createdAt.toISOString(),
    expiresAt: handoff.expiresAt.toISOString(),
    latestCorrectionNumber: handoff.response?.corrections[0]?.revision ?? null,
    reservedAt: handoff.reservedAt?.toISOString() ?? null,
    status,
    submittedAt: handoff.response?.submission?.createdAt.toISOString() ?? null,
    updatedAt: handoff.updatedAt.toISOString(),
  };
}
function deletedExternalPrefillStatus(
  createdAt: Date
): ReturnType<typeof externalPrefillStatus> & { deletedAt: string } {
  const timestamp = createdAt.toISOString();
  return {
    consumedAt: null,
    createdAt: timestamp,
    deletedAt: timestamp,
    expiresAt: timestamp,
    latestCorrectionNumber: null,
    reservedAt: null,
    status: "deleted",
    submittedAt: null,
    updatedAt: timestamp,
  };
}

export async function pollPrefillHandoffStatus(
  externalReference: string,
  clock: () => Date = () => new Date()
): Promise<ReturnType<typeof externalPrefillStatus>> {
  const relations = {
    response: {
      select: {
        corrections: {
          orderBy: { revision: "desc" as const },
          select: { revision: true },
          take: 1,
        },
        status: true,
        submission: { select: { createdAt: true } },
      },
    },
  } as const;
  const externalReferenceDigest = tokenDigest(externalReference);
  const handoffResult = await prisma.$transaction(async (tx) => {
    const deletionTombstone = await tx.deletionTombstone.findUnique({
      where: { externalReferenceDigest },
    });
    if (deletionTombstone) {
      const pendingCleanup = await tx.objectCleanupIntent.count({
        where: {
          deletionResponseLookupDigest: deletionTombstone.responseLookupDigest,
        },
      });
      return pendingCleanup > 0
        ? { cleanupPending: true }
        : { deletedAt: deletionTombstone.createdAt };
    }
    const current = await tx.handoff.findFirst({
      include: relations,
      orderBy: { createdAt: "desc" },
      where: { externalReferenceDigest },
    });
    if (!current) {
      handoffUnavailable();
    }
    if (
      current.status === HandoffStatus.deleted &&
      current.deletionResponseLookupDigest
    ) {
      const pendingCleanup = await tx.objectCleanupIntent.count({
        where: {
          deletionResponseLookupDigest: current.deletionResponseLookupDigest,
        },
      });
      if (pendingCleanup > 0) {
        return { cleanupPending: true };
      }
    }
    const now = clock();
    if (
      (current.status === HandoffStatus.pending ||
        current.status === HandoffStatus.reserved) &&
      current.expiresAt <= now
    ) {
      await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "pending_claims"
          WHERE "handoff_id" = ${current.id}::uuid
          FOR UPDATE
        `
      );
      const expired = await tx.handoff.updateMany({
        data: {
          codeDigest: null,
          configurationHash: null,
          filteredValues: Prisma.JsonNull,
          formId: null,
          normalizedEmail: null,
          responseId: null,
          status: HandoffStatus.expired,
        },
        where: {
          expiresAt: { lte: now },
          id: current.id,
          status: { in: [HandoffStatus.pending, HandoffStatus.reserved] },
        },
      });
      if (expired.count === 1) {
        await tx.pendingClaim.deleteMany({ where: { handoffId: current.id } });
      }
      const refreshed = await tx.handoff.findUnique({
        include: relations,
        where: { id: current.id },
      });
      if (!refreshed) {
        handoffUnavailable();
      }
      return { handoff: refreshed };
    }
    return { handoff: current };
  });
  if (handoffResult.cleanupPending) {
    handoffUnavailable();
  }
  if (handoffResult.deletedAt instanceof Date) {
    return deletedExternalPrefillStatus(handoffResult.deletedAt);
  }
  if (!handoffResult.handoff) {
    handoffUnavailable();
  }
  if (handoffResult.handoff.status === HandoffStatus.deleted) {
    return deletedExternalPrefillStatus(handoffResult.handoff.updatedAt);
  }
  return externalPrefillStatus(handoffResult.handoff);
}
