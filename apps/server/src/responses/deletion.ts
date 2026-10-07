import {
  prisma,
  Prisma,
  ResponseStatus,
  OperationStatus,
  OperationTargetType,
  HandoffStatus,
  AuditOutcome,
} from "@onlyoffice/db";

import type { AiAuthoringSessions } from "../ai-authoring";
import { createResponseDeletionAudit } from "../audit/events";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Identity } from "../auth/identity";
import { endAiAuthoringSessions } from "../auth/password-session";
import { tokenDigest } from "../digests";
import { fail } from "../http/errors";
import { operationMetadata } from "../operations/model";
import {
  drainObjectCleanupIntents,
  uniqueObjectKeys,
} from "../operations/object-cleanup";

interface ResponseDeletionResult {
  alreadyDeleted: boolean;
  deleted: boolean;
  ownerSessionIds: string[];
}

function responseDeletionPrefixes(
  responseId: string,
  submissionId: string | null,
  operationPrefixes: readonly string[]
): string[] {
  return [
    `responses/${responseId}/`,
    ...(submissionId ? [`submissions/${submissionId}/`] : []),
    ...operationPrefixes,
  ];
}

function responseDeletionKeyBelongs(
  key: string,
  prefixes: readonly string[]
): boolean {
  return prefixes.some((prefix) => key.startsWith(prefix));
}

export async function deleteResponseData({
  actor,
  aiAuthoring,
  allowActiveLease,
  missingOk,
  removeObject,
  responseId,
  revokeOwnerSessions,
}: {
  actor: Identity;
  aiAuthoring: AiAuthoringSessions;
  allowActiveLease: boolean;
  missingOk: boolean;
  removeObject: (key: string) => Promise<void>;
  responseId: string;
  revokeOwnerSessions: boolean;
}): Promise<ResponseDeletionResult> {
  const responseLookupDigest = tokenDigest(responseId);
  const existingTombstone = await prisma.deletionTombstone.findUnique({
    where: { responseLookupDigest },
  });
  if (existingTombstone) {
    await drainObjectCleanupIntents(
      undefined,
      responseLookupDigest,
      removeObject
    );
    const remaining = await prisma.objectCleanupIntent.count({
      where: { deletionResponseLookupDigest: responseLookupDigest },
    });
    if (remaining > 0) {
      fail(503, "deletion_cleanup_failed", "Response objects remain");
    }
    return { alreadyDeleted: true, deleted: true, ownerSessionIds: [] };
  }
  const result = await prisma.$transaction(
    async (tx) => {
      const [lockedResponse] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "responses"
          WHERE "id" = ${responseId}::uuid
          FOR UPDATE
        `
      );
      if (!lockedResponse) {
        if (missingOk) {
          return {
            alreadyDeleted: false,
            deleted: false,
            ownerSessionIds: [],
          };
        }
        fail(404, "not_found", "Response was not found");
      }
      const response = await tx.response.findUnique({
        include: {
          corrections: { orderBy: { revision: "asc" } },
          submission: true,
        },
        where: { id: responseId },
      });
      if (!response) {
        if (missingOk) {
          return {
            alreadyDeleted: false,
            deleted: false,
            ownerSessionIds: [],
          };
        }
        fail(404, "not_found", "Response was not found");
      }
      if (actor.role !== "admin" && response.userId !== actor.id) {
        fail(403, "forbidden", "You may only delete your own Draft");
      }
      if (actor.role !== "admin" && response.status !== ResponseStatus.draft) {
        fail(409, "draft_unavailable", "Only a Draft can be discarded");
      }
      const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "forms"
          WHERE "id" = ${response.formId}::uuid
          FOR UPDATE
        `
      );
      if (!lockedForm) {
        fail(409, "stale_response", "The Response form is unavailable");
      }
      const operations = await tx.operation.findMany({
        select: {
          id: true,
          metadata: true,
          stagingObjectKey: true,
          status: true,
        },
        where: {
          OR: [
            { responseId: response.id },
            ...(response.submission
              ? [{ submissionId: response.submission.id }]
              : []),
            ...(response.corrections.length > 0
              ? [
                  {
                    correctionId: {
                      in: response.corrections.map(
                        (correction) => correction.id
                      ),
                    },
                  },
                ]
              : []),
          ],
        },
      });
      if (
        operations.some(
          (operation) =>
            operation.status === OperationStatus.pending ||
            operation.status === OperationStatus.processing
        )
      ) {
        fail(409, "operation_in_progress", "The Response operation is active");
      }
      const correctionIds = response.corrections.map(
        (correction) => correction.id
      );
      const leases = await tx.editorLease.findMany({
        select: { id: true, workspaceObjectKey: true },
        where: {
          OR: [
            {
              targetId: response.id,
              targetType: OperationTargetType.response,
            },
            {
              targetId: response.id,
              targetType: OperationTargetType.correction,
            },
          ],
        },
      });
      if (leases.length > 0 && !allowActiveLease) {
        fail(409, "editor_in_use", "The Response is open in an editor");
      }

      const candidateKeys = [
        response.draftObjectKey,
        response.submission?.objectKey,
        ...response.corrections.map((correction) => correction.objectKey),
        ...leases.map((lease) => lease.workspaceObjectKey),
        ...operations.flatMap((operation) => {
          const metadata = operationMetadata(operation.metadata);
          return [
            operation.stagingObjectKey,
            metadata.finalObjectKey,
            metadata.stagedObjectKey,
            ...(metadata.cleanupObjectKeys ?? []),
            metadata.workspaceObjectKey,
          ];
        }),
      ];
      const linkedObjectPrefixes = [
        ...new Set(
          uniqueObjectKeys(candidateKeys).flatMap((key) => {
            const [scope, pathId] = key.split("/", 3);
            return (scope === "operations" || scope === "submissions") && pathId
              ? [`${scope}/${pathId}/`]
              : [];
          })
        ),
      ];
      const prefixes = responseDeletionPrefixes(
        response.id,
        response.submission?.id ?? null,
        linkedObjectPrefixes
      );
      for (const key of uniqueObjectKeys(candidateKeys)) {
        if (!responseDeletionKeyBelongs(key, prefixes)) {
          fail(
            500,
            "invalid_object_key",
            "Response object ownership is invalid"
          );
        }
      }
      const cleanupIntents = await tx.objectCleanupIntent.findMany({
        select: { objectKey: true },
        where: {
          OR: prefixes.map((prefix) => ({
            objectKey: { startsWith: prefix },
          })),
        },
      });
      const objectKeys = uniqueObjectKeys([
        ...candidateKeys,
        ...cleanupIntents.map((intent) => intent.objectKey),
      ]);
      for (const objectKeyValue of objectKeys) {
        await tx.objectCleanupIntent.upsert({
          create: {
            deletionOwnerUserId: response.userId,
            deletionResponseLookupDigest: responseLookupDigest,
            objectKey: objectKeyValue,
          },
          update: {
            cleanupAfter: new Date(),
            deletionOwnerUserId: response.userId,
            deletionResponseLookupDigest: responseLookupDigest,
          },
          where: { objectKey: objectKeyValue },
        });
      }

      const handoffs = await tx.handoff.findMany({
        select: { id: true },
        where: {
          OR: [
            { responseId: response.id },
            ...(response.externalReferenceDigest
              ? [{ externalReferenceDigest: response.externalReferenceDigest }]
              : []),
          ],
        },
      });
      if (handoffs.length > 0) {
        const handoffIds = handoffs.map((handoff) => handoff.id);
        await tx.pendingClaim.deleteMany({
          where: { handoffId: { in: handoffIds } },
        });
        await tx.handoff.updateMany({
          data: {
            codeDigest: null,
            configurationHash: null,
            consumedAt: null,
            deletionResponseLookupDigest: responseLookupDigest,
            filteredValues: Prisma.JsonNull,
            formId: null,
            normalizedEmail: null,
            reservedAt: null,
            responseId: null,
            status: HandoffStatus.deleted,
          },
          where: { id: { in: handoffIds } },
        });
      }
      await tx.editorLease.deleteMany({
        where: {
          OR: [
            {
              targetId: response.id,
              targetType: OperationTargetType.response,
            },
            {
              targetId: response.id,
              targetType: OperationTargetType.correction,
            },
          ],
        },
      });
      await tx.operation.deleteMany({
        where: {
          OR: [
            { responseId: response.id },
            ...(response.submission
              ? [{ submissionId: response.submission.id }]
              : []),
            ...(correctionIds.length > 0
              ? [{ correctionId: { in: correctionIds } }]
              : []),
          ],
        },
      });
      const ownerSessions = revokeOwnerSessions
        ? await tx.session.findMany({
            select: { id: true },
            where: { userId: response.userId },
          })
        : [];
      const ownerSessionIds = ownerSessions.map(({ id }) => id);
      if (revokeOwnerSessions) {
        await tx.session.deleteMany({ where: { userId: response.userId } });
      }
      await tx.correction.deleteMany({ where: { responseId: response.id } });
      await tx.submission.deleteMany({ where: { responseId: response.id } });
      await tx.prefillSnapshot.deleteMany({
        where: { responseId: response.id },
      });
      await tx.response.delete({ where: { id: response.id } });
      await tx.deletionTombstone.create({
        data: {
          actorId: actor.id,
          externalReferenceDigest: response.externalReferenceDigest,
          id: crypto.randomUUID(),
          outcome: AuditOutcome.success,
          responseLookupDigest,
        },
      });
      await createResponseDeletionAudit({
        action: "delete_response_pending",
        actorId: actor.id,
        outcome: AuditOutcome.success,
        targetId: response.id,
        tx,
      });
      return { alreadyDeleted: false, deleted: true, ownerSessionIds };
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  if (!result.deleted) {
    return result;
  }
  const cleanupResults = await Promise.allSettled([
    endAiAuthoringSessions(aiAuthoring, result.ownerSessionIds),
    (async () => {
      await drainObjectCleanupIntents(
        undefined,
        responseLookupDigest,
        removeObject
      );
      const remaining = await prisma.objectCleanupIntent.count({
        where: { deletionResponseLookupDigest: responseLookupDigest },
      });
      if (remaining > 0) {
        fail(503, "deletion_cleanup_failed", "Response objects remain");
      }
    })(),
  ]);
  const cleanupErrors = cleanupResults.flatMap((cleanup) =>
    cleanup.status === "rejected" ? [cleanup.reason] : []
  );
  if (cleanupErrors.length === 1) {
    throw cleanupErrors[0];
  }
  if (cleanupErrors.length > 1) {
    throw new AggregateError(cleanupErrors, "Response deletion cleanup failed");
  }
  await createResponseDeletionAudit({
    actorId: actor.id,
    outcome: AuditOutcome.success,
    targetId: responseId,
  });
  return result;
}
