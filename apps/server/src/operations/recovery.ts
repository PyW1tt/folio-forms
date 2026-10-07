import { prisma, OperationStatus, Prisma } from "@onlyoffice/db";

import { expireDuePrefillHandoffs } from "../prefill/expiry";
import { operationTimeoutMs, updateOperationFailed } from "./lifecycle";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import {
  drainObjectCleanupIntents,
  deleteObjectUnlessCanonical,
} from "./object-cleanup";

export async function reconcileRecoverableState(): Promise<void> {
  await drainObjectCleanupIntents();
  const now = new Date();
  await expireDuePrefillHandoffs(now);
  const staleBefore = new Date(now.getTime() - operationTimeoutMs);
  const staleOperations = await prisma.operation.findMany({
    select: { id: true },
    where: {
      status: { in: [OperationStatus.pending, OperationStatus.processing] },
      updatedAt: { lte: staleBefore },
    },
  });
  for (const operation of staleOperations) {
    await updateOperationFailed(operation.id, "operation_timeout", staleBefore);
  }
  const expiredLeases = await prisma.editorLease.findMany({
    select: {
      id: true,
      targetId: true,
      targetType: true,
    },
    where: {
      OR: [
        { expiresAt: { lte: now } },
        { holderSession: { expiresAt: { lte: now } } },
      ],
    },
  });
  for (const lease of expiredLeases) {
    const workspaceObjectKey = await prisma.$transaction(async (tx) => {
      const [locked] = await tx.$queryRaw<
        {
          expiresAt: Date;
          sessionExpiresAt: Date;
          workspaceObjectKey: string | null;
        }[]
      >(
        Prisma.sql`
          SELECT
            "editor_leases"."expires_at" AS "expiresAt",
            "session"."expires_at" AS "sessionExpiresAt",
            "editor_leases"."workspace_object_key" AS "workspaceObjectKey"
          FROM "editor_leases"
          INNER JOIN "session"
            ON "session"."id" = "editor_leases"."holder_session_id"
          WHERE "editor_leases"."id" = ${lease.id}::uuid
          FOR UPDATE
        `
      );
      if (
        !locked ||
        (locked.expiresAt > now && locked.sessionExpiresAt > now)
      ) {
        return null;
      }
      const activeOperation = await tx.operation.findFirst({
        select: { id: true },
        where: {
          status: { in: [OperationStatus.pending, OperationStatus.processing] },
          targetId: lease.targetId,
          targetType: lease.targetType,
        },
      });
      if (activeOperation) {
        return null;
      }
      const deleted = await tx.editorLease.deleteMany({
        where: { id: lease.id },
      });
      return deleted.count === 1 ? locked.workspaceObjectKey : null;
    });
    if (workspaceObjectKey) {
      await deleteObjectUnlessCanonical(workspaceObjectKey);
    }
  }
  await prisma.callbackClaim.deleteMany({
    where: { expiresAt: { lte: now } },
  });
  await drainObjectCleanupIntents();
}
