// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { prisma, Prisma, HandoffStatus } from "@onlyoffice/db";

const handoffExpirySweepBatchSize = 100;
export async function expireDuePrefillHandoffs(now: Date): Promise<void> {
  await prisma.$transaction(async (tx) => {
    const dueHandoffs = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`
        SELECT "id"
        FROM "handoffs"
        WHERE "status" IN ('pending', 'reserved')
          AND "expires_at" <= ${now}
        ORDER BY "expires_at" ASC
        LIMIT ${handoffExpirySweepBatchSize}
      `
    );
    if (dueHandoffs.length === 0) {
      return;
    }
    for (const handoff of dueHandoffs) {
      await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "pending_claims"
          WHERE "handoff_id" = ${handoff.id}::uuid
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
          id: handoff.id,
          status: { in: [HandoffStatus.pending, HandoffStatus.reserved] },
        },
      });
      if (expired.count === 1) {
        await tx.pendingClaim.deleteMany({
          where: { handoffId: handoff.id },
        });
      }
    }
  });
}
