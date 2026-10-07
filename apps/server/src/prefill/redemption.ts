import {
  prisma,
  ResponseStatus,
  Prisma,
  HandoffStatus,
  OperationStatus,
  PrefillPolicy,
  OperationTargetType,
  AuditOutcome,
} from "@onlyoffice/db";

import { createFormAudit, createFormFailureAudit } from "../audit/events";
import type { Identity } from "../auth/identity";
import { normalizeEmail } from "../auth/identity";
import { databaseErrorCode, isSerializationConflict } from "../db-errors";
import { tokenDigest } from "../digests";
import { fail, HttpError } from "../http/errors";
import { jsonRecord, publicIdPattern } from "../http/input";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type {
  ResponseWithSnapshot,
  FormWithDocuments,
  JsonRecord,
} from "../model-types";
import {
  uniqueObjectKeys,
  deleteObjectUnlessCanonical,
} from "../operations/object-cleanup";
import { jsonValue } from "../responses/data";
import {
  objectExists,
  readObject,
  objectKey,
  putObject,
  DOCX_CONTENT_TYPE,
} from "../storage";
import { handoffUnavailable } from "./input";
import {
  publishedPrefillConfiguration,
  externalValueMatchesSchema,
} from "./values";

interface PrefillHandoffRedeemResult {
  cleanupObjectKeys: string[];
  response: ResponseWithSnapshot;
}

export async function redeemPrefillHandoff(
  form: FormWithDocuments,
  identity: Identity,
  claimToken: string,
  clock: () => Date = () => new Date()
): Promise<PrefillHandoffRedeemResult> {
  const responseHint = await prisma.response.findUnique({
    select: { id: true, status: true },
    where: { formId_userId: { formId: form.id, userId: identity.id } },
  });
  if (responseHint?.status === ResponseStatus.submitted) {
    handoffUnavailable();
  }
  const responseId = responseHint?.id ?? crypto.randomUUID();
  const publishedTemplate = form.publishedTemplate;
  if (!publishedTemplate?.objectKey || !publishedTemplate.documentKey) {
    handoffUnavailable();
  }
  if (!(await objectExists(publishedTemplate.objectKey))) {
    handoffUnavailable();
  }
  const document = await readObject(publishedTemplate.objectKey);
  const draftObjectKey = objectKey(
    "responses",
    responseId,
    "draft",
    crypto.randomUUID(),
    "docx"
  );
  const draftDocumentKey = `response-${responseId}-${crypto.randomUUID()}`;
  await putObject(draftObjectKey, document, DOCX_CONTENT_TYPE);
  const claimDigest = tokenDigest(claimToken);
  let oldDraftObjectKey: string | null = null;
  let unusedDraftObjectKey: string | null = null;
  try {
    const response = await prisma.$transaction(
      async (tx) => {
        const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "forms"
            WHERE "id" = ${form.id}::uuid
            FOR UPDATE
          `
        );
        if (!lockedForm) {
          handoffUnavailable();
        }
        const currentForm = await tx.form.findUnique({
          include: {
            publishedTemplate: {
              include: {
                prefillConfiguration: { include: { fields: true } },
              },
            },
          },
          where: { id: lockedForm.id },
        });
        const configuration = publishedPrefillConfiguration(currentForm);
        if (
          !currentForm ||
          currentForm.publicId !== form.publicId ||
          !configuration ||
          !currentForm.publishedTemplate
        ) {
          handoffUnavailable();
        }
        const [lockedClaim] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "pending_claims"
            WHERE "claim_digest" = ${claimDigest}
            FOR UPDATE
          `
        );
        if (!lockedClaim) {
          handoffUnavailable();
        }
        const pendingClaim = await tx.pendingClaim.findUnique({
          include: { handoff: true },
          where: { id: lockedClaim.id },
        });
        const handoff = pendingClaim?.handoff;
        const now = clock();
        if (
          !pendingClaim ||
          !handoff ||
          pendingClaim.consumedAt ||
          pendingClaim.expiresAt <= now ||
          handoff.formId !== currentForm.id ||
          handoff.status !== HandoffStatus.reserved ||
          handoff.consumedAt ||
          handoff.expiresAt <= now ||
          handoff.normalizedEmail !== normalizeEmail(identity.email) ||
          handoff.configurationHash !== configuration.configurationHash ||
          configuration.publishedTemplateId !==
            currentForm.publishedTemplate.id ||
          configuration.configurationHash !==
            currentForm.publishedTemplate.contentHash
        ) {
          handoffUnavailable();
        }
        const deletedReference = await tx.deletionTombstone.findUnique({
          select: { id: true },
          where: {
            externalReferenceDigest: handoff.externalReferenceDigest,
          },
        });
        if (deletedReference) {
          handoffUnavailable();
        }
        const currentResponse = await tx.response.findUnique({
          include: { prefillSnapshot: true, submission: true },
          where: {
            formId_userId: { formId: currentForm.id, userId: identity.id },
          },
        });
        if (
          currentResponse?.status === ResponseStatus.submitted ||
          currentResponse?.status === ResponseStatus.submitting ||
          (currentResponse &&
            (await tx.operation.findFirst({
              select: { id: true },
              where: {
                responseId: currentResponse.id,
                status: {
                  in: [OperationStatus.pending, OperationStatus.processing],
                },
              },
            })))
        ) {
          handoffUnavailable();
        }
        const reuseExistingDraft =
          currentResponse?.status === ResponseStatus.draft &&
          currentResponse.publishedVersion === currentForm.version &&
          currentResponse.draftDocumentKey !== null &&
          currentResponse.draftObjectKey !== null;
        const storedValues = jsonRecord(
          handoff.filteredValues,
          "The prefill handoff values are invalid"
        );
        const values: JsonRecord = {};
        const lockedFields: JsonRecord = {};
        for (const field of configuration.fields) {
          const value = storedValues[field.tag];
          if (value === undefined) {
            continue;
          }
          if (!externalValueMatchesSchema(field.pointer, value)) {
            handoffUnavailable();
          }
          values[field.tag] = value;
          lockedFields[field.tag] =
            field.policy === PrefillPolicy.lock_when_available;
        }
        const responseTargetId = currentResponse?.id ?? responseId;
        if (reuseExistingDraft) {
          unusedDraftObjectKey = draftObjectKey;
        } else {
          oldDraftObjectKey = currentResponse?.draftObjectKey ?? null;
          if (currentResponse) {
            await tx.editorLease.deleteMany({
              where: {
                targetId: currentResponse.id,
                targetType: OperationTargetType.response,
              },
            });
            await tx.operation.deleteMany({
              where: { responseId: currentResponse.id },
            });
            await tx.prefillSnapshot.deleteMany({
              where: { responseId: currentResponse.id },
            });
            await tx.response.update({
              data: {
                draftData: Prisma.DbNull,
                draftDocumentKey,
                draftObjectKey,
                externalReferenceDigest: handoff.externalReferenceDigest,
                publishedTemplateId: currentForm.publishedTemplate.id,
                publishedVersion: currentForm.version,
                status: ResponseStatus.draft,
                updatedAt: now,
              },
              where: { id: responseTargetId },
            });
          } else {
            await tx.response.create({
              data: {
                draftData: Prisma.DbNull,
                draftDocumentKey,
                draftObjectKey,
                externalReferenceDigest: handoff.externalReferenceDigest,
                form: { connect: { id: currentForm.id } },
                id: responseTargetId,
                owner: { connect: { id: identity.id } },
                publishedTemplate: {
                  connect: { id: currentForm.publishedTemplate.id },
                },
                publishedVersion: currentForm.version,
                status: ResponseStatus.draft,
              },
            });
          }
          await tx.prefillSnapshot.create({
            data: {
              form: { connect: { id: currentForm.id } },
              id: crypto.randomUUID(),
              lockedFields: jsonValue(lockedFields),
              owner: { connect: { id: identity.id } },
              response: { connect: { id: responseTargetId } },
              values: jsonValue(values),
            },
          });
        }
        const consumedClaim = await tx.pendingClaim.updateMany({
          data: { consumedAt: now },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: pendingClaim.id,
          },
        });
        if (consumedClaim.count !== 1) {
          handoffUnavailable();
        }
        const consumedHandoff = await tx.handoff.updateMany({
          data: {
            consumedAt: now,
            responseId: responseTargetId,
            status: HandoffStatus.consumed,
          },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: handoff.id,
            status: HandoffStatus.reserved,
          },
        });
        if (consumedHandoff.count !== 1) {
          handoffUnavailable();
        }
        await createFormAudit(tx, {
          action: "redeem_handoff",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: currentForm.publicId,
        });
        const created = reuseExistingDraft
          ? currentResponse
          : await tx.response.findUnique({
              include: { prefillSnapshot: true },
              where: { id: responseTargetId },
            });
        if (!created) {
          fail(500, "start_failed", "Unable to redeem prefill handoff");
        }
        return created;
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
    return {
      cleanupObjectKeys: uniqueObjectKeys([
        oldDraftObjectKey,
        unusedDraftObjectKey,
      ]),
      response,
    };
  } catch (error) {
    await deleteObjectUnlessCanonical(draftObjectKey);
    const normalizedError =
      databaseErrorCode(error) === "P2002" || isSerializationConflict(error)
        ? new HttpError(
            409,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          )
        : error;
    try {
      await createFormFailureAudit({
        action: "redeem_handoff",
        actorId: identity.id,
        error: normalizedError,
        targetId: publicIdPattern.test(form.publicId) ? form.publicId : null,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw normalizedError;
  }
}
export async function consumeSubmittedPrefillHandoff(
  form: FormWithDocuments,
  identity: Identity,
  claimToken: string,
  responseId: string,
  clock: () => Date = () => new Date()
): Promise<void> {
  const claimDigest = tokenDigest(claimToken);
  try {
    await prisma.$transaction(
      async (tx) => {
        const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
          SELECT "id"
          FROM "forms"
          WHERE "id" = ${form.id}::uuid
          FOR UPDATE
        `
        );
        if (!lockedForm) {
          handoffUnavailable();
        }
        const currentForm = await tx.form.findUnique({
          include: {
            publishedTemplate: {
              include: { prefillConfiguration: { include: { fields: true } } },
            },
          },
          where: { id: lockedForm.id },
        });
        const configuration = publishedPrefillConfiguration(currentForm);
        const currentTemplate = currentForm?.publishedTemplate;
        const [lockedClaim] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
          SELECT "id"
          FROM "pending_claims"
          WHERE "claim_digest" = ${claimDigest}
          FOR UPDATE
        `
        );
        if (!lockedClaim) {
          handoffUnavailable();
        }
        const pendingClaim = await tx.pendingClaim.findUnique({
          include: { handoff: true },
          where: { id: lockedClaim.id },
        });
        const handoff = pendingClaim?.handoff;
        const currentResponse = await tx.response.findUnique({
          select: { id: true, status: true },
          where: { id: responseId },
        });
        const now = clock();
        if (
          !currentForm ||
          currentForm.publicId !== form.publicId ||
          !currentTemplate ||
          !configuration ||
          !currentResponse ||
          currentResponse.status !== ResponseStatus.submitted ||
          currentResponse.id !== responseId ||
          !pendingClaim ||
          !handoff ||
          pendingClaim.consumedAt ||
          pendingClaim.expiresAt <= now ||
          handoff.formId !== currentForm.id ||
          handoff.responseId ||
          handoff.status !== HandoffStatus.reserved ||
          handoff.consumedAt ||
          handoff.expiresAt <= now ||
          handoff.normalizedEmail !== normalizeEmail(identity.email) ||
          handoff.configurationHash !== configuration.configurationHash ||
          configuration.publishedTemplateId !== currentTemplate.id ||
          configuration.configurationHash !== currentTemplate.contentHash
        ) {
          handoffUnavailable();
        }
        const consumedClaim = await tx.pendingClaim.updateMany({
          data: { consumedAt: now },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: pendingClaim.id,
          },
        });
        if (consumedClaim.count !== 1) {
          handoffUnavailable();
        }
        const consumedHandoff = await tx.handoff.updateMany({
          data: {
            consumedAt: now,
            responseId,
            status: HandoffStatus.consumed,
          },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: handoff.id,
            responseId: null,
            status: HandoffStatus.reserved,
          },
        });
        if (consumedHandoff.count !== 1) {
          handoffUnavailable();
        }
        await createFormAudit(tx, {
          action: "redeem_handoff",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: currentForm.publicId,
        });
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
  } catch (error) {
    const normalizedError =
      databaseErrorCode(error) === "P2002" || isSerializationConflict(error)
        ? new HttpError(
            409,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          )
        : error;
    try {
      await createFormFailureAudit({
        action: "redeem_handoff",
        actorId: identity.id,
        error: normalizedError,
        targetId: publicIdPattern.test(form.publicId) ? form.publicId : null,
      });
    } catch {
      // Preserve the retryable response if the failure audit cannot be written.
    }
    throw normalizedError;
  }
}
