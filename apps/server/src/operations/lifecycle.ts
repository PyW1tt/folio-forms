import {
  OperationStatus,
  prisma,
  OperationTargetType,
  AuditOutcome,
  ResponseStatus,
} from "@onlyoffice/db";
import type { Prisma } from "@onlyoffice/db";

import { createFormAudit, formAuditErrorCodeFromCode } from "../audit/events";
import { tokenDigest } from "../digests";
import { fail } from "../http/errors";
import { publicIdPattern } from "../http/input";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Operation, JsonRecord } from "../model-types";
import { jsonValue } from "../responses/data";
import { objectExists } from "../storage";
import type { OperationErrorCode, OperationMetadata } from "./model";
import { operationMetadata } from "./model";
import { deleteObjects, deleteObjectUnlessCanonical } from "./object-cleanup";

export const operationTimeoutMs = 4 * 60_000;
export const callbackClaimLifetimeSeconds = 5 * 60;

async function activeAfterRecovery(
  operation: Operation | null
): Promise<boolean> {
  if (!operation) {
    return false;
  }
  const current = await expireOperationIfNeeded(operation);
  return (
    current.status === OperationStatus.pending ||
    current.status === OperationStatus.processing
  );
}

export async function activeOperationForForm(formId: string): Promise<boolean> {
  return activeAfterRecovery(
    await prisma.operation.findFirst({
      where: {
        formId,
        status: { in: [OperationStatus.pending, OperationStatus.processing] },
        targetType: OperationTargetType.template_draft,
      },
    })
  );
}

export async function activeOperationForResponse(
  responseId: string
): Promise<boolean> {
  return activeAfterRecovery(
    await prisma.operation.findFirst({
      where: {
        responseId,
        status: { in: [OperationStatus.pending, OperationStatus.processing] },
      },
    })
  );
}

export async function updateOperationFailed(
  operationId: string,
  errorCode: OperationErrorCode,
  updatedBefore?: Date
): Promise<boolean> {
  const operation = await prisma.$transaction(async (tx) => {
    const current = await tx.operation.findUnique({
      select: { actorId: true, metadata: true, stagingObjectKey: true },
      where: { id: operationId },
    });
    if (!current) {
      return null;
    }
    const metadata = operationMetadata(current.metadata);
    const failed = await tx.operation.updateMany({
      data: {
        errorCode,
        status: OperationStatus.failed,
        updatedAt: new Date(),
      },
      where: {
        id: operationId,
        ...(updatedBefore ? { updatedAt: { lte: updatedBefore } } : {}),
        status: { in: [OperationStatus.pending, OperationStatus.processing] },
      },
    });
    if (failed.count !== 1) {
      return null;
    }
    if (metadata.action === "save-template" || metadata.action === "publish") {
      let targetId =
        typeof metadata.publicId === "string" &&
        publicIdPattern.test(metadata.publicId)
          ? metadata.publicId
          : null;
      if (!targetId) {
        const targetForm = await tx.form.findUnique({
          select: { publicId: true },
          where: { id: metadata.formId },
        });
        targetId = targetForm?.publicId ?? null;
      }
      await createFormAudit(tx, {
        action:
          metadata.action === "publish"
            ? "publish_form"
            : "save_template_draft",
        actorId: current.actorId,
        outcome: AuditOutcome.failure,
        safeMetadata: { errorCode: formAuditErrorCodeFromCode(errorCode) },
        targetId,
      });
    }
    if (
      metadata.action === "save-correction" &&
      metadata.responseId &&
      current.actorId
    ) {
      await tx.auditEvent.create({
        data: {
          action: "create_correction",
          actorId: current.actorId,
          outcome: AuditOutcome.failure,
          safeMetadata: jsonValue({ errorCode, state: "submitted" }),
          targetId: metadata.responseId,
          targetType: "response",
        },
      });
    }
    if (metadata.action === "submit" && metadata.responseId) {
      await tx.response.updateMany({
        data: { status: ResponseStatus.draft, updatedAt: new Date() },
        where: {
          id: metadata.responseId,
          status: ResponseStatus.submitting,
        },
      });
    }
    return { ...current, metadata };
  });
  if (!operation) {
    return false;
  }
  await deleteObjects([operation.stagingObjectKey]);
  await deleteObjectUnlessCanonical(operation.metadata.finalObjectKey);
  return true;
}

export async function expireOperationIfNeeded(
  operation: Operation
): Promise<Operation> {
  const active =
    operation.status === OperationStatus.pending ||
    operation.status === OperationStatus.processing;
  const staleBefore = new Date(Date.now() - operationTimeoutMs);
  const stale = operation.updatedAt <= staleBefore;
  if (!active || !stale) {
    return operation;
  }
  await updateOperationFailed(operation.id, "operation_timeout", staleBefore);
  return (
    (await prisma.operation.findUnique({ where: { id: operation.id } })) ??
    operation
  );
}

type CallbackClaimConsumption = "claimed" | "invalid" | "replayed";

export async function consumeCallbackClaim(
  operationId: string,
  userdata: string
): Promise<CallbackClaimConsumption> {
  const digest = tokenDigest(userdata);
  const consumed = await prisma.callbackClaim.updateMany({
    data: { consumedAt: new Date() },
    where: {
      consumedAt: null,
      expiresAt: { gt: new Date() },
      operationId,
      tokenDigest: digest,
    },
  });
  if (consumed.count === 1) {
    return "claimed";
  }
  const existing = await prisma.callbackClaim.findFirst({
    select: { consumedAt: true },
    where: { operationId, tokenDigest: digest },
  });
  return existing?.consumedAt ? "replayed" : "invalid";
}

export async function operationDocumentKey(
  documentKey: string
): Promise<string | null> {
  const pending = await prisma.operation.findMany({
    orderBy: { updatedAt: "desc" },
    select: { metadata: true },
    take: 10,
    where: {
      documentKey,
      status: { in: [OperationStatus.pending, OperationStatus.processing] },
    },
  });
  for (const operation of pending) {
    const metadata = operationMetadata(operation.metadata);
    if (await objectExists(metadata.stagedObjectKey)) {
      return metadata.stagedObjectKey;
    }
  }

  const templateDraft = await prisma.templateDraft.findUnique({
    where: { documentKey },
  });
  if (templateDraft) {
    return templateDraft.objectKey;
  }
  const publishedTemplate = await prisma.publishedTemplate.findUnique({
    where: { documentKey },
  });
  if (publishedTemplate) {
    return publishedTemplate.objectKey;
  }
  const response = await prisma.response.findUnique({
    select: { draftObjectKey: true },
    where: { draftDocumentKey: documentKey },
  });
  if (response?.draftObjectKey) {
    return response.draftObjectKey;
  }
  const correctionWorkspace = await prisma.editorLease.findUnique({
    select: { workspaceObjectKey: true },
    where: { workspaceDocumentKey: documentKey },
  });
  if (correctionWorkspace?.workspaceObjectKey) {
    return correctionWorkspace.workspaceObjectKey;
  }
  const correction = await prisma.correction.findUnique({
    select: { objectKey: true },
    where: { documentKey },
  });
  if (correction) {
    return correction.objectKey;
  }
  const submission = await prisma.submission.findUnique({
    select: { objectKey: true },
    where: { documentKey },
  });
  return submission?.objectKey ?? null;
}

export async function markOperationCompleted(
  tx: Prisma.TransactionClient,
  operationId: string,
  result: JsonRecord,
  metadata: OperationMetadata,
  cleanupObjectKeys: string[],
  submissionId?: string
): Promise<void> {
  const completed = await tx.operation.updateMany({
    data: {
      errorCode: null,
      metadata: jsonValue({ ...metadata, cleanupObjectKeys }),
      result: jsonValue(result),
      status: OperationStatus.completed,
      ...(submissionId ? { submissionId } : {}),
      updatedAt: new Date(),
    },
    where: { id: operationId, status: OperationStatus.processing },
  });
  if (completed.count !== 1) {
    fail(409, "stale_operation", "The document operation is no longer active");
  }
}
