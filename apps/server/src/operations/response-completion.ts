import {
  prisma,
  ResponseStatus,
  Prisma,
  OperationStatus,
  OperationTargetType,
  AuditOutcome,
} from "@onlyoffice/db";

import { validateResponseDocument } from "../documents/native-eligibility";
import { fail, HttpError } from "../http/errors";
import { jsonRecord } from "../http/input";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Operation } from "../model-types";
import { jsonValue, changedResponseData } from "../responses/data";
import { nativeResponseDocument } from "../responses/native-document";
import type { putObject } from "../storage";
import { readObject, DOCX_CONTENT_TYPE } from "../storage";
import { markOperationCompleted, updateOperationFailed } from "./lifecycle";
import type { OperationMetadata, OperationCompletion } from "./model";
import { operationMetadata } from "./model";
import { deleteObjects, deleteObjectUnlessCanonical } from "./object-cleanup";

export async function completeDraftOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  if (!documentKey) {
    fail(500, "invalid_operation", "Draft operation has no document key");
  }
  if (!metadata.responseId || !metadata.data) {
    fail(500, "invalid_operation", "Draft metadata is incomplete");
  }
  const response = await prisma.response.findUnique({
    where: { id: metadata.responseId },
  });
  if (
    !response ||
    response.status !== ResponseStatus.draft ||
    response.draftDocumentKey !== documentKey
  ) {
    fail(409, "stale_operation", "The response is no longer editable");
  }
  await validateResponseDocument(response.publishedTemplateId, bytes, false);
  const nextDocumentKey = metadata.nextDocumentKey ?? documentKey;
  const result = {
    documentKey: nextDocumentKey,
    publicId: metadata.publicId,
    responseId: response.id,
  };
  const cleanupObjectKeys = [
    metadata.stagedObjectKey,
    ...(response.draftObjectKey ? [response.draftObjectKey] : []),
  ];
  await prisma.$transaction(
    async (tx) => {
      const updated = await tx.response.updateMany({
        data: {
          draftData: jsonValue(metadata.data),
          draftDocumentKey: nextDocumentKey,
          draftObjectKey: metadata.finalObjectKey,
          updatedAt: new Date(),
        },
        where: {
          draftDocumentKey: documentKey,
          id: response.id,
          status: ResponseStatus.draft,
        },
      });
      if (updated.count !== 1) {
        fail(409, "stale_operation", "The response is no longer editable");
      }
      await markOperationCompleted(
        tx,
        operation.id,
        result,
        metadata,
        cleanupObjectKeys
      );
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}

export async function completeSubmitOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  if (!documentKey) {
    fail(500, "invalid_operation", "Submit operation has no document key");
  }
  if (!metadata.responseId || !metadata.submissionId || !metadata.data) {
    fail(500, "invalid_operation", "Submit metadata is incomplete");
  }
  const { responseId, submissionId, data } = metadata;
  const response = await prisma.response.findUnique({
    where: { id: responseId },
  });
  if (
    !response ||
    response.status !== ResponseStatus.submitting ||
    response.draftDocumentKey !== documentKey
  ) {
    fail(
      409,
      "stale_operation",
      "The response is no longer pending submission"
    );
  }
  await validateResponseDocument(response.publishedTemplateId, bytes, true);

  const submissionDocumentKey = metadata.submissionDocumentKey ?? documentKey;
  const result = {
    publicId: metadata.publicId,
    responseId: response.id,
    submissionId,
  };
  const cleanupObjectKeys = [
    metadata.stagedObjectKey,
    ...(response.draftObjectKey ? [response.draftObjectKey] : []),
  ];
  await prisma.$transaction(
    async (tx) => {
      const claimed = await tx.operation.updateMany({
        data: { updatedAt: new Date() },
        where: { id: operation.id, status: OperationStatus.processing },
      });
      if (claimed.count !== 1) {
        fail(
          409,
          "stale_operation",
          "The submission operation is no longer active"
        );
      }
      await tx.submission.create({
        data: {
          data: jsonValue(data),
          documentKey: submissionDocumentKey,
          form: { connect: { id: response.formId } },
          id: submissionId,
          objectKey: metadata.finalObjectKey,
          owner: { connect: { id: response.userId } },
          response: { connect: { id: response.id } },
        },
      });
      const updatedResponse = await tx.response.updateMany({
        data: {
          draftData: Prisma.DbNull,
          draftDocumentKey: null,
          draftObjectKey: null,
          status: ResponseStatus.submitted,
          updatedAt: new Date(),
        },
        where: { id: response.id, status: ResponseStatus.submitting },
      });
      if (updatedResponse.count !== 1) {
        fail(
          409,
          "stale_operation",
          "The response is no longer pending submission"
        );
      }
      await markOperationCompleted(
        tx,
        operation.id,
        result,
        metadata,
        cleanupObjectKeys,
        submissionId
      );
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}
export async function processNativeResponseOperation(
  operation: Operation,
  storeObject: typeof putObject
): Promise<void> {
  const metadata = operationMetadata(operation.metadata);
  try {
    if (
      (metadata.action !== "save-draft" && metadata.action !== "submit") ||
      !operation.documentKey ||
      !metadata.responseId ||
      !metadata.data ||
      !metadata.finalObjectKey ||
      !metadata.stagedObjectKey
    ) {
      fail(500, "invalid_operation", "Native response operation is incomplete");
    }
    const claimed = await prisma.operation.updateMany({
      data: { status: OperationStatus.processing, updatedAt: new Date() },
      where: { id: operation.id, status: OperationStatus.pending },
    });
    if (claimed.count !== 1) {
      return;
    }
    const response = await prisma.response.findUnique({
      select: {
        draftDocumentKey: true,
        publishedTemplateId: true,
        status: true,
      },
      where: { id: metadata.responseId },
    });
    const expectedStatus =
      metadata.action === "submit"
        ? ResponseStatus.submitting
        : ResponseStatus.draft;
    if (
      !response ||
      response.status !== expectedStatus ||
      response.draftDocumentKey !== operation.documentKey
    ) {
      fail(409, "stale_operation", "The response is no longer editable");
    }
    const bytes = metadata.nativeDocumentStaged
      ? await readObject(metadata.stagedObjectKey)
      : await nativeResponseDocument(
          response.publishedTemplateId,
          metadata.data
        );
    if (!metadata.nativeDocumentStaged) {
      await storeObject(metadata.stagedObjectKey, bytes, DOCX_CONTENT_TYPE);
    }
    await storeObject(metadata.finalObjectKey, bytes, DOCX_CONTENT_TYPE);
    const completion =
      metadata.action === "submit"
        ? await completeSubmitOperation(operation, metadata, bytes)
        : await completeDraftOperation(operation, metadata, bytes);
    await deleteObjects(completion.cleanupObjectKeys);
  } catch (error) {
    await deleteObjects([metadata.stagedObjectKey]);
    if (metadata.finalObjectKey) {
      await deleteObjectUnlessCanonical(metadata.finalObjectKey);
    }
    await updateOperationFailed(
      operation.id,
      error instanceof HttpError && error.code === "invalid_template"
        ? "invalid_template"
        : "document_save_failed"
    );
  }
}
export async function completeCorrectionOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  const {
    baseDocumentKey,
    baseRevision,
    data,
    nextDocumentKey,
    reason,
    responseId,
    submissionId,
  } = metadata;
  const actorId = operation.actorId;
  if (
    !documentKey ||
    !responseId ||
    !submissionId ||
    !data ||
    typeof baseDocumentKey !== "string" ||
    typeof baseRevision !== "number" ||
    !nextDocumentKey ||
    !reason ||
    !actorId
  ) {
    fail(500, "invalid_operation", "Correction metadata is incomplete");
  }
  const response = await prisma.response.findUnique({
    include: { submission: true },
    where: { id: responseId },
  });
  const submission = response?.submission;
  if (
    !response ||
    response.status !== ResponseStatus.submitted ||
    !submission ||
    submission.id !== submissionId
  ) {
    fail(409, "stale_operation", "The Submission is no longer correctable");
  }
  await validateResponseDocument(response.publishedTemplateId, bytes, true);
  const cleanupObjectKeys = [
    metadata.stagedObjectKey,
    metadata.workspaceObjectKey,
  ].filter((key): key is string => Boolean(key));
  await prisma.$transaction(
    async (tx) => {
      const [lockedResponse] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "responses"
          WHERE "id" = ${response.id}::uuid
          FOR UPDATE
        `
      );
      if (!lockedResponse) {
        fail(409, "stale_operation", "The Submission is no longer correctable");
      }
      const current = await tx.response.findUnique({
        include: {
          corrections: {
            orderBy: { revision: "desc" },
            take: 1,
          },
          submission: true,
        },
        where: { id: response.id },
      });
      const currentSubmission = current?.submission;
      const latest = current?.corrections[0];
      const currentRevision = latest?.revision ?? 0;
      const currentDocumentKey =
        latest?.documentKey ?? currentSubmission?.documentKey;
      if (
        !current ||
        current.status !== ResponseStatus.submitted ||
        !currentSubmission ||
        currentSubmission.id !== submissionId ||
        currentRevision !== baseRevision ||
        currentDocumentKey !== baseDocumentKey
      ) {
        fail(409, "stale_operation", "A newer Correction is already effective");
      }
      const previousData = jsonRecord(latest?.data ?? currentSubmission.data);
      const correction = await tx.correction.create({
        data: {
          actorId,
          changedData: jsonValue(changedResponseData(previousData, data)),
          data: jsonValue(data),
          documentKey: nextDocumentKey,
          objectKey: metadata.finalObjectKey,
          reason,
          responseId: current.id,
          revision: currentRevision + 1,
          submissionId,
        },
      });
      await tx.response.update({
        data: { updatedAt: new Date() },
        where: { id: current.id },
      });
      await tx.editorLease.updateMany({
        data: {
          workspaceBaseDocumentKey: null,
          workspaceBaseRevision: null,
          workspaceDocumentKey: null,
          workspaceObjectKey: null,
        },
        where: {
          targetId: current.id,
          targetType: OperationTargetType.correction,
          workspaceDocumentKey: metadata.workspaceDocumentKey,
        },
      });
      const completedResult = {
        correctionId: correction.id,
        publicId: metadata.publicId,
        responseId: current.id,
        revision: correction.revision,
        submissionId,
      };
      await markOperationCompleted(
        tx,
        operation.id,
        completedResult,
        metadata,
        cleanupObjectKeys
      );
      await tx.auditEvent.create({
        data: {
          action: "create_correction",
          actorId,
          outcome: AuditOutcome.success,
          safeMetadata: jsonValue({
            revision: correction.revision,
            state: "submitted",
          }),
          targetId: correction.id,
          targetType: "correction",
        },
      });
      return completedResult;
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}
