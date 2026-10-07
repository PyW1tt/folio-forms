import {
  prisma,
  ResponseStatus,
  OperationTargetType,
  AuditOutcome,
  FillMethod,
} from "@onlyoffice/db";

import { createResponseAudit } from "../audit/events";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Identity } from "../auth/identity";
import { onlyOfficeFieldTagAliases } from "../documents/office-values";
import { actionEditorCapability } from "../editor/authorization";
import type {
  EditorLeaseGrant,
  CorrectionWorkspaceInput,
} from "../editor/leases";
import {
  renewEditorLease,
  editorLeaseProof,
  releaseEditorLease,
  claimEditorLease,
  editorLeaseBridge,
} from "../editor/leases";
import { nativeFields } from "../forms/query";
import { fail } from "../http/errors";
import { validateId, jsonRecord } from "../http/input";
import type { FormWithDocuments } from "../model-types";
import { editorConfig } from "../onlyoffice";
import { operationTimeoutMs } from "../operations/lifecycle";
import { deleteObjectUnlessCanonical } from "../operations/object-cleanup";
import {
  lockedPrefillForSnapshot,
  editableFieldsForSnapshot,
} from "../prefill/values";
import {
  objectExists,
  objectKey,
  putObject,
  readObject,
  DOCX_CONTENT_TYPE,
} from "../storage";

export async function correctionEditorConfig(
  responseId: string,
  identity: Identity
): Promise<Record<string, unknown>> {
  validateId(responseId, "Response");
  const response = await prisma.response.findUnique({
    include: {
      corrections: {
        orderBy: { revision: "desc" },
        take: 1,
      },
      prefillSnapshot: true,
      submission: true,
    },
    where: { id: responseId },
  });
  if (
    !response ||
    response.status !== ResponseStatus.submitted ||
    !response.submission
  ) {
    fail(404, "not_found", "Submitted Response was not found");
  }
  const form = await prisma.form.findUnique({
    include: { publishedTemplate: true, templateDraft: true },
    where: { id: response.formId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  const latest = response.corrections[0];
  const baseRevision = latest?.revision ?? 0;
  const baseDocumentKey =
    latest?.documentKey ?? response.submission.documentKey;
  const sourceObjectKey = latest?.objectKey ?? response.submission.objectKey;
  if (!(await objectExists(sourceObjectKey))) {
    fail(409, "document_unavailable", "Response document is unavailable");
  }

  const now = new Date();
  const priorLease = await prisma.editorLease.findFirst({
    select: {
      expiresAt: true,
      id: true,
      workspaceBaseDocumentKey: true,
      workspaceBaseRevision: true,
      workspaceDocumentKey: true,
      workspaceObjectKey: true,
    },
    where: {
      holderSessionId: identity.sessionId,
      holderUserId: identity.id,
      targetId: response.id,
      targetType: OperationTargetType.correction,
    },
  });
  let lease: EditorLeaseGrant;
  if (
    priorLease &&
    priorLease.expiresAt > now &&
    priorLease.workspaceBaseDocumentKey === baseDocumentKey &&
    priorLease.workspaceBaseRevision === baseRevision &&
    priorLease.workspaceDocumentKey &&
    priorLease.workspaceObjectKey &&
    (await objectExists(priorLease.workspaceObjectKey))
  ) {
    const renewed = await renewEditorLease(identity, priorLease.id);
    lease = {
      ...renewed,
      proof: editorLeaseProof(identity, "correction", response.id),
      workspaceBaseDocumentKey: priorLease.workspaceBaseDocumentKey,
      workspaceBaseRevision: priorLease.workspaceBaseRevision,
      workspaceDocumentKey: priorLease.workspaceDocumentKey,
      workspaceObjectKey: priorLease.workspaceObjectKey,
    };
  } else {
    if (priorLease) {
      await releaseEditorLease(identity, priorLease.id);
    }
    const workspace: CorrectionWorkspaceInput = {
      baseDocumentKey,
      baseRevision,
      documentKey: `correction-workspace-${response.id}-${crypto.randomUUID()}`,
      objectKey: objectKey(
        "responses",
        response.id,
        "correction-workspaces",
        crypto.randomUUID(),
        "docx"
      ),
    };
    const workspaceCleanupAfter = new Date(Date.now() + operationTimeoutMs);
    await prisma.objectCleanupIntent.upsert({
      create: {
        cleanupAfter: workspaceCleanupAfter,
        objectKey: workspace.objectKey,
      },
      update: { cleanupAfter: workspaceCleanupAfter },
      where: { objectKey: workspace.objectKey },
    });
    await putObject(
      workspace.objectKey,
      await readObject(sourceObjectKey),
      DOCX_CONTENT_TYPE
    );
    try {
      lease = await claimEditorLease(
        identity,
        "correction",
        response.id,
        form.id,
        workspace
      );
    } catch (error) {
      if (await deleteObjectUnlessCanonical(workspace.objectKey)) {
        await prisma.objectCleanupIntent.deleteMany({
          where: { objectKey: workspace.objectKey },
        });
      }
      throw error;
    }
    if (
      lease.workspaceDocumentKey !== workspace.documentKey ||
      lease.workspaceObjectKey !== workspace.objectKey
    ) {
      if (await deleteObjectUnlessCanonical(workspace.objectKey)) {
        await prisma.objectCleanupIntent.deleteMany({
          where: { objectKey: workspace.objectKey },
        });
      }
      fail(
        409,
        "editor_lease_inactive",
        "The correction workspace is unavailable"
      );
    }
    await prisma.objectCleanupIntent.deleteMany({
      where: { objectKey: workspace.objectKey },
    });
    if (
      priorLease?.workspaceObjectKey &&
      priorLease.workspaceObjectKey !== lease.workspaceObjectKey
    ) {
      await deleteObjectUnlessCanonical(priorLease.workspaceObjectKey);
    }
  }
  const currentRevision = await prisma.response.findUnique({
    select: {
      corrections: {
        orderBy: { revision: "desc" },
        select: {
          documentKey: true,
          id: true,
          objectKey: true,
          revision: true,
        },
        take: 1,
      },
      submission: { select: { documentKey: true, objectKey: true } },
    },
    where: { id: response.id },
  });
  const currentLatest = currentRevision?.corrections[0];
  const currentBaseRevision = currentLatest?.revision ?? 0;
  const currentBaseDocumentKey =
    currentLatest?.documentKey ?? currentRevision?.submission?.documentKey;
  const currentSourceObjectKey =
    currentLatest?.objectKey ?? currentRevision?.submission?.objectKey;
  if (
    currentBaseRevision !== baseRevision ||
    currentBaseDocumentKey !== baseDocumentKey ||
    currentSourceObjectKey !== sourceObjectKey
  ) {
    await releaseEditorLease(identity, lease.id);
    fail(
      409,
      "stale_document",
      "The Response changed while the correction editor was opening"
    );
  }
  await createResponseAudit({
    action: currentLatest ? "view_correction" : "view_response",
    actorId: identity.id,
    outcome: AuditOutcome.success,
    safeMetadata: { revision: currentBaseRevision, state: "submitted" },
    targetId: currentLatest?.id ?? response.id,
    targetType: currentLatest ? "correction" : "response",
  });
  const capabilityScope = {
    documentKey: lease.workspaceDocumentKey as string,
    formId: form.id,
    targetId: response.id,
    targetType: "correction",
  } as const;
  return editorConfig(
    {
      action: "correction",
      capabilities: {
        "save-correction": actionEditorCapability(
          identity,
          capabilityScope,
          "save-correction",
          lease
        ),
      },
      documentKey: capabilityScope.documentKey,
      lease: editorLeaseBridge(lease),
      prefill: lockedPrefillForSnapshot(response.prefillSnapshot),
      publicId: form.publicId,
      responseId: response.id,
      tagAliases: form.publishedTemplate?.objectKey
        ? onlyOfficeFieldTagAliases(
            await readObject(form.publishedTemplate.objectKey)
          )
        : {},
    },
    identity
  );
}

export async function userEditorConfig(
  form: FormWithDocuments,
  identity: Identity,
  responseId: string | undefined,
  requestedAction?: string
): Promise<Record<string, unknown>> {
  const publishedTemplate = form.publishedTemplate;
  if (!publishedTemplate?.objectKey || !publishedTemplate.documentKey) {
    fail(409, "not_published", "This form has no published document");
  }
  const response = responseId
    ? await prisma.response.findFirst({
        include: { prefillSnapshot: true },
        where: {
          formId: form.id,
          id: responseId,
          userId: identity.id,
        },
      })
    : await prisma.response.findFirst({
        include: { prefillSnapshot: true },
        where: { formId: form.id, userId: identity.id },
      });
  if (!response) {
    fail(404, "not_found", "Start a response before opening the editor");
  }
  if (response.status === ResponseStatus.submitted) {
    fail(409, "already_submitted", "This response has already been submitted");
  }
  if (!response.draftDocumentKey || !response.draftObjectKey) {
    fail(409, "document_unavailable", "Response document is unavailable");
  }
  if (!(await objectExists(response.draftObjectKey))) {
    fail(
      409,
      "document_unavailable",
      "Response document artifact is unavailable"
    );
  }
  const snapshot = response.prefillSnapshot;
  const capabilityScope = {
    documentKey: response.draftDocumentKey,
    formId: form.id,
    targetId: response.id,
    targetType: "response",
  } as const;
  const lease = await claimEditorLease(
    identity,
    capabilityScope.targetType,
    capabilityScope.targetId,
    form.id
  );
  if (form.fillMethod === FillMethod.native) {
    const nativeFieldConfig = await nativeFields(
      publishedTemplate.id,
      await readObject(response.draftObjectKey)
    );
    return {
      capabilities: {
        "save-draft": actionEditorCapability(
          identity,
          capabilityScope,
          "save-draft",
          lease
        ),
        submit: actionEditorCapability(
          identity,
          capabilityScope,
          "submit",
          lease
        ),
      },
      data:
        response.draftData === null
          ? jsonRecord(snapshot?.values ?? {})
          : jsonRecord(response.draftData),
      documentKey: response.draftDocumentKey,
      fields: nativeFieldConfig.fields,
      fillMethod: FillMethod.native,
      lockedFields: snapshot ? jsonRecord(snapshot.lockedFields) : {},
      pictures: nativeFieldConfig.pictures,
      responseId: response.id,
    };
  }
  return {
    ...editorConfig(
      {
        action:
          requestedAction === "submit"
            ? "submit"
            : requestedAction === "fill"
              ? "fill"
              : "draft",
        capabilities: {
          "save-draft": actionEditorCapability(
            identity,
            capabilityScope,
            "save-draft",
            lease
          ),
          submit: actionEditorCapability(
            identity,
            capabilityScope,
            "submit",
            lease
          ),
        },
        documentKey: response.draftDocumentKey,
        lease: editorLeaseBridge(lease),
        prefill: snapshot
          ? requestedAction === "fill"
            ? {
                data: jsonRecord(snapshot.values),
                editableFields: editableFieldsForSnapshot(snapshot),
              }
            : lockedPrefillForSnapshot(snapshot)
          : undefined,
        publicId: form.publicId,
        responseId: response.id,
        tagAliases: onlyOfficeFieldTagAliases(
          await readObject(publishedTemplate.objectKey)
        ),
      },
      identity
    ),
    fillMethod: FillMethod.onlyoffice,
  };
}
