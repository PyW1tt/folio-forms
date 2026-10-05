import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";

import type { createApp } from "../../../src/app";
import {
  createEditorCapability,
  createCallbackUserdata,
} from "../../../src/onlyoffice";
import { reconcileRecoverableState } from "../../../src/operations/recovery";
import {
  readObject,
  objectExists,
  DOCX_CONTENT_TYPE,
  putObject,
  objectKey,
} from "../../../src/storage";
import { docxFixture } from "../../fixtures/documents";
import { jsonHeaders, persistCallbackClaim } from "../../fixtures/http";
import type { FormLifecycleOutput } from "./form-lifecycle";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";
import type { SubmissionExportsOutput } from "./submission";

export interface DiscardAndRestartInput {
  app: ReturnType<typeof createApp>;
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  archivedNoResponseEmail: FormLifecycleOutput["archivedNoResponseEmail"];
  handoffCandidateValues: PrefillEntryOutput["handoffCandidateValues"];
  prefillHandoffSecret: PrefillEntryOutput["prefillHandoffSecret"];
  archivedNoResponseBearer: FormLifecycleOutput["archivedNoResponseBearer"];
  userBearer: PrefillEntryOutput["userBearer"];
  adminId: BootstrapAndCreationOutput["adminId"];
  formId: BootstrapAndCreationOutput["formId"];
  publishedContractBefore: FormLifecycleOutput["publishedContractBefore"];
}

export const runDiscardAndRestart = async (
  input: DiscardAndRestartInput
): Promise<void> => {
  const {
    app,
    formPublicId,
    adminBearer,
    archivedNoResponseEmail,
    handoffCandidateValues,
    prefillHandoffSecret,
    archivedNoResponseBearer,
    userBearer,
    adminId,
    formId,
    publishedContractBefore,
  } = input;
  const unarchiveResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({ status: "published" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(unarchiveResponse.status).toBe(200);
  expect(await unarchiveResponse.json()).toMatchObject({
    form: {
      publicId: formPublicId,
      status: "published",
      version: 1,
    },
  });
  const archivedHandoffReference = `ticket-17-archived-${crypto.randomUUID()}`;
  const archivedHandoffCreate = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: archivedNoResponseEmail,
        externalReference: archivedHandoffReference,
        publicId: formPublicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(archivedHandoffCreate.status).toBe(200);
  const archivedHandoffCode = (
    (await archivedHandoffCreate.json()) as { code: string }
  ).code;
  const archivedHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: archivedHandoffCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    })
  );
  expect(archivedHandoffLaunch.status).toBe(303);
  const archivedPendingCookie = archivedHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!archivedPendingCookie) {
    throw new Error("The archived handoff did not set a pending claim cookie");
  }
  const unarchivedStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}/start`, {
      headers: {
        Authorization: `Bearer ${archivedNoResponseBearer}`,
        Cookie: archivedPendingCookie,
      },
      method: "POST",
    })
  );
  expect(unarchivedStartResponse.status).toBe(200);
  const unarchivedStartBody = (await unarchivedStartResponse.json()) as {
    response?: { id?: string };
  };
  const discardedResponseId = unarchivedStartBody.response?.id;
  if (!discardedResponseId) {
    throw new Error("The discard fixture was not created");
  }
  const discardResponseBefore = await prisma.response.findUnique({
    select: { draftObjectKey: true },
    where: { id: discardedResponseId },
  });
  if (!discardResponseBefore?.draftObjectKey) {
    throw new Error("The discard Draft document was not created");
  }
  const discardEditorConfigResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${formPublicId}/editor-config?responseId=${discardedResponseId}&action=draft`,
      { headers: { Authorization: `Bearer ${archivedNoResponseBearer}` } }
    )
  );
  expect(discardEditorConfigResponse.status).toBe(200);
  const unauthorizedDiscard = await app.handle(
    new Request(`http://test.local/api/responses/${discardedResponseId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(unauthorizedDiscard.status).toBe(403);
  const discardResponse = await app.handle(
    new Request(`http://test.local/api/responses/${discardedResponseId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "DELETE",
    })
  );
  expect(discardResponse.status).toBe(200);
  expect(await discardResponse.json()).toEqual({ discarded: true });
  expect(
    await prisma.response.findUnique({ where: { id: discardedResponseId } })
  ).toBeNull();
  expect(
    await prisma.editorLease.count({
      where: {
        targetId: discardedResponseId,
        targetType: "response",
      },
    })
  ).toBe(0);
  expect(
    await prisma.operation.count({ where: { responseId: discardedResponseId } })
  ).toBe(0);
  expect(await objectExists(discardResponseBefore.draftObjectKey)).toBe(false);
  const repeatedDiscardResponse = await app.handle(
    new Request(`http://test.local/api/responses/${discardedResponseId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "DELETE",
    })
  );
  expect(repeatedDiscardResponse.status).toBe(200);
  expect(await repeatedDiscardResponse.json()).toEqual({ discarded: true });
  const responsesAfterDiscard = await app.handle(
    new Request("http://test.local/api/responses/me", {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
    })
  );
  expect(await responsesAfterDiscard.json()).toEqual({ responses: [] });
  const restartHandoffReference = `ticket-17-restart-${crypto.randomUUID()}`;
  const restartHandoffCreate = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: archivedNoResponseEmail,
        externalReference: restartHandoffReference,
        publicId: formPublicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(restartHandoffCreate.status).toBe(200);
  const restartHandoffCode = (
    (await restartHandoffCreate.json()) as { code: string }
  ).code;
  const restartHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: restartHandoffCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    })
  );
  expect(restartHandoffLaunch.status).toBe(303);
  const restartPendingCookie = restartHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!restartPendingCookie) {
    throw new Error("The restart handoff did not set a pending claim cookie");
  }
  const restartAfterDiscard = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}/start`, {
      headers: {
        Authorization: `Bearer ${archivedNoResponseBearer}`,
        Cookie: restartPendingCookie,
      },
      method: "POST",
    })
  );
  expect(restartAfterDiscard.status).toBe(200);
  const restartAfterDiscardBody = (await restartAfterDiscard.json()) as {
    response?: { id?: string };
  };
  const restartedResponseId = restartAfterDiscardBody.response?.id;
  if (!restartedResponseId) {
    throw new Error("The response was not restartable after discard");
  }
  expect(restartedResponseId).not.toBe(discardedResponseId);
  const discardRestartResponse = await app.handle(
    new Request(`http://test.local/api/responses/${restartedResponseId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "DELETE",
    })
  );
  expect(discardRestartResponse.status).toBe(200);
  const unarchivedFormListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  const unarchivedFormListBody = (await unarchivedFormListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      status: string;
      submissionCount: number;
    }[];
  };
  expect(
    unarchivedFormListBody.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({
    activeDraftCount: 0,
    publicId: formPublicId,
    status: "published",
    submissionCount: 1,
  });
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "unarchive_form",
        outcome: "success",
        targetId: formPublicId,
      },
    })
  ).toMatchObject({ actorId: adminId, targetId: formPublicId });
  expect(
    await prisma.publishedTemplate.findUnique({
      include: {
        manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
        prefillConfiguration: {
          include: { fields: { orderBy: { tag: "asc" } } },
        },
      },
      where: { formId },
    })
  ).toEqual(publishedContractBefore);
};

export interface RecoverableStateInput {
  app: ReturnType<typeof createApp>;
  templateDraft: SubmissionExportsOutput["templateDraft"];
  sourceDocument: SubmissionExportsOutput["sourceDocument"];
  adminId: BootstrapAndCreationOutput["adminId"];
  formId: BootstrapAndCreationOutput["formId"];
  otherUser: SubmissionExportsOutput["otherUser"];
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
}

export const runRecoverableState = async (
  input: RecoverableStateInput
): Promise<void> => {
  const {
    app,
    templateDraft,
    sourceDocument,
    adminId,
    formId,
    otherUser,
    formPublicId,
    adminBearer,
  } = input;
  const staleOperationId = crypto.randomUUID();
  const staleStagingObjectKey = objectKey(
    "operations",
    staleOperationId,
    "staged.docx"
  );
  const staleFinalObjectKey = objectKey(
    "operations",
    staleOperationId,
    "final.docx"
  );
  const staleOperationUserdata = createCallbackUserdata({
    documentKey: templateDraft.documentKey,
    expiresAt: Math.floor(Date.now() / 1000) - 1,
    operationId: staleOperationId,
    operationType: "save_template_draft",
  });
  await Promise.all([
    putObject(staleStagingObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
    putObject(staleFinalObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
  ]);
  await prisma.operation.create({
    data: {
      actorId: adminId,
      documentKey: templateDraft.documentKey,
      errorCode: null,
      formId,
      id: staleOperationId,
      metadata: {
        action: "save-template",
        finalObjectKey: staleFinalObjectKey,
        formId,
        stagedObjectKey: staleStagingObjectKey,
      },
      ownerUserId: adminId,
      stagingObjectKey: staleStagingObjectKey,
      status: "processing",
      targetId: templateDraft.id,
      targetType: "template_draft",
      type: "save_template_draft",
      updatedAt: new Date(0),
    },
  });
  await persistCallbackClaim({
    expiresAt: new Date(0),
    operationId: staleOperationId,
    userdata: staleOperationUserdata,
  });
  const staleOperationCapability = createEditorCapability({
    action: "poll-operation",
    actorId: adminId,
    documentKey: templateDraft.documentKey,
    formId,
    operationId: staleOperationId,
    role: "admin",
    targetId: templateDraft.id,
    targetType: "template-draft",
  });

  const publishedTemplate = await prisma.publishedTemplate.findUnique({
    select: { id: true, version: true },
    where: { formId },
  });
  if (!publishedTemplate) {
    throw new Error("The published template was not found");
  }
  const staleResponseId = crypto.randomUUID();
  const staleResponseDocumentKey = `response-${staleResponseId}-${crypto.randomUUID()}`;
  const staleResponseObjectKey = objectKey(
    "responses",
    staleResponseId,
    "draft",
    crypto.randomUUID(),
    "docx"
  );
  const staleSubmitOperationId = crypto.randomUUID();
  const staleSubmitStagingObjectKey = objectKey(
    "operations",
    staleSubmitOperationId,
    "staged.docx"
  );
  const staleSubmitFinalObjectKey = objectKey(
    "operations",
    staleSubmitOperationId,
    "final.docx"
  );
  const staleSubmissionId = crypto.randomUUID();
  const staleSubmitUserdata = createCallbackUserdata({
    documentKey: staleResponseDocumentKey,
    expiresAt: Math.floor(Date.now() / 1000) - 1,
    operationId: staleSubmitOperationId,
    operationType: "submit_response",
  });
  await Promise.all([
    putObject(staleResponseObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
    putObject(staleSubmitStagingObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
    putObject(staleSubmitFinalObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
  ]);
  await prisma.response.create({
    data: {
      draftData: {},
      draftDocumentKey: staleResponseDocumentKey,
      draftObjectKey: staleResponseObjectKey,
      form: { connect: { id: formId } },
      id: staleResponseId,
      owner: { connect: { id: otherUser.id } },
      publishedTemplate: { connect: { id: publishedTemplate.id } },
      publishedVersion: publishedTemplate.version,
      status: "submitting",
    },
  });
  await prisma.operation.create({
    data: {
      actorId: otherUser.id,
      documentKey: staleResponseDocumentKey,
      errorCode: null,
      formId,
      id: staleSubmitOperationId,
      metadata: {
        action: "submit",
        finalObjectKey: staleSubmitFinalObjectKey,
        formId,
        responseId: staleResponseId,
        stagedObjectKey: staleSubmitStagingObjectKey,
        submissionDocumentKey: `submission-${staleSubmissionId}-${crypto.randomUUID()}`,
        submissionId: staleSubmissionId,
      },
      ownerUserId: otherUser.id,
      responseId: staleResponseId,
      stagingObjectKey: staleSubmitStagingObjectKey,
      status: "processing",
      targetId: staleResponseId,
      targetType: "response",
      type: "submit_response",
      updatedAt: new Date(0),
    },
  });
  await persistCallbackClaim({
    expiresAt: new Date(0),
    operationId: staleSubmitOperationId,
    userdata: staleSubmitUserdata,
  });

  const expiredLease = await prisma.editorLease.findUnique({
    where: {
      targetType_targetId: {
        targetId: templateDraft.id,
        targetType: "template_draft",
      },
    },
  });
  if (!expiredLease) {
    throw new Error("The active template lease was not found");
  }
  await prisma.editorLease.update({
    data: { createdAt: new Date(0), expiresAt: new Date(1) },
    where: { id: expiredLease.id },
  });
  await prisma.operation.update({
    data: { updatedAt: new Date() },
    where: { id: staleOperationId },
  });
  const cleanupIntentObjectKey = objectKey(
    "cleanup-intents",
    crypto.randomUUID(),
    "orphan.docx"
  );
  await putObject(
    cleanupIntentObjectKey,
    docxFixture("ticket-08-cleanup-intent"),
    DOCX_CONTENT_TYPE
  );
  await prisma.objectCleanupIntent.create({
    data: { objectKey: cleanupIntentObjectKey },
  });
  expect(await objectExists(cleanupIntentObjectKey)).toBe(true);
  const inFlightCleanupObjectKey = objectKey(
    "cleanup-intents",
    crypto.randomUUID(),
    "in-flight.docx"
  );
  await putObject(
    inFlightCleanupObjectKey,
    docxFixture("ticket-08-in-flight-cleanup"),
    DOCX_CONTENT_TYPE
  );
  await prisma.objectCleanupIntent.create({
    data: {
      cleanupAfter: new Date(Date.now() + 60_000),
      objectKey: inFlightCleanupObjectKey,
    },
  });
  await reconcileRecoverableState();
  expect(
    await prisma.editorLease.findUnique({ where: { id: expiredLease.id } })
  ).not.toBeNull();
  expect(await objectExists(cleanupIntentObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: cleanupIntentObjectKey },
    })
  ).toBeNull();
  expect(await objectExists(inFlightCleanupObjectKey)).toBe(true);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: inFlightCleanupObjectKey },
    })
  ).not.toBeNull();
  await prisma.objectCleanupIntent.update({
    data: { cleanupAfter: new Date(0) },
    where: { objectKey: inFlightCleanupObjectKey },
  });
  await prisma.operation.update({
    data: { updatedAt: new Date(0) },
    where: { id: staleOperationId },
  });
  await reconcileRecoverableState();
  expect(await objectExists(inFlightCleanupObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: inFlightCleanupObjectKey },
    })
  ).toBeNull();

  expect(
    await prisma.operation.findUnique({
      select: { errorCode: true, status: true },
      where: { id: staleOperationId },
    })
  ).toEqual({ errorCode: "operation_timeout", status: "failed" });
  expect(
    await prisma.operation.findUnique({
      select: { errorCode: true, status: true },
      where: { id: staleSubmitOperationId },
    })
  ).toEqual({ errorCode: "operation_timeout", status: "failed" });
  expect(
    await prisma.response.findUnique({
      select: {
        draftDocumentKey: true,
        draftObjectKey: true,
        status: true,
      },
      where: { id: staleResponseId },
    })
  ).toEqual({
    draftDocumentKey: staleResponseDocumentKey,
    draftObjectKey: staleResponseObjectKey,
    status: "draft",
  });
  expect(
    await prisma.editorLease.findUnique({ where: { id: expiredLease.id } })
  ).toBeNull();
  expect(
    await prisma.callbackClaim.findUnique({
      where: { operationId: staleOperationId },
    })
  ).toBeNull();
  expect(
    await prisma.callbackClaim.findUnique({
      where: { operationId: staleSubmitOperationId },
    })
  ).toBeNull();
  expect(await objectExists(staleStagingObjectKey)).toBe(false);
  expect(await objectExists(staleFinalObjectKey)).toBe(false);
  expect(await objectExists(staleSubmitStagingObjectKey)).toBe(false);
  expect(await objectExists(staleSubmitFinalObjectKey)).toBe(false);
  expect(await objectExists(staleResponseObjectKey)).toBe(true);
  expect(await objectExists(templateDraft.objectKey)).toBe(true);
  expect(await readObject(templateDraft.objectKey)).toEqual(sourceDocument);

  const reconciledOperationResponse = await app.handle(
    new Request(`http://test.local/api/operations/${staleOperationId}`, {
      headers: { "X-Editor-Capability": staleOperationCapability },
    })
  );
  expect(reconciledOperationResponse.status).toBe(200);
  expect(await reconciledOperationResponse.json()).toMatchObject({
    operation: {
      error: "operation_timeout",
      id: staleOperationId,
      status: "failed",
    },
  });
  const reclaimedAfterReconcileResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(reclaimedAfterReconcileResponse.status).toBe(409);
  expect(await reclaimedAfterReconcileResponse.json()).toMatchObject({
    error: "published_immutable",
  });
};
