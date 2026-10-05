// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { zipSync, strToU8, unzipSync } from "fflate";

import type { createApp } from "../../../src/app";
import { pluginGuid, verifyEditorCapability } from "../../../src/onlyoffice";
import {
  readObject,
  objectExists,
  DOCX_CONTENT_TYPE,
  putObject,
} from "../../../src/storage";
import { jsonHeaders, waitForOperation } from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import type { FieldConfigurationOutput } from "./editor-access";
import {
  refreshResponseEditor,
  capabilityHeaders,
  draftRequest,
  submitRequest,
} from "./helpers";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";

export interface DraftLifecycleInput {
  app: ReturnType<typeof createApp>;
  formRecord: PrefillEntryOutput["formRecord"];
  userBearer: PrefillEntryOutput["userBearer"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  formId: BootstrapAndCreationOutput["formId"];
  userId: PrefillEntryOutput["userId"];
  selectedPointer: FieldConfigurationOutput["selectedPointer"];
  secretFormTitle: BootstrapAndCreationOutput["secretFormTitle"];
}

export interface DraftLifecycleOutput {
  responseId: string;
  responseDocumentKey: string;
  saveDraftCapability: string;
  submitCapability: string;
  userLease: {
    expiresAt: string;
    id: string;
    releaseUrl: string;
    renewUrl: string;
  };
  savedDraftData: {
    accept_terms: boolean;
    department: string;
    description_1: string;
    description_2: string;
    full_name: string | boolean;
    start_date: string;
  };
  saveBody: { operationCapability?: string; operationId?: string } & {
    operationId: string;
    operationCapability: string;
  };
  savedResponseDocumentKey: string;
  savedResponseObjectKey: string;
  savedResponseDocumentBytes: Uint8Array<ArrayBufferLike>;
}

export const runDraftLifecycle = async (
  input: DraftLifecycleInput
): Promise<DraftLifecycleOutput> => {
  const {
    app,
    formRecord,
    userBearer,
    adminBearer,
    formPublicId,
    formId,
    userId,
    selectedPointer,
    secretFormTitle,
  } = input;

  const startResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(startResponse.status).toBe(200);
  const startBody = (await startResponse.json()) as {
    response?: { id?: string };
  };
  const responseId = startBody.response?.id;
  if (!responseId) {
    throw new Error("The response was not started");
  }
  const draftCountListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(draftCountListResponse.status).toBe(200);
  const draftCountList = (await draftCountListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      submissionCount: number;
    }[];
  };
  expect(
    draftCountList.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({ activeDraftCount: 1, submissionCount: 0 });
  const formLifecycleBeforeResponseGuard = await prisma.form.findUnique({
    select: { status: true, version: true },
    where: { id: formId },
  });
  if (!formLifecycleBeforeResponseGuard) {
    throw new Error("The Form lifecycle guard baseline was not found");
  }
  await prisma.form.update({
    data: { status: "draft", version: 0 },
    where: { id: formId },
  });
  const responseProtectedDelete = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(responseProtectedDelete.status).toBe(409);
  expect(await responseProtectedDelete.json()).toMatchObject({
    error: "form_has_responses",
  });
  const responseProtectedTemplate = await prisma.templateDraft.findUnique({
    select: { objectKey: true },
    where: { formId },
  });
  expect(responseProtectedTemplate).not.toBeNull();
  if (responseProtectedTemplate) {
    expect(await objectExists(responseProtectedTemplate.objectKey)).toBe(true);
  }
  await prisma.form.update({
    data: formLifecycleBeforeResponseGuard,
    where: { id: formId },
  });

  const editorResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${formRecord.publicId}/editor-config?responseId=${responseId}&action=fill`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(editorResponse.status).toBe(200);
  const editorConfig = (await editorResponse.json()) as EditorConfigBody;
  let responseDocumentKey = editorConfig.config.document.key;
  const originalResponseDocumentKey = responseDocumentKey;
  const userPluginOptions =
    editorConfig.config.editorConfig.plugins.options[pluginGuid];
  let saveDraftCapability: string =
    editorConfig.bridge.capabilities["save-draft"] ?? "";
  let submitCapability: string = editorConfig.bridge.capabilities.submit ?? "";
  if (
    !userPluginOptions ||
    !responseDocumentKey ||
    !saveDraftCapability ||
    !submitCapability
  ) {
    throw new Error("The User editor capabilities were not returned");
  }
  let userLease = editorConfig.bridge.lease;
  expect(userLease).toEqual({
    expiresAt: expect.any(String),
    id: expect.any(String),
    releaseUrl: expect.any(String),
    renewUrl: expect.any(String),
  });
  for (const leaseValue of Object.values(userLease)) {
    expect(JSON.stringify(editorConfig.config)).not.toContain(leaseValue);
  }
  expect(userPluginOptions).not.toHaveProperty("lease");
  const originalSaveDraftCapability = saveDraftCapability;
  await prisma.editorLease.update({
    data: { createdAt: new Date(0), expiresAt: new Date(1) },
    where: { id: userLease.id },
  });

  const previousUserLeaseId = userLease.id;
  ({ responseDocumentKey, saveDraftCapability, submitCapability, userLease } =
    await refreshResponseEditor(
      app,
      formRecord.publicId,
      responseId,
      userBearer
    ));
  expect(userLease.id).not.toBe(previousUserLeaseId);
  const staleUserLeaseRelease = await app.handle(
    new Request(`http://test.local/api/editor-leases/${previousUserLeaseId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(staleUserLeaseRelease.status).toBe(409);
  expect(await staleUserLeaseRelease.json()).toMatchObject({
    error: "editor_lease_inactive",
  });
  const activeLeaseOperationId = crypto.randomUUID();
  await prisma.operation.create({
    data: {
      actorId: userId,
      documentKey: responseDocumentKey,
      formId,
      id: activeLeaseOperationId,
      metadata: {
        action: "save-draft",
        finalObjectKey: `operations/${activeLeaseOperationId}/final.docx`,
        formId,
        responseId,
        stagedObjectKey: `operations/${activeLeaseOperationId}/staged.docx`,
      },
      ownerUserId: userId,
      responseId,
      stagingObjectKey: `operations/${activeLeaseOperationId}/staged.docx`,
      status: "processing",
      targetId: responseId,
      targetType: "response",
      type: "save_draft",
    },
  });
  const releaseDuringOperationResponse = await app.handle(
    new Request(`http://test.local${userLease.releaseUrl}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(releaseDuringOperationResponse.status).toBe(200);
  expect(
    await prisma.editorLease.findUnique({ where: { id: userLease.id } })
  ).not.toBeNull();
  await prisma.operation.delete({ where: { id: activeLeaseOperationId } });
  const sessionOnlyUserDraftResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyUserDraftResponse.status).toBe(401);
  expect(await sessionOnlyUserDraftResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const sessionOnlyUserSubmitResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/submit`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyUserSubmitResponse.status).toBe(401);
  expect(await sessionOnlyUserSubmitResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const userEditorSerialized = JSON.stringify(editorConfig);
  expect(userEditorSerialized).not.toContain(userBearer);
  expect(userEditorSerialized).not.toContain('"authToken"');
  expect(JSON.stringify(editorConfig.config)).not.toContain(
    saveDraftCapability
  );
  expect(editorConfig.bridge.id).toBe(userPluginOptions.bridgeId);
  const saveDraftClaims = verifyEditorCapability(saveDraftCapability);
  expect(saveDraftClaims).toMatchObject({
    action: "save-draft",
    actorId: userId,
    documentKey: responseDocumentKey,
    formId,
    leaseId: userLease.id,
    leaseProof: expect.any(String),
    role: "user",
    targetId: responseId,
    targetType: "response",
  });
  if (!saveDraftClaims) {
    throw new Error("The save-draft capability was invalid");
  }
  const submitClaims = verifyEditorCapability(submitCapability);
  expect(submitClaims).toMatchObject({
    action: "submit",
    actorId: userId,
    documentKey: responseDocumentKey,
    formId,
    leaseId: userLease.id,
    leaseProof: expect.any(String),
    role: "user",
    targetId: responseId,
    targetType: "response",
  });
  if (!submitClaims) {
    throw new Error("The submit capability was invalid");
  }
  await prisma.user.update({
    data: { enabled: false },
    where: { id: userId },
  });
  const disabledActorResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect(disabledActorResponse.status).toBe(401);
  await prisma.user.update({
    data: { enabled: true, role: "admin" },
    where: { id: userId },
  });
  const changedRoleResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect(changedRoleResponse.status).toBe(401);
  await prisma.user.update({
    data: { role: "user" },
    where: { id: userId },
  });
  const userCrossActionResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(submitCapability),
      method: "POST",
    })
  );
  expect(userCrossActionResponse.status).toBe(403);
  const invalidDraftCases = [
    { data: { unknown_tag: "value" }, status: 422 },
    { data: { accept_terms: "true" }, status: 422 },
    { data: { department: "not-an-option" }, status: 422 },
    { data: { start_date: "2026-02-30" }, status: 422 },
    { data: { description_1: "x".repeat(10_001) }, status: 422 },
    {
      data: { description_1: "🙂".repeat(70_000) },
      error: "response_too_large",
      status: 413,
    },
    {
      data: { description_1: "🙂".repeat(100_000) },
      error: "payload_too_large",
      status: 413,
    },
  ];
  for (const invalidDraftCase of invalidDraftCases) {
    const invalidDraftResponse = await draftRequest(
      app,
      formRecord.publicId,
      responseDocumentKey,
      responseId,
      saveDraftCapability,
      invalidDraftCase.data
    );
    expect(invalidDraftResponse.status).toBe(invalidDraftCase.status);
    expect(await invalidDraftResponse.json()).toHaveProperty(
      "error",
      invalidDraftCase.error ?? "invalid_response_data"
    );
  }
  const trustedPrefillValue = selectedPointer.endsWith("/active")
    ? true
    : "Ticket 17 Prefill";
  const lockedSnapshotCases = [
    { error: "invalid_response_data", status: 422, value: "x".repeat(10_001) },
    { error: "response_too_large", status: 413, value: "🙂".repeat(100_000) },
    { error: "invalid_response_data", status: 422, value: true },
  ];
  for (const lockedSnapshotCase of lockedSnapshotCases) {
    await prisma.prefillSnapshot.update({
      data: { values: { full_name: lockedSnapshotCase.value } },
      where: { responseId },
    });
    const invalidTrustedValueResponse = await draftRequest(
      app,
      formRecord.publicId,
      responseDocumentKey,
      responseId,
      saveDraftCapability,
      {
        full_name: "client value",
      }
    );
    expect(invalidTrustedValueResponse.status).toBe(lockedSnapshotCase.status);
    expect(await invalidTrustedValueResponse.json()).toMatchObject({
      error: lockedSnapshotCase.error,
    });
  }
  await prisma.prefillSnapshot.update({
    data: { values: { full_name: trustedPrefillValue } },
    where: { responseId },
  });
  const draftDataBeforeLockedEdit = await prisma.response.findUniqueOrThrow({
    select: { draftData: true },
    where: { id: responseId },
  });
  const lockedClientEditResponse = await draftRequest(
    app,
    formRecord.publicId,
    responseDocumentKey,
    responseId,
    saveDraftCapability,
    {
      full_name: "Tampered client value",
    }
  );
  expect(lockedClientEditResponse.status).toBe(422);
  expect(await lockedClientEditResponse.json()).toMatchObject({
    error: "invalid_response_data",
  });
  expect(
    await prisma.response.findUniqueOrThrow({
      select: { draftData: true },
      where: { id: responseId },
    })
  ).toEqual(draftDataBeforeLockedEdit);
  const savedDraftData = {
    accept_terms: true,
    department: "engineering",
    description_1: "line one\nline two",
    description_2: "",
    full_name: selectedPointer.endsWith("/active") ? true : "Ticket 17 Prefill",
    start_date: "2026-09-15",
  };
  const invalidSubmitCases = [
    { ...savedDraftData, accept_terms: false },
    { ...savedDraftData, department: "not-an-option" },
    { ...savedDraftData, start_date: "2026-02-30" },
  ];
  for (const invalidSubmitData of invalidSubmitCases) {
    const invalidSubmitResponse = await submitRequest(
      app,
      formRecord.publicId,
      responseDocumentKey,
      responseId,
      submitCapability,
      invalidSubmitData
    );
    expect(invalidSubmitResponse.status).toBe(422);
    expect(await invalidSubmitResponse.json()).toMatchObject({
      error: "invalid_response_data",
    });
  }

  const draftBeforeClear = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: responseId },
  });
  if (!draftBeforeClear.draftObjectKey) {
    throw new Error("The editable response document is missing");
  }
  const clearedDraftArchive = unzipSync(
    await readObject(draftBeforeClear.draftObjectKey)
  );
  const draftBeforeClearXml = new TextDecoder().decode(
    clearedDraftArchive["word/document.xml"]
  );
  const descriptionControl =
    /(?<opening><w:tag\b[^>]*\bw:val="description_2"\s*\/>[\s\S]*?<w:sdtContent\b[^>]*>)(?<content>[\s\S]*?)<\/w:sdtContent>/u;
  const descriptionContent =
    descriptionControl.exec(draftBeforeClearXml)?.groups?.content;
  if (descriptionContent === undefined) {
    throw new Error("The editable description field is missing");
  }
  // Model the editor clearing this field before its unchanged force-save reply.
  const clearedDescriptionContent = descriptionContent
    .replaceAll(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/gu, "<w:t></w:t>")
    .replaceAll(/<w:br\b[^>]*\/>/gu, "");
  clearedDraftArchive["word/document.xml"] = strToU8(
    draftBeforeClearXml.replace(
      descriptionControl,
      `$<opening>${clearedDescriptionContent}</w:sdtContent>`
    )
  );
  await putObject(
    draftBeforeClear.draftObjectKey,
    zipSync(clearedDraftArchive),
    DOCX_CONTENT_TYPE
  );

  const saveResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: savedDraftData,
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect(saveResponse.status).toBe(202);
  const saveBody = (await saveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!saveBody.operationCapability || !saveBody.operationId) {
    throw new Error("The draft operation was not created");
  }
  const saveOperation = await waitForOperation(app, saveBody.operationId, {
    "X-Editor-Capability": saveBody.operationCapability,
  });
  expect(saveOperation.status).toBe("completed");
  const saveOperationResult = saveOperation.result as {
    documentKey?: unknown;
    responseId?: unknown;
  };
  expect(saveOperationResult.documentKey).toEqual(expect.any(String));
  expect(saveOperationResult.responseId).toBe(responseId);
  const savedResponse = await prisma.response.findUnique({
    select: {
      draftData: true,
      draftDocumentKey: true,
      draftObjectKey: true,
      status: true,
      updatedAt: true,
    },
    where: { id: responseId },
  });
  const savedResponseDocumentKey = savedResponse?.draftDocumentKey;
  const savedResponseObjectKey = savedResponse?.draftObjectKey;
  expect(savedResponse).toMatchObject({
    draftData: savedDraftData,
    draftDocumentKey: expect.any(String),
    status: "draft",
    updatedAt: expect.any(Date),
  });
  if (!savedResponseDocumentKey || !savedResponseObjectKey) {
    throw new Error("The saved response document was not persisted");
  }
  expect(savedResponseDocumentKey).not.toBe(originalResponseDocumentKey);
  expect(saveOperationResult.documentKey).toBe(savedResponseDocumentKey);
  const savedResponseDocumentBytes = await readObject(savedResponseObjectKey);
  const staleDraftResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: savedDraftData,
        documentKey: originalResponseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(originalSaveDraftCapability),
      method: "POST",
    })
  );
  expect(staleDraftResponse.status).toBe(409);
  expect(await staleDraftResponse.json()).toMatchObject({
    error: "stale_response",
  });
  responseDocumentKey = savedResponseDocumentKey;
  const resumedStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(resumedStartResponse.status).toBe(200);
  expect(await resumedStartResponse.json()).toMatchObject({
    response: { id: responseId, status: "draft" },
  });
  const resumedEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${formRecord.publicId}/editor-config?responseId=${responseId}&action=draft`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(resumedEditorResponse.status).toBe(200);
  const resumedEditorConfig =
    (await resumedEditorResponse.json()) as EditorConfigBody;
  expect(resumedEditorConfig).toMatchObject({
    config: { document: { key: responseDocumentKey } },
  });
  const resumedSaveDraftCapability =
    resumedEditorConfig.bridge.capabilities["save-draft"];
  const resumedSubmitCapability =
    resumedEditorConfig.bridge.capabilities.submit;
  if (!resumedSaveDraftCapability || !resumedSubmitCapability) {
    throw new Error("The resumed User editor capabilities were not returned");
  }
  saveDraftCapability = resumedSaveDraftCapability;
  submitCapability = resumedSubmitCapability;
  userLease = resumedEditorConfig.bridge.lease;
  const responseList = await app.handle(
    new Request("http://test.local/api/responses/me", {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(await responseList.json()).toMatchObject({
    responses: [
      {
        formPublicId: formRecord.publicId,
        formTitle: secretFormTitle,
        id: responseId,
        status: "draft",
        updatedAt: expect.any(String),
      },
    ],
  });
  return {
    responseDocumentKey,
    responseId,
    saveBody: {
      ...saveBody,
      operationCapability: saveBody.operationCapability,
      operationId: saveBody.operationId,
    },
    saveDraftCapability,
    savedDraftData,
    savedResponseDocumentBytes,
    savedResponseDocumentKey,
    savedResponseObjectKey,
    submitCapability,
    userLease,
  };
};
