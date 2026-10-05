// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";
import { createHash } from "node:crypto";

import type { Submission, Correction } from "@onlyoffice/db";
import { prisma } from "@onlyoffice/db";

import { createApp } from "../../../src/app";
import {
  pluginGuid,
  verifyEditorCapability,
  createOnlyOfficeAuthorization,
  createCallbackUserdata,
  createOnlyOfficeBodyToken,
} from "../../../src/onlyoffice";
import {
  readObject,
  objectExists,
  DOCX_CONTENT_TYPE,
  putObject,
  objectKey,
} from "../../../src/storage";
import { docxFixture } from "../../fixtures/documents";
import { jsonHeaders, waitForOperation } from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import type { DraftLifecycleOutput } from "./draft";
import { capabilityHeaders } from "./helpers";
import type { PrefillEntryOutput } from "./prefill-entry";
import type {
  BootstrapAndCreationOutput,
  TemplateUploadsOutput,
} from "./setup";
import type {
  SubmitFailuresAndCompletionOutput,
  SubmissionExportsOutput,
} from "./submission";

export interface CorrectionInput {
  app: ReturnType<typeof createApp>;
  completedSubmissionId: SubmitFailuresAndCompletionOutput["completedSubmissionId"];
  responseId: DraftLifecycleOutput["responseId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  responseDocumentKey: SubmitFailuresAndCompletionOutput["responseDocumentKey"];
  formRecord: PrefillEntryOutput["formRecord"];
  adminId: BootstrapAndCreationOutput["adminId"];
  formId: BootstrapAndCreationOutput["formId"];
  competingAdminBearer: TemplateUploadsOutput["competingAdminBearer"];
  savedDraftData: DraftLifecycleOutput["savedDraftData"];
  userBearer: PrefillEntryOutput["userBearer"];
  dataBody: SubmissionExportsOutput["dataBody"];
  otherUserBearer: SubmissionExportsOutput["otherUserBearer"];
  mock: PrefillEntryOutput["mock"];
  postSubmitExternalReference: SubmitFailuresAndCompletionOutput["postSubmitExternalReference"];
}

export interface CorrectionOutput {
  originalSubmissionBeforeCorrection: Pick<
    Submission,
    "data" | "objectKey" | "documentKey"
  >;
  correctionData: {
    description_1: string;
    accept_terms: boolean;
    department: string;
    description_2: string;
    full_name: string | boolean;
    start_date: string;
  };
  correction: Correction;
}

export const runCorrection = async (
  input: CorrectionInput
): Promise<CorrectionOutput> => {
  const {
    app,
    completedSubmissionId,
    responseId,
    adminBearer,
    responseDocumentKey,
    formRecord,
    adminId,
    formId,
    competingAdminBearer,
    savedDraftData,
    userBearer,
    dataBody,
    otherUserBearer,
    mock,
    postSubmitExternalReference,
  } = input;
  const originalSubmissionBeforeCorrection =
    await prisma.submission.findUniqueOrThrow({
      select: {
        data: true,
        documentKey: true,
        objectKey: true,
      },
      where: { id: completedSubmissionId },
    });
  const correctionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(correctionEditorResponse.status).toBe(200);
  const correctionEditorConfig =
    (await correctionEditorResponse.json()) as EditorConfigBody;
  const correctionDocumentKey = correctionEditorConfig.config.document.key;
  const correctionCapability =
    correctionEditorConfig.bridge.capabilities["save-correction"];
  if (!correctionCapability || !correctionDocumentKey) {
    throw new Error("The correction editor capabilities were not returned");
  }
  expect(correctionDocumentKey).not.toBe(responseDocumentKey);
  expect(
    correctionEditorConfig.config.editorConfig.plugins.options[pluginGuid]
  ).toMatchObject({
    bridgeId: correctionEditorConfig.bridge.id,
    publicId: formRecord.publicId,
  });
  const correctionClaims = verifyEditorCapability(correctionCapability);
  expect(correctionClaims).toMatchObject({
    action: "save-correction",
    actorId: adminId,
    documentKey: correctionDocumentKey,
    formId,
    leaseId: correctionEditorConfig.bridge.lease.id,
    role: "admin",
    targetId: responseId,
    targetType: "correction",
  });
  const competingCorrectionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${competingAdminBearer}` } }
    )
  );
  expect(competingCorrectionEditorResponse.status).toBe(409);
  expect(await competingCorrectionEditorResponse.json()).toMatchObject({
    error: "editor_in_use",
  });
  const correctionData = {
    ...savedDraftData,
    description_1: "แก้ไขข้อมูลโดยผู้ดูแล",
  };
  const correctionInputData = { ...correctionData, department: "Engineering" };
  const correctionSaveResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction`,
      {
        body: JSON.stringify({
          data: correctionInputData,
          documentKey: correctionDocumentKey,
          reason: "แก้ไขตามเอกสารต้นฉบับ",
        }),
        headers: capabilityHeaders(correctionCapability),
        method: "POST",
      }
    )
  );
  expect(correctionSaveResponse.status).toBe(202);
  const correctionSaveBody = (await correctionSaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !correctionSaveBody.operationCapability ||
    !correctionSaveBody.operationId
  ) {
    throw new Error("The correction operation was not created");
  }
  const correctionOperation = await waitForOperation(
    app,
    correctionSaveBody.operationId,
    { "X-Editor-Capability": correctionSaveBody.operationCapability }
  );
  const correctionOperationRecord = await prisma.operation.findUniqueOrThrow({
    select: { result: true },
    where: { id: correctionSaveBody.operationId },
  });
  const correctionResult = correctionOperationRecord.result as
    | { correctionId?: string; revision?: number }
    | undefined;
  expect(correctionOperation).toMatchObject({
    result: { correctionId: expect.any(String), revision: 1 },
    status: "completed",
  });
  if (!correctionResult?.correctionId) {
    throw new Error("The correction result was not returned");
  }
  expect(correctionResult.revision).toBe(1);
  const correctedSubmission = await prisma.submission.findUniqueOrThrow({
    select: { data: true, documentKey: true, objectKey: true },
    where: { id: completedSubmissionId },
  });
  expect(correctedSubmission).toEqual(originalSubmissionBeforeCorrection);
  const correction = await prisma.correction.findFirstOrThrow({
    where: { responseId, revision: 1 },
  });
  expect(String(correctionResult.correctionId)).toBe(correction.id);
  expect(correction).toMatchObject({
    actorId: adminId,
    data: correctionData,
    reason: "แก้ไขตามเอกสารต้นฉบับ",
    responseId,
    revision: 1,
    submissionId: completedSubmissionId,
  });
  expect(await objectExists(correction.objectKey)).toBe(true);
  const correctionReplayClaim = await prisma.callbackClaim.findUnique({
    where: { operationId: correctionSaveBody.operationId },
  });
  if (!correctionReplayClaim) {
    throw new Error("The correction callback claim was not persisted");
  }
  const correctionReplayUserdata = createCallbackUserdata({
    documentKey: correctionDocumentKey,
    expiresAt: Math.floor(correctionReplayClaim.expiresAt.getTime() / 1000),
    operationId: correctionSaveBody.operationId,
    operationType: "save_correction",
  });
  expect(
    createHash("sha256").update(correctionReplayUserdata).digest("hex")
  ).toBe(correctionReplayClaim.tokenDigest);
  const correctionReplayPayload = {
    key: correctionDocumentKey,
    status: 6,
    userdata: correctionReplayUserdata,
  };
  const correctionReplayResponse = await app.handle(
    new Request("http://test.local/onlyoffice/callback", {
      body: JSON.stringify({
        ...correctionReplayPayload,
        token: createOnlyOfficeBodyToken(correctionReplayPayload),
      }),
      headers: {
        Authorization: createOnlyOfficeAuthorization(correctionReplayPayload),
        ...jsonHeaders,
      },
      method: "POST",
    })
  );
  expect(correctionReplayResponse.status).toBe(200);
  expect(await correctionReplayResponse.json()).toEqual({ error: 0 });
  expect(await prisma.correction.count({ where: { responseId } })).toBe(1);
  const correctionHistoryResponse = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/corrections`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(correctionHistoryResponse.status).toBe(200);
  expect(await correctionHistoryResponse.json()).toMatchObject({
    latestRevision: 1,
    revisions: [
      {
        data: savedDraftData,
        reason: null,
        revision: 0,
      },
      {
        data: correctionData,
        reason: "แก้ไขตามเอกสารต้นฉบับ",
        revision: 1,
      },
    ],
  });
  const latestDataResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestDataResponse.status).toBe(200);
  const latestDataBody =
    (await latestDataResponse.json()) as typeof dataBody & {
      correction: { reason: string; revision: number } | null;
      revision: number;
    };
  expect(latestDataBody).toMatchObject({
    correction: { reason: "แก้ไขตามเอกสารต้นฉบับ", revision: 1 },
    data: correctionData,
    fields: dataBody.fields,
    revision: 1,
  });
  const latestJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestJsonResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}-revision-1.json"`
  );
  expect(JSON.parse(await latestJsonResponse.text())).toEqual(correctionData);
  const originalJsonAfterCorrection = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json?revision=original`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(JSON.parse(await originalJsonAfterCorrection.text())).toEqual(
    savedDraftData
  );
  const latestDocxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestDocxResponse.status).toBe(200);
  expect(latestDocxResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}-revision-1.docx"`
  );
  const latestPdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestPdfResponse.status).toBe(200);
  expect(latestPdfResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}-revision-1.pdf"`
  );
  expect(await latestPdfResponse.text()).toBe("%PDF-test");
  const correctedResultsResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}?revision=latest`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(correctedResultsResponse.status).toBe(200);
  expect(await correctedResultsResponse.json()).toMatchObject({
    result: {
      correction: { reason: "แก้ไขตามเอกสารต้นฉบับ", revision: 1 },
      data: correctionData,
      document: { available: true, state: "correction" },
      latestCorrectionNumber: 1,
      revision: 1,
    },
  });
  const originalAdminResultResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}?revision=original`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(await originalAdminResultResponse.json()).toMatchObject({
    result: { data: savedDraftData, revision: 0 },
  });
  const forbiddenAdminResultResponse = await app.handle(
    new Request(`http://test.local/api/admin/results/${responseId}`, {
      headers: { Authorization: `Bearer ${otherUserBearer}` },
    })
  );
  expect(forbiddenAdminResultResponse.status).toBe(403);
  const forbiddenViewerConfigResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenViewerConfigResponse.status).toBe(403);
  const invalidViewerRevisionResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config?revision=bogus`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(invalidViewerRevisionResponse.status).toBe(400);
  const latestViewerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(latestViewerResponse.status).toBe(200);
  const latestViewerConfig = (await latestViewerResponse.json()) as {
    bridge?: unknown;
    config: {
      document: {
        fileType: string;
        key: string;
        permissions: Record<string, unknown>;
        url: string;
      };
      editorConfig: Record<string, unknown>;
      token: string;
    };
  };
  expect(latestViewerConfig).not.toHaveProperty("bridge");
  expect(latestViewerConfig).not.toHaveProperty("capabilities");
  expect(latestViewerConfig.config.document).toMatchObject({
    fileType: "docx",
    key: correction.documentKey,
    permissions: {
      comment: false,
      download: false,
      edit: false,
      fillForms: false,
      review: false,
    },
  });
  expect(latestViewerConfig.config.editorConfig).toMatchObject({
    mode: "view",
  });
  expect(latestViewerConfig.config.editorConfig).not.toHaveProperty("plugins");
  expect(latestViewerConfig.config.editorConfig).not.toHaveProperty(
    "callbackUrl"
  );
  const [, viewerTokenPayload] = latestViewerConfig.config.token.split(".");
  if (!viewerTokenPayload) {
    throw new Error("The signed viewer configuration was not returned");
  }
  const signedViewerConfig = JSON.parse(
    Buffer.from(viewerTokenPayload, "base64url").toString("utf-8")
  ) as Record<string, unknown>;
  expect(signedViewerConfig).toMatchObject({
    document: {
      permissions: { download: false, edit: false, fillForms: false },
    },
    editorConfig: { mode: "view" },
  });
  const originalViewerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config?revision=original`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(originalViewerResponse.status).toBe(200);
  const originalViewerConfig =
    (await originalViewerResponse.json()) as typeof latestViewerConfig;
  expect(originalViewerConfig.config.document.key).toBe(
    originalSubmissionBeforeCorrection.documentKey
  );
  const unauthorizedOfficeDocumentResponse = await app.handle(
    new Request(
      `http://test.local/onlyoffice/document/${encodeURIComponent(
        correction.documentKey
      )}`
    )
  );
  expect(unauthorizedOfficeDocumentResponse.status).toBe(401);
  const viewerAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "view_correction",
      actorId: adminId,
      targetId: correction.id,
    },
  });
  expect(viewerAudit).toMatchObject({
    outcome: "success",
    safeMetadata: { revision: 1, state: "submitted" },
  });
  const correctionAudits = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: {
      actorId: adminId,
      targetId: { in: [responseId, correction.id, completedSubmissionId] },
    },
  });
  for (const audit of correctionAudits) {
    expect(JSON.stringify(audit.safeMetadata)).not.toContain(
      "แก้ไขตามเอกสารต้นฉบับ"
    );
    expect(JSON.stringify(audit.safeMetadata)).not.toContain(
      correction.objectKey
    );
  }
  const externalStatusAfterCorrection = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(await externalStatusAfterCorrection.json()).toMatchObject({
    latestCorrectionNumber: 1,
    status: "submitted",
  });
  return {
    correction,
    correctionData,
    originalSubmissionBeforeCorrection,
  };
};

export interface CorrectionFailureAndRevisionsInput {
  app: ReturnType<typeof createApp>;
  convertedDocumentKeys: string[];
  responseId: DraftLifecycleOutput["responseId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  correctionData: CorrectionOutput["correctionData"];
  completedSubmissionId: SubmitFailuresAndCompletionOutput["completedSubmissionId"];
  adminId: BootstrapAndCreationOutput["adminId"];
  userBearer: PrefillEntryOutput["userBearer"];
  savedDraftData: DraftLifecycleOutput["savedDraftData"];
  submissionDocument: SubmissionExportsOutput["submissionDocument"];
  originalSubmissionBeforeCorrection: CorrectionOutput["originalSubmissionBeforeCorrection"];
  correction: CorrectionOutput["correction"];
  otherUserBearer: SubmissionExportsOutput["otherUserBearer"];
  userId: PrefillEntryOutput["userId"];
}

export const runCorrectionFailureAndRevisions = async (
  input: CorrectionFailureAndRevisionsInput
): Promise<void> => {
  const {
    app,
    convertedDocumentKeys,
    responseId,
    adminBearer,
    correctionData,
    completedSubmissionId,
    adminId,
    userBearer,
    savedDraftData,
    submissionDocument,
    originalSubmissionBeforeCorrection,
    correction,
    otherUserBearer,
    userId,
  } = input;
  const failedCorrectionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(failedCorrectionEditorResponse.status).toBe(200);
  const failedCorrectionEditorConfig =
    (await failedCorrectionEditorResponse.json()) as EditorConfigBody;
  const failedCorrectionCapability =
    failedCorrectionEditorConfig.bridge.capabilities["save-correction"];
  if (!failedCorrectionCapability) {
    throw new Error("The failed correction capability was not returned");
  }
  const correctionFailureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.resolve(new TextEncoder().encode("%PDF-test")),
      forceSave: () =>
        Promise.reject(
          new Error("deterministic correction force-save failure")
        ),
    },
  });
  const failedCorrectionSaveResponse = await correctionFailureApp.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction`,
      {
        body: JSON.stringify({
          data: { description_1: "ควรไม่ถูกบันทึก" },
          documentKey: failedCorrectionEditorConfig.config.document.key,
          reason: "การแก้ไขที่ล้มเหลว",
        }),
        headers: capabilityHeaders(failedCorrectionCapability),
        method: "POST",
      }
    )
  );
  expect(failedCorrectionSaveResponse.status).toBe(202);
  const failedCorrectionSaveBody =
    (await failedCorrectionSaveResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !failedCorrectionSaveBody.operationCapability ||
    !failedCorrectionSaveBody.operationId
  ) {
    throw new Error("The failed correction operation was not created");
  }
  const failedCorrectionOperation = await waitForOperation(
    app,
    failedCorrectionSaveBody.operationId,
    { "X-Editor-Capability": failedCorrectionSaveBody.operationCapability }
  );
  expect(failedCorrectionOperation).toMatchObject({
    error: "force_save_failed",
    status: "failed",
  });
  expect(await prisma.correction.count({ where: { responseId } })).toBe(1);
  expect(
    await prisma.correction.findFirstOrThrow({
      orderBy: { revision: "desc" },
      where: { responseId },
    })
  ).toMatchObject({ data: correctionData, revision: 1 });
  const failedCorrectionReleaseResponse = await app.handle(
    new Request(
      `http://test.local/api/editor-leases/${failedCorrectionEditorConfig.bridge.lease.id}`,
      { headers: { Authorization: `Bearer ${adminBearer}` }, method: "DELETE" }
    )
  );
  expect(failedCorrectionReleaseResponse.status).toBe(200);
  const latestCorrectionData = {
    ...correctionData,
    description_1: "แก้ไขครั้งล่าสุด",
  };
  const latestCorrectionDocumentKey = crypto.randomUUID();
  const latestCorrectionObjectKey = objectKey(
    "submissions",
    completedSubmissionId,
    "revision-2",
    crypto.randomUUID(),
    "docx"
  );
  await putObject(
    latestCorrectionObjectKey,
    docxFixture("Ticket 04 latest correction"),
    DOCX_CONTENT_TYPE
  );
  await prisma.correction.create({
    data: {
      actorId: adminId,
      changedData: { description_1: "แก้ไขครั้งล่าสุด" },
      data: latestCorrectionData,
      documentKey: latestCorrectionDocumentKey,
      objectKey: latestCorrectionObjectKey,
      reason: "Correction ล่าสุด",
      responseId,
      revision: 2,
      submissionId: completedSubmissionId,
    },
  });
  const allRevisionHistoryResponse = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/corrections`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(allRevisionHistoryResponse.status).toBe(200);
  expect(await allRevisionHistoryResponse.json()).toMatchObject({
    latestRevision: 2,
    revisions: [
      { data: savedDraftData, revision: 0 },
      { data: correctionData, revision: 1 },
      { data: latestCorrectionData, revision: 2 },
    ],
  });
  const implicitLatestDataResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(implicitLatestDataResponse.status).toBe(200);
  expect(await implicitLatestDataResponse.json()).toMatchObject({
    correction: { revision: 2 },
    data: latestCorrectionData,
    revision: 2,
  });
  const implicitLatestDocxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(implicitLatestDocxResponse.status).toBe(200);
  expect(
    new Uint8Array(await implicitLatestDocxResponse.arrayBuffer())
  ).toEqual(Uint8Array.from(await readObject(latestCorrectionObjectKey)));
  const implicitLatestPdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(implicitLatestPdfResponse.status).toBe(200);
  expect(convertedDocumentKeys.at(-1)).toBe(latestCorrectionDocumentKey);
  const revisionArtifacts = [
    {
      data: savedDraftData,
      document: submissionDocument,
      documentKey: originalSubmissionBeforeCorrection.documentKey,
      revision: 0,
    },
    {
      data: correctionData,
      document: await readObject(correction.objectKey),
      documentKey: correction.documentKey,
      revision: 1,
    },
    {
      data: latestCorrectionData,
      document: await readObject(latestCorrectionObjectKey),
      documentKey: latestCorrectionDocumentKey,
      revision: 2,
    },
  ];
  for (const revision of revisionArtifacts) {
    const revisionQuery = `?revision=${revision.revision}`;
    const revisionDataResponse = await app.handle(
      new Request(
        `http://test.local/api/submissions/${completedSubmissionId}/data${revisionQuery}`,
        { headers: { Authorization: `Bearer ${userBearer}` } }
      )
    );
    expect(revisionDataResponse.status).toBe(200);
    expect(await revisionDataResponse.json()).toMatchObject({
      data: revision.data,
      revision: revision.revision,
    });
    const revisionDocxResponse = await app.handle(
      new Request(
        `http://test.local/api/submissions/${completedSubmissionId}/docx${revisionQuery}`,
        { headers: { Authorization: `Bearer ${userBearer}` } }
      )
    );
    expect(revisionDocxResponse.status).toBe(200);
    expect(new Uint8Array(await revisionDocxResponse.arrayBuffer())).toEqual(
      Uint8Array.from(revision.document)
    );
    const revisionPdfResponse = await app.handle(
      new Request(
        `http://test.local/api/submissions/${completedSubmissionId}/pdf${revisionQuery}`,
        { headers: { Authorization: `Bearer ${userBearer}` } }
      )
    );
    expect(revisionPdfResponse.status).toBe(200);
    expect(revisionPdfResponse.headers.get("content-disposition")).toContain(
      revision.revision === 0 ? ".pdf" : `-revision-${revision.revision}.pdf`
    );
    expect(convertedDocumentKeys.at(-1)).toBe(revision.documentKey);
    const adminRevisionQuery = `?revision=${revision.revision}`;
    const adminResultResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/results/${responseId}${adminRevisionQuery}`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    expect(adminResultResponse.status).toBe(200);
    expect(await adminResultResponse.json()).toMatchObject({
      result: {
        correction:
          revision.revision === 0
            ? null
            : {
                reason:
                  revision.revision === 1
                    ? "แก้ไขตามเอกสารต้นฉบับ"
                    : "Correction ล่าสุด",
                revision: revision.revision,
              },
        data: revision.data,
        document: {
          available: true,
          state: revision.revision === 0 ? "submission" : "correction",
        },
        latestCorrectionNumber: 2,
        revision: revision.revision,
      },
    });
    const adminViewerResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/results/${responseId}/viewer-config${adminRevisionQuery}`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    expect(adminViewerResponse.status).toBe(200);
    const adminViewerBody = (await adminViewerResponse.json()) as {
      config: { document: { key: string } };
    };
    expect(adminViewerBody.config.document.key).toBe(revision.documentKey);
  }
  const implicitAdminResultResponse = await app.handle(
    new Request(`http://test.local/api/admin/results/${responseId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(await implicitAdminResultResponse.json()).toMatchObject({
    result: {
      data: latestCorrectionData,
      latestCorrectionNumber: 2,
      revision: 2,
    },
  });
  const implicitAdminViewerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  const implicitAdminViewerBody =
    (await implicitAdminViewerResponse.json()) as {
      config: { document: { key: string } };
    };
  expect(implicitAdminViewerBody.config.document.key).toBe(
    latestCorrectionDocumentKey
  );
  const missingAdminResultRevisionResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}?revision=3`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(missingAdminResultRevisionResponse.status).toBe(404);
  const missingAdminViewerRevisionResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config?revision=3`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(missingAdminViewerRevisionResponse.status).toBe(404);
  const missingRevisionResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data?revision=3`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(missingRevisionResponse.status).toBe(404);
  const otherOwnerRevisionResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data?revision=1`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(otherOwnerRevisionResponse.status).toBe(403);
  const exportAudits = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: {
      action: "export_response",
      actorId: adminId,
      targetId: completedSubmissionId,
      targetType: "submission",
    },
  });
  expect(exportAudits.map((audit) => audit.safeMetadata)).toEqual([
    { format: "json", revision: 0, state: "submitted" },
    { format: "docx", revision: 0, state: "submitted" },
    { format: "pdf", revision: 0, state: "submitted" },
  ]);
  const stableSubmissionBeforeConversion = await prisma.submission.findUnique({
    select: {
      data: true,
      documentKey: true,
      objectKey: true,
    },
    where: { id: completedSubmissionId },
  });
  const conversionFailureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.reject(new Error("deterministic conversion failure")),
      forceSave: () => Promise.resolve(false),
    },
  });
  const conversionFailureResponse = await conversionFailureApp.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf?revision=original`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(conversionFailureResponse.status).toBe(500);
  const failedPdfAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "export_response",
      actorId: userId,
      outcome: "failure",
      targetId: completedSubmissionId,
      targetType: "submission",
    },
  });
  expect(failedPdfAudit.safeMetadata).toEqual({
    errorCode: "pdf_conversion_failed",
    format: "pdf",
    revision: 0,
    state: "submitted",
  });
  expect(
    await prisma.submission.findUnique({
      select: {
        data: true,
        documentKey: true,
        objectKey: true,
      },
      where: { id: completedSubmissionId },
    })
  ).toEqual(stableSubmissionBeforeConversion);
  if (!stableSubmissionBeforeConversion) {
    throw new Error("The stable submission was not found");
  }
  expect(await readObject(stableSubmissionBeforeConversion.objectKey)).toEqual(
    submissionDocument
  );
};
