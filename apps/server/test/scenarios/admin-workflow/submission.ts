// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";
import { createHash } from "node:crypto";

import type { Handoff } from "@onlyoffice/db";
import { prisma } from "@onlyoffice/db";

import { createApp } from "../../../src/app";
import { readObject, DOCX_CONTENT_TYPE } from "../../../src/storage";
import type { EditorConfigBody } from "../../fixtures/http";
import {
  jsonHeaders,
  createCredentialFixture,
  bearerFor,
  waitForOperation,
} from "../../fixtures/http";
import type { DraftLifecycleOutput } from "./draft";
import {
  capabilityHeaders,
  draftRequest,
  refreshResponseEditor,
  submitRequest,
} from "./helpers";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";
import type { PrimaryPublicationOutput } from "./template-contract";

export interface SubmitFailuresAndCompletionInput {
  app: ReturnType<typeof createApp>;
  responseId: DraftLifecycleOutput["responseId"];
  formRecord: PrefillEntryOutput["formRecord"];
  savedDraftData: DraftLifecycleOutput["savedDraftData"];
  responseDocumentKey: DraftLifecycleOutput["responseDocumentKey"];
  saveDraftCapability: DraftLifecycleOutput["saveDraftCapability"];
  userBearer: PrefillEntryOutput["userBearer"];
  submitCapability: DraftLifecycleOutput["submitCapability"];
  formId: BootstrapAndCreationOutput["formId"];
  userId: PrefillEntryOutput["userId"];
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  mock: PrefillEntryOutput["mock"];
  userEmail: BootstrapAndCreationOutput["userEmail"];
  handoffCandidateValues: PrefillEntryOutput["handoffCandidateValues"];
  saveBody: DraftLifecycleOutput["saveBody"];
}

export interface SubmitFailuresAndCompletionOutput {
  completedSubmissionId: string;
  postSubmitExternalReference: string;
  postSubmitHandoffRecord: Handoff;
  responseDocumentKey: string;
  saveDraftCapability: string;
  submitCapability: string;
  userLease: {
    expiresAt: string;
    id: string;
    releaseUrl: string;
    renewUrl: string;
  };
}

export const runSubmitFailuresAndCompletion = async (
  input: SubmitFailuresAndCompletionInput
): Promise<SubmitFailuresAndCompletionOutput> => {
  const {
    app,
    responseId,
    formRecord,
    savedDraftData,
    userBearer,
    formId,
    userId,
    formPublicId,
    adminBearer,
    mock,
    userEmail,
    handoffCandidateValues,
    saveBody,
  } = input;
  let { responseDocumentKey, saveDraftCapability, submitCapability } = input;

  const stableBeforeFailure = await prisma.response.findUnique({
    select: { draftData: true, draftObjectKey: true },
    where: { id: responseId },
  });
  const failureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.reject(new Error("deterministic draft failure")),
      forceSave: () => Promise.reject(new Error("deterministic draft failure")),
    },
  });
  const failedSaveResponse = await failureApp.handle(
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
  expect(failedSaveResponse.status).toBe(202);
  const failedSaveBody = (await failedSaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!failedSaveBody.operationCapability || !failedSaveBody.operationId) {
    throw new Error("The failed draft operation was not created");
  }
  const failedSaveOperation = await waitForOperation(
    app,
    failedSaveBody.operationId,
    { "X-Editor-Capability": failedSaveBody.operationCapability }
  );
  expect(failedSaveOperation).toMatchObject({
    error: "force_save_failed",
    status: "failed",
  });
  expect(
    await prisma.response.findUnique({
      select: { draftData: true, draftObjectKey: true },
      where: { id: responseId },
    })
  ).toEqual(stableBeforeFailure);
  const retryDocumentKeyBeforeSave = responseDocumentKey;
  const retrySaveResponse = await draftRequest(
    app,
    formRecord.publicId,
    responseDocumentKey,
    responseId,
    saveDraftCapability,
    savedDraftData
  );
  expect(retrySaveResponse.status).toBe(202);
  const retrySaveBody = (await retrySaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!retrySaveBody.operationCapability || !retrySaveBody.operationId) {
    throw new Error("The retry draft operation was not created");
  }
  const retrySaveOperation = await waitForOperation(
    app,
    retrySaveBody.operationId,
    {
      "X-Editor-Capability": retrySaveBody.operationCapability,
    }
  );
  expect(retrySaveOperation.status).toBe("completed");
  const refreshedEditor = await refreshResponseEditor(
    app,
    formRecord.publicId,
    responseId,
    userBearer
  );
  ({ responseDocumentKey, saveDraftCapability, submitCapability } =
    refreshedEditor);
  const userLease: EditorConfigBody["bridge"]["lease"] =
    refreshedEditor.userLease;
  expect(responseDocumentKey).not.toBe(retryDocumentKeyBeforeSave);
  const stableBeforeSubmitFailure = await prisma.response.findUnique({
    select: { draftData: true, draftObjectKey: true, status: true },
    where: { id: responseId },
  });
  const failedSubmitResponse = await submitRequest(
    app,
    formRecord.publicId,
    responseDocumentKey,
    responseId,
    submitCapability,
    savedDraftData,
    failureApp
  );
  expect(failedSubmitResponse.status).toBe(202);
  const failedSubmitBody = (await failedSubmitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!failedSubmitBody.operationCapability || !failedSubmitBody.operationId) {
    throw new Error("The failed submit operation was not created");
  }
  const failedSubmitOperation = await waitForOperation(
    app,
    failedSubmitBody.operationId,
    { "X-Editor-Capability": failedSubmitBody.operationCapability }
  );
  expect(failedSubmitOperation).toMatchObject({
    error: "force_save_failed",
    status: "failed",
  });
  expect(
    await prisma.response.findUnique({
      select: { draftData: true, draftObjectKey: true, status: true },
      where: { id: responseId },
    })
  ).toEqual(stableBeforeSubmitFailure);
  expect(await prisma.submission.count({ where: { responseId } })).toBe(0);

  const concurrentSubmitResponses = await Promise.all([
    submitRequest(
      app,
      formRecord.publicId,
      responseDocumentKey,
      responseId,
      submitCapability,
      savedDraftData
    ),
    submitRequest(
      app,
      formRecord.publicId,
      responseDocumentKey,
      responseId,
      submitCapability,
      savedDraftData
    ),
  ]);
  expect(
    concurrentSubmitResponses.filter((response) => response.status === 202)
  ).toHaveLength(1);
  const rejectedSubmitResponses = concurrentSubmitResponses.filter(
    (response) => response.status === 409
  );
  expect(rejectedSubmitResponses).toHaveLength(1);
  expect(await rejectedSubmitResponses[0]?.json()).toMatchObject({
    error: "operation_in_progress",
  });
  const submitResponse = concurrentSubmitResponses.find(
    (response) => response.status === 202
  );
  if (!submitResponse) {
    throw new Error("The concurrent submit operation was not created");
  }
  const submitBody = (await submitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
    submissionId?: string;
  };
  const completedSubmissionId = submitBody.submissionId;
  if (
    !submitBody.operationCapability ||
    !submitBody.operationId ||
    !completedSubmissionId
  ) {
    throw new Error("The submit operation was not created");
  }
  const submitOperation = await waitForOperation(app, submitBody.operationId, {
    "X-Editor-Capability": submitBody.operationCapability,
  });
  expect(submitOperation.status).toBe("completed");
  const repeatStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(repeatStartResponse.status).toBe(200);
  expect(await repeatStartResponse.json()).toMatchObject({
    receiptUrl: `/receipt/${completedSubmissionId}`,
    response: {
      id: responseId,
      status: "submitted",
      submissionId: completedSubmissionId,
    },
    submissionId: completedSubmissionId,
  });
  expect(
    await prisma.response.count({
      where: { formId, userId },
    })
  ).toBe(1);
  const postSubmitUnarchive = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({ status: "published" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(postSubmitUnarchive.status).toBe(200);
  const postSubmitUnarchiveBody = (await postSubmitUnarchive.json()) as {
    form?: { publicId?: string; status?: string };
  };
  expect(postSubmitUnarchiveBody.form).toMatchObject({
    publicId: formPublicId,
    status: "published",
  });
  const postSubmitExternalReference = `ticket-17-post-submit-${crypto.randomUUID()}`;
  const postSubmitHandoffCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: postSubmitExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(postSubmitHandoffCreate.status).toBe(200);
  const postSubmitCode = (
    (await postSubmitHandoffCreate.json()) as {
      code: string;
    }
  ).code;
  if (!postSubmitCode) {
    throw new Error("The post-submit handoff code was not created");
  }
  const postSubmitLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: postSubmitCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(postSubmitLaunch.status).toBe(303);
  expect(postSubmitLaunch.headers.get("location")).toBe(
    `/forms/${formRecord.publicId}/fill`
  );
  const postSubmitCookie = postSubmitLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!postSubmitCookie) {
    throw new Error("The post-submit handoff did not set a cookie");
  }
  const postSubmitStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: postSubmitCookie,
      },
      method: "POST",
    })
  );
  expect(postSubmitStart.status).toBe(200);
  expect(await postSubmitStart.json()).toMatchObject({
    receiptUrl: `/receipt/${completedSubmissionId}`,
    response: { id: responseId, status: "submitted" },
    submissionId: completedSubmissionId,
  });
  expect(postSubmitStart.headers.get("set-cookie")).toContain("Max-Age=0");
  const postSubmitHandoffRecord = await prisma.handoff.findFirstOrThrow({
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(postSubmitExternalReference)
        .digest("hex"),
    },
  });
  expect(postSubmitHandoffRecord).toMatchObject({
    responseId,
    status: "consumed",
  });
  expect(
    await prisma.response.count({
      where: { formId, userId },
    })
  ).toBe(1);
  const replayPostSubmitStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: postSubmitCookie,
      },
      method: "POST",
    })
  );
  expect(replayPostSubmitStart.status).toBe(409);
  expect(await replayPostSubmitStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "redeem_handoff",
        outcome: "failure",
        targetId: formPublicId,
      },
    })
  ).toMatchObject({ targetId: formPublicId });
  const submittedStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(submittedStatusResponse.status).toBe(200);
  expect(await submittedStatusResponse.json()).toMatchObject({
    latestCorrectionNumber: null,
    status: "submitted",
    submittedAt: expect.any(String),
  });
  const immutableDraftResponse = await draftRequest(
    app,
    formRecord.publicId,
    responseDocumentKey,
    responseId,
    saveDraftCapability,
    savedDraftData
  );
  expect(immutableDraftResponse.status).toBe(409);
  expect(await immutableDraftResponse.json()).toMatchObject({
    error: "stale_response",
  });
  const submittedCountListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(submittedCountListResponse.status).toBe(200);
  const submittedCountList = (await submittedCountListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      submissionCount: number;
    }[];
  };
  expect(
    submittedCountList.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({ activeDraftCount: 0, submissionCount: 1 });
  const ownerOperationVisibilityResponse = await app.handle(
    new Request(`http://test.local/api/operations/${submitBody.operationId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(ownerOperationVisibilityResponse.status).toBe(200);
  expect(await ownerOperationVisibilityResponse.json()).toMatchObject({
    operation: { id: submitBody.operationId, status: "completed" },
  });
  const adminOperationVisibilityResponse = await app.handle(
    new Request(`http://test.local/api/operations/${submitBody.operationId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminOperationVisibilityResponse.status).toBe(200);
  expect(await adminOperationVisibilityResponse.json()).toMatchObject({
    operation: { id: submitBody.operationId, status: "completed" },
  });
  const crossOperationPollResponse = await app.handle(
    new Request(`http://test.local/api/operations/${submitBody.operationId}`, {
      headers: {
        "X-Editor-Capability": saveBody.operationCapability,
      },
    })
  );
  expect(crossOperationPollResponse.status).toBe(403);
  return {
    completedSubmissionId,
    postSubmitExternalReference,
    postSubmitHandoffRecord,
    responseDocumentKey,
    saveDraftCapability,
    submitCapability,
    userLease,
  };
};

export interface SubmissionExportsInput {
  app: ReturnType<typeof createApp>;
  otherUserEmail: BootstrapAndCreationOutput["otherUserEmail"];
  password: BootstrapAndCreationOutput["password"];
  completedSubmissionId: SubmitFailuresAndCompletionOutput["completedSubmissionId"];
  publishedManifest: PrimaryPublicationOutput["publishedManifest"];
  userBearer: PrefillEntryOutput["userBearer"];
  savedDraftData: DraftLifecycleOutput["savedDraftData"];
  responseId: DraftLifecycleOutput["responseId"];
  publishedManifestRecord: PrimaryPublicationOutput["publishedManifestRecord"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  formId: BootstrapAndCreationOutput["formId"];
  formRecord: PrefillEntryOutput["formRecord"];
}

export interface SubmissionExportsOutput {
  otherUser: { email: string; id: string };
  otherUserBearer: string;
  dataBody: {
    data: Record<string, unknown>;
    fields: {
      label: string;
      options: { displayText: string; value: string }[];
      placeholder: string | null;
      position: number;
      tag: string;
      type: string;
    }[];
    returnUrl: string;
    submission: Record<string, unknown>;
  };
  submissionDocument: Uint8Array<ArrayBuffer>;
  templateDraft: { id: string; objectKey: string; documentKey: string };
  sourceDocument: Uint8Array<ArrayBufferLike>;
}

export const runSubmissionExports = async (
  input: SubmissionExportsInput
): Promise<SubmissionExportsOutput> => {
  const {
    app,
    otherUserEmail,
    password,
    completedSubmissionId,
    publishedManifest,
    userBearer,
    savedDraftData,
    responseId,
    publishedManifestRecord,
    adminBearer,
    formId,
    formRecord,
  } = input;
  const otherUser = await createCredentialFixture({
    email: otherUserEmail,
    name: "Ticket 04 Other User",
    password,
  });
  const otherUserBearer = await bearerFor(app, otherUserEmail, password);
  const forbiddenDocxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenDocxResponse.status).toBe(403);
  const forbiddenPdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenPdfResponse.status).toBe(403);
  const forbiddenDataResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenDataResponse.status).toBe(403);

  await prisma.$transaction(async (tx) => {
    await tx.$executeRaw`ALTER TABLE "field_manifests" DISABLE TRIGGER "field_manifests_immutable"`;
    await tx.$executeRaw`ALTER TABLE "manifest_fields" DISABLE TRIGGER "manifest_fields_immutable"`;
    try {
      await tx.$executeRaw`
        UPDATE "manifest_fields"
        SET "position" = "position" + 100000
        WHERE "manifest_id" = ${publishedManifest.id}::uuid
      `;
      await tx.$executeRaw`
        WITH ranked_fields AS (
          SELECT
            "id",
            (ROW_NUMBER() OVER (ORDER BY "tag") - 1)::integer AS "position"
          FROM "manifest_fields"
          WHERE "manifest_id" = ${publishedManifest.id}::uuid
        )
        UPDATE "manifest_fields" AS field
        SET
          "label" = field."tag",
          "placeholder" = NULL,
          "position" = ranked_fields."position"
        FROM ranked_fields
        WHERE field."id" = ranked_fields."id"
      `;
      await tx.fieldManifest.update({
        data: { displayMetadataVersion: 0 },
        where: { id: publishedManifest.id },
      });
    } finally {
      await tx.$executeRaw`ALTER TABLE "manifest_fields" ENABLE TRIGGER "manifest_fields_immutable"`;
      await tx.$executeRaw`ALTER TABLE "field_manifests" ENABLE TRIGGER "field_manifests_immutable"`;
    }
  });
  const dataResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(dataResponse.status).toBe(200);
  const dataBody = (await dataResponse.json()) as {
    data: Record<string, unknown>;
    fields: {
      label: string;
      options: { displayText: string; value: string }[];
      placeholder: string | null;
      position: number;
      tag: string;
      type: string;
    }[];
    returnUrl: string;
    submission: Record<string, unknown>;
  };
  expect(dataBody).toMatchObject({
    data: savedDraftData,
    returnUrl: "https://source.example.test/forms/return",
    submission: { id: completedSubmissionId, responseId },
  });
  expect(dataBody.submission).not.toHaveProperty("userEmail");
  expect(
    dataBody.fields.map(({ label, placeholder, position, tag, type }) => ({
      label,
      placeholder,
      position,
      tag,
      type,
    }))
  ).toEqual(
    [...publishedManifestRecord.manifest.fields]
      .toSorted((left, right) => left.position - right.position)
      .map(({ label, placeholder, position, tag, type }) => ({
        label,
        placeholder,
        position,
        tag,
        type,
      }))
  );
  expect(
    dataBody.fields.find((field) => field.tag === "department")?.options
  ).toEqual([
    { displayText: "Choose an item", value: "" },
    { displayText: "Engineering", value: "engineering" },
    { displayText: "Human Resources", value: "hr" },
    { displayText: "Finance", value: "finance" },
  ]);
  await prisma.$transaction(async (tx) => {
    await tx.$executeRaw`ALTER TABLE "field_manifests" DISABLE TRIGGER "field_manifests_immutable"`;
    await tx.$executeRaw`ALTER TABLE "manifest_fields" DISABLE TRIGGER "manifest_fields_immutable"`;
    try {
      await tx.$executeRaw`
        UPDATE "manifest_fields"
        SET "position" = "position" + 100000
        WHERE "manifest_id" = ${publishedManifest.id}::uuid
      `;
      for (const field of dataBody.fields) {
        await tx.$executeRaw`
          UPDATE "manifest_fields"
          SET
            "label" = ${field.label},
            "placeholder" = ${field.placeholder},
            "position" = ${field.position}
          WHERE "manifest_id" = ${publishedManifest.id}::uuid AND "tag" = ${field.tag}
        `;
      }
      await tx.fieldManifest.update({
        data: {
          displayMetadataVersion: publishedManifest.displayMetadataVersion,
        },
        where: { id: publishedManifest.id },
      });
    } finally {
      await tx.$executeRaw`ALTER TABLE "manifest_fields" ENABLE TRIGGER "manifest_fields_immutable"`;
      await tx.$executeRaw`ALTER TABLE "field_manifests" ENABLE TRIGGER "field_manifests_immutable"`;
    }
  });
  const ownerJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(ownerJsonResponse.status).toBe(200);
  expect(ownerJsonResponse.headers.get("content-type")).toBe(
    "application/json; charset=utf-8"
  );
  expect(ownerJsonResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.json"`
  );
  expect(JSON.parse(await ownerJsonResponse.text())).toEqual(savedDraftData);
  const adminJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminJsonResponse.status).toBe(200);
  expect(adminJsonResponse.headers.get("content-type")).toBe(
    "application/json; charset=utf-8"
  );
  expect(adminJsonResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.json"`
  );
  expect(JSON.parse(await adminJsonResponse.text())).toEqual(savedDraftData);
  const forbiddenJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenJsonResponse.status).toBe(403);

  const docxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(docxResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.docx"`
  );
  expect(docxResponse.status).toBe(200);
  expect(docxResponse.headers.get("content-type")).toBe(DOCX_CONTENT_TYPE);
  const submissionDocument = new Uint8Array(await docxResponse.arrayBuffer());
  const templateDraft = await prisma.templateDraft.findUnique({
    select: { documentKey: true, id: true, objectKey: true },
    where: { formId },
  });
  if (!templateDraft) {
    throw new Error("The test template draft was not found");
  }
  const sourceDocument = await readObject(templateDraft.objectKey);
  const adminDocxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminDocxResponse.status).toBe(200);
  expect(adminDocxResponse.headers.get("content-type")).toBe(DOCX_CONTENT_TYPE);
  expect(adminDocxResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.docx"`
  );
  await adminDocxResponse.arrayBuffer();

  const pdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(pdfResponse.status).toBe(200);
  expect(pdfResponse.headers.get("content-type")).toBe("application/pdf");
  expect(pdfResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.pdf"`
  );
  expect(await pdfResponse.text()).toBe("%PDF-test");
  const adminPdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminPdfResponse.status).toBe(200);
  expect(adminPdfResponse.headers.get("content-type")).toBe("application/pdf");
  expect(adminPdfResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.pdf"`
  );
  expect(await adminPdfResponse.text()).toBe("%PDF-test");
  const adminSubmittedResultsResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&state=submitted`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminSubmittedResultsResponse.status).toBe(200);
  expect(await adminSubmittedResultsResponse.json()).toMatchObject({
    results: [
      {
        formPublicId: formRecord.publicId,
        id: responseId,
        latestCorrectionNumber: null,
        state: "submitted",
        submissionId: completedSubmissionId,
      },
    ],
  });
  const adminSubmittedDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/results/${responseId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminSubmittedDetailResponse.status).toBe(200);
  expect(await adminSubmittedDetailResponse.json()).toMatchObject({
    result: {
      data: savedDraftData,
      document: { available: true, state: "submission" },
      id: responseId,
      state: "submitted",
      submissionId: completedSubmissionId,
    },
  });
  return {
    dataBody,
    otherUser,
    otherUserBearer,
    sourceDocument,
    submissionDocument,
    templateDraft,
  };
};
