// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

import { readObject } from "../../../src/storage";
import { waitForOperation } from "../../fixtures/http";
import {
  patchFillMethod,
  getNativeConfig,
  nativeSubmitRequest,
  onlyOfficeDraftRequest,
  refreshOnlyOfficeEditorConfig,
} from "./helpers";
import type { ScenarioApp } from "./helpers";
import type { OfficeCanonicalValuesOutput } from "./office-values";
import type { NativeSetupOutput } from "./setup";

export interface NativeResubmitInput {
  app: ScenarioApp;
  adminBearer: string;
  publicId: string;
  responseId: string;
  userBearer: string;
  editorConfigUrl: string;
  prefillLocks: NativeSetupOutput["prefillLocks"];
  onlyOfficeCanonicalData: OfficeCanonicalValuesOutput["onlyOfficeCanonicalData"];
}

export const runNativeResubmit = async (
  input: NativeResubmitInput
): Promise<void> => {
  const {
    app,
    adminBearer,
    publicId,
    responseId,
    userBearer,
    editorConfigUrl,
    prefillLocks,
    onlyOfficeCanonicalData,
  } = input;
  let onlyOfficeSaveCapability: string;
  let onlyOfficeDocumentKey: string;

  const saveOnlyOfficeDisplayDraft = async (data: Record<string, unknown>) => {
    const response = await onlyOfficeDraftRequest(
      app,
      publicId,
      responseId,
      onlyOfficeDocumentKey,
      onlyOfficeSaveCapability,
      data
    );
    expect(response.status).toBe(202);
    const operation = (await response.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
    if (!operation.operationCapability || !operation.operationId) {
      throw new Error("The display draft operation was not created");
    }
    expect(
      await waitForOperation(app, operation.operationId, {
        "X-Editor-Capability": operation.operationCapability,
      })
    ).toMatchObject({ status: "completed" });
    return prisma.response.findUniqueOrThrow({
      select: { draftData: true },
      where: { id: responseId },
    });
  };
  for (const { data, expected } of [
    {
      data: { category: "Empty option label" },
      expected: { category: "" },
    },
    {
      data: { category: "Choice B" },
      expected: { category: "Empty option label" },
    },
    {
      data: { category: "Friendly label" },
      expected: { category: "stored_value" },
    },
    {
      data: { custom_category: "" },
      expected: { custom_category: " " },
    },
    {
      data: { custom_category: "Empty combo label" },
      expected: { custom_category: "" },
    },
    {
      data: { custom_category: "Choice B" },
      expected: { custom_category: "Empty combo label" },
    },
    {
      data: { custom_category: "custom-browser-value" },
      expected: { custom_category: "custom-browser-value" },
    },
    {
      data: { custom_category: "Whitespace combo label" },
      expected: { custom_category: " " },
    },
  ]) {
    ({ onlyOfficeSaveCapability, onlyOfficeDocumentKey } =
      await refreshOnlyOfficeEditorConfig(app, publicId, userBearer));
    const displayDraft = await saveOnlyOfficeDisplayDraft(data);
    expect(displayDraft.draftData).toMatchObject(expected);
  }
  ({ onlyOfficeSaveCapability, onlyOfficeDocumentKey } =
    await refreshOnlyOfficeEditorConfig(app, publicId, userBearer));
  const explicitlyClearedDraft = await saveOnlyOfficeDisplayDraft({
    category: null,
    enabled: false,
    start_date: null,
  });
  expect(explicitlyClearedDraft.draftData).toMatchObject({
    category: null,
    enabled: false,
    start_date: null,
  });
  ({ onlyOfficeSaveCapability, onlyOfficeDocumentKey } =
    await refreshOnlyOfficeEditorConfig(app, publicId, userBearer));
  const restoredDraft = await saveOnlyOfficeDisplayDraft({
    category: "Friendly label",
    enabled: "☒",
    start_date: "Thursday, February 29, 2024 d literal",
  });
  expect(restoredDraft.draftData).toMatchObject({
    category: "stored_value",
    enabled: true,
    start_date: "2024-02-29",
  });

  const nativeAgainResponse = await patchFillMethod(
    app,
    publicId,
    adminBearer,
    "native"
  );
  expect(nativeAgainResponse.status).toBe(200);
  const nativeAgainStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  const nativeAgainStart = (await nativeAgainStartResponse.json()) as {
    editorConfigUrl?: string;
    fillMethod?: string;
  };
  expect(nativeAgainStart.fillMethod).toBe("native");
  const nativeConfig = await getNativeConfig(app, editorConfigUrl, userBearer);
  expect(nativeConfig).toMatchObject({
    data: {
      category: "stored_value",
      cleared_category: null,
      cleared_date: null,
      comments: "Saved\nnative\tanswer",

      custom_category: " ",
      empty_category: " ",
      enabled: true,
      full_name: "Trusted\nPrefill",
      nullable_category: "",
      start_date: "2024-02-29",
      state_checkbox: true,
      value_checkbox: false,
    },
    lockedFields: prefillLocks,
  });
  const submitResponse = await nativeSubmitRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    {
      ...onlyOfficeCanonicalData,
      empty_category: "",
    }
  );
  expect(submitResponse.status).toBe(202);
  const submitBody = (await submitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
    submissionId?: string;
  };
  if (!submitBody.operationCapability || !submitBody.operationId) {
    throw new Error("The Ticket 06 submit operation was not created");
  }
  expect(
    await waitForOperation(app, submitBody.operationId, {
      "X-Editor-Capability": submitBody.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const submission = await prisma.submission.findUniqueOrThrow({
    select: { data: true, objectKey: true, responseId: true },
    where: { id: submitBody.submissionId },
  });
  expect(submission).toMatchObject({
    data: {
      category: "stored_value",
      cleared_category: null,
      cleared_date: null,
      comments: "Saved\nnative\tanswer",

      custom_category: " ",
      empty_category: "",
      full_name: "Trusted\nPrefill",
      nullable_category: "",
      start_date: "2024-02-29",
      state_checkbox: true,
      value_checkbox: false,
    },
    responseId,
  });
  const submissionXml = new TextDecoder().decode(
    unzipSync(await readObject(submission.objectKey))["word/document.xml"]
  );
  expect(submissionXml).toContain("<w:rPr><w:b/></w:rPr><w:t>Trusted</w:t>");
  expect(submissionXml).toContain(
    '<w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Prefill</w:t>'
  );
  const submissionCommentsContent = submissionXml.match(
    /<w:tag w:val="comments"\/>[\s\S]*?<w:sdtContent>(?<content>[\s\S]*?)<\/w:sdtContent>/u
  )?.groups?.content;
  expect(submissionCommentsContent).toContain("<w:br/>");
  expect(submissionCommentsContent).toContain("<w:tab/>");
  expect(submissionXml).toContain("Whitespace combo label");
  const submittedResponse = await prisma.response.findUniqueOrThrow({
    select: { status: true },
    where: { id: responseId },
  });
  expect(submittedResponse.status).toBe("submitted");
  const submittedStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(await submittedStartResponse.json()).toMatchObject({
    submissionId: submitBody.submissionId,
  });
};
