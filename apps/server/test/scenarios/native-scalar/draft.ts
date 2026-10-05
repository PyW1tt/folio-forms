import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { createApp } from "../../../src/app";
import { readObject } from "../../../src/storage";
import { waitForOperation } from "../../fixtures/http";
import {
  getNativeConfig,
  nativeDraftRequest,
  nativeSubmitRequest,
} from "./helpers";
import type { ScenarioApp, NativeResponseConfig } from "./helpers";
import type { NativeSetupOutput } from "./setup";

export interface NativeDraftInput {
  app: ScenarioApp;
  editorConfigUrl: string;
  userBearer: string;
  publicId: string;
  responseId: string;
  prefillValues: NativeSetupOutput["prefillValues"];
  prefillLocks: NativeSetupOutput["prefillLocks"];
}

export interface NativeDraftOutput {
  nativeConfig: NativeResponseConfig;
  draftData: {
    category: string;
    cleared_category: null;
    cleared_date: null;
    comments: string;
    custom_category: string;
    empty_category: string;
    enabled: boolean;
    full_name: string;
    nullable_category: null;
    start_date: string;
    state_checkbox: boolean;
    value_checkbox: boolean;
  };
  savedDraftDocumentKey: string;
}

export const runNativeDraft = async (
  input: NativeDraftInput
): Promise<NativeDraftOutput> => {
  const {
    app,
    editorConfigUrl,
    userBearer,
    publicId,
    responseId,
    prefillValues,
    prefillLocks,
  } = input;

  let nativeConfig = await getNativeConfig(app, editorConfigUrl, userBearer);
  expect(nativeConfig).toMatchObject({
    data: prefillValues,
    fillMethod: "native",
    lockedFields: prefillLocks,
    responseId,
  });
  expect(
    nativeConfig.fields.map(
      ({ label, options, placeholder, position, required, tag, type }) => ({
        label,
        options,
        placeholder,
        position,
        required,
        tag,
        type,
      })
    )
  ).toEqual([
    {
      label: "Full name",
      options: [],
      placeholder: "Enter full name",
      position: 0,
      required: true,
      tag: "full_name",
      type: "text",
    },
    {
      label: "Comments",
      options: [],
      placeholder: "Add comments",
      position: 1,
      required: false,
      tag: "comments",
      type: "text",
    },
    {
      label: "Enabled",
      options: [],
      placeholder: null,
      position: 2,
      required: true,
      tag: "enabled",
      type: "checkbox",
    },
    {
      label: "Checkbox with state children",
      options: [],
      placeholder: null,
      position: 3,
      required: false,
      tag: "state_checkbox",
      type: "checkbox",
    },
    {
      label: "Checkbox with unchecked marker value",
      options: [],
      placeholder: null,
      position: 4,
      required: false,
      tag: "value_checkbox",
      type: "checkbox",
    },
    {
      label: "Start date",
      options: [],
      placeholder: null,
      position: 5,
      required: false,
      tag: "start_date",
      type: "date",
    },
    {
      label: "Cleared date",
      options: [],
      placeholder: null,
      position: 6,
      required: false,
      tag: "cleared_date",
      type: "date",
    },
    {
      label: "Category",
      options: [
        { displayText: "Friendly label", value: "stored_value" },
        { displayText: "Empty option label", value: "" },
        { displayText: "Choice B", value: "Empty option label" },
      ],
      placeholder: null,
      position: 7,
      required: false,
      tag: "category",
      type: "dropdown",
    },
    {
      label: "Empty category",
      options: [
        { displayText: "Empty option label", value: "" },
        { displayText: "Whitespace option label", value: " " },
      ],
      placeholder: null,
      position: 8,
      required: true,
      tag: "empty_category",
      type: "dropdown",
    },
    {
      label: "Nullable category",
      options: [{ displayText: "Nullable empty option label", value: "" }],
      placeholder: null,
      position: 9,
      required: false,
      tag: "nullable_category",
      type: "dropdown",
    },
    {
      label: "Cleared category",
      options: [{ displayText: "Old label", value: "old_value" }],
      placeholder: null,
      position: 10,
      required: false,
      tag: "cleared_category",
      type: "dropdown",
    },
    {
      label: "Custom category",
      options: [
        { displayText: "Suggested label", value: "suggested" },
        { displayText: "Empty combo label", value: "" },
        { displayText: "Choice B", value: "Empty combo label" },
        { displayText: "Whitespace combo label", value: " " },
      ],
      placeholder: null,
      position: 11,
      required: true,
      tag: "custom_category",
      type: "combo",
    },
  ]);
  const operationsBeforeInvalidInput = await prisma.operation.count({
    where: { responseId },
  });
  const invalidDraftResponse = await nativeDraftRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    {
      comments: "x".repeat(10_001),
    }
  );
  expect(invalidDraftResponse.status).toBe(422);
  expect(await invalidDraftResponse.json()).toMatchObject({
    error: "invalid_response_data",
  });
  for (const invalidValue of [
    { start_date: "2023-02-29" },
    { category: "unknown" },
    { custom_category: "x".repeat(10_001) },
    { full_name: "Tampered locked Prefill" },
  ]) {
    const response = await nativeDraftRequest(
      app,
      publicId,
      responseId,
      nativeConfig,
      invalidValue
    );
    expect(response.status).toBe(422);
    expect(await response.json()).toMatchObject({
      error: "invalid_response_data",
    });
  }
  await prisma.prefillSnapshot.update({
    data: { lockedFields: {}, values: {} },
    where: { responseId },
  });
  const invalidRequiredSubmit = await nativeSubmitRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    {
      comments: "Missing required name",
    }
  );
  expect(invalidRequiredSubmit.status).toBe(422);
  expect(await invalidRequiredSubmit.json()).toMatchObject({
    error: "invalid_response_data",
  });
  const uncheckedRequiredSubmit = await nativeSubmitRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    {
      enabled: false,
      full_name: "Present",
    }
  );
  expect(uncheckedRequiredSubmit.status).toBe(422);
  expect(await uncheckedRequiredSubmit.json()).toMatchObject({
    error: "invalid_response_data",
  });
  const nullRequiredDropdownSubmit = await nativeSubmitRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    {
      empty_category: null,
      enabled: true,
      full_name: "Present",
    }
  );
  expect(nullRequiredDropdownSubmit.status).toBe(422);
  expect(await nullRequiredDropdownSubmit.json()).toMatchObject({
    error: "invalid_response_data",
  });
  const unselectedRequiredComboSubmit = await nativeSubmitRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    {
      custom_category: null,
      empty_category: "",
      enabled: true,
      full_name: "Present",
    }
  );
  expect(unselectedRequiredComboSubmit.status).toBe(422);
  expect(await unselectedRequiredComboSubmit.json()).toMatchObject({
    error: "invalid_response_data",
  });
  const unlistedWhitespaceComboSubmit = await nativeSubmitRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    {
      custom_category: "  ",
      empty_category: "",
      enabled: true,
      full_name: "Present",
    }
  );
  expect(unlistedWhitespaceComboSubmit.status).toBe(422);
  expect(await unlistedWhitespaceComboSubmit.json()).toMatchObject({
    error: "invalid_response_data",
  });
  expect(await prisma.operation.count({ where: { responseId } })).toBe(
    operationsBeforeInvalidInput
  );
  await prisma.prefillSnapshot.update({
    data: { lockedFields: prefillLocks, values: prefillValues },
    where: { responseId },
  });
  const draftData = {
    category: "stored_value",
    cleared_category: null,
    cleared_date: null,
    comments: "Saved native answer",
    custom_category: "suggested",
    empty_category: "",
    enabled: true,
    full_name: "Trusted\nPrefill",
    nullable_category: null,
    start_date: "2024-02-29",
    state_checkbox: true,
    value_checkbox: false,
  };
  const failedStorageApp = createApp({
    legacySso: null,
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.resolve(new TextEncoder().encode("%PDF-native")),
      forceSave: () => Promise.resolve(false),
    },
    prefillReturnUrl: "https://source.example.test/forms/return",
    putObject: () => Promise.reject(new Error("Ticket 06 storage failure")),
  });
  const stableBeforeFailure = await prisma.response.findUniqueOrThrow({
    select: {
      draftData: true,
      draftDocumentKey: true,
      draftObjectKey: true,
      status: true,
    },
    where: { id: responseId },
  });
  const stableDraftObjectKey = stableBeforeFailure.draftObjectKey;
  if (!stableDraftObjectKey) {
    throw new Error("The stable Ticket 06 draft has no object key");
  }
  const stableDocument = await readObject(stableDraftObjectKey);
  const failedSaveResponse = await nativeDraftRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    draftData,
    failedStorageApp
  );
  expect(failedSaveResponse.status).toBe(202);
  const failedSaveBody = (await failedSaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!failedSaveBody.operationCapability || !failedSaveBody.operationId) {
    throw new Error("The Ticket 06 failed save operation was not created");
  }
  expect(
    await waitForOperation(app, failedSaveBody.operationId, {
      "X-Editor-Capability": failedSaveBody.operationCapability,
    })
  ).toMatchObject({ error: "document_save_failed", status: "failed" });
  expect(
    await prisma.response.findUniqueOrThrow({
      select: {
        draftData: true,
        draftDocumentKey: true,
        draftObjectKey: true,
        status: true,
      },
      where: { id: responseId },
    })
  ).toEqual(stableBeforeFailure);
  expect(await readObject(stableDraftObjectKey)).toEqual(stableDocument);
  const saveResponse = await nativeDraftRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    draftData
  );
  expect(saveResponse.status).toBe(202);
  const saveBody = (await saveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!saveBody.operationCapability || !saveBody.operationId) {
    throw new Error("The Ticket 06 save operation was not created");
  }
  expect(
    await waitForOperation(app, saveBody.operationId, {
      "X-Editor-Capability": saveBody.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const savedResponse = await prisma.response.findUniqueOrThrow({
    select: { draftData: true, draftDocumentKey: true, draftObjectKey: true },
    where: { id: responseId },
  });
  const savedDraftDocumentKey = savedResponse.draftDocumentKey;
  if (!savedDraftDocumentKey) {
    throw new Error("The saved Ticket 06 draft has no document key");
  }
  expect(savedResponse.draftData).toEqual({
    category: "stored_value",
    cleared_category: null,
    cleared_date: null,
    comments: "Saved native answer",
    custom_category: "suggested",
    empty_category: "",
    enabled: true,
    full_name: "Trusted\nPrefill",
    nullable_category: null,
    start_date: "2024-02-29",
    state_checkbox: true,
    value_checkbox: false,
  });
  const savedDraftObjectKey = savedResponse.draftObjectKey;
  if (!savedDraftObjectKey) {
    throw new Error("The saved Ticket 06 draft has no object key");
  }
  const savedArchive = unzipSync(await readObject(savedDraftObjectKey));
  const savedDocumentXml = new TextDecoder().decode(
    savedArchive["word/document.xml"]
  );
  expect(savedDocumentXml).toContain("Static layout");
  expect(savedDocumentXml).not.toContain("showingPlcHdr");
  const nullableCategoryContent = savedDocumentXml.match(
    /<w:tag w:val="nullable_category"\/>[\s\S]*?<w:sdtContent>(?<content>[\s\S]*?)<\/w:sdtContent>/u
  )?.[1];
  expect(nullableCategoryContent).toBe("<w:r><w:t></w:t></w:r>");
  expect(savedDocumentXml).toContain(
    '<w:tag xmlns:ext="urn:fixture" ext:val="metadata" w:val="full_name"/>'
  );
  expect(savedDocumentXml).toContain("<w:rPr><w:b/></w:rPr><w:t>Trusted</w:t>");
  expect(savedDocumentXml).toContain(
    '<w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Prefill</w:t>'
  );
  expect(savedDocumentXml).toContain("Saved native answer");
  expect(savedDocumentXml).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>');
  expect(savedDocumentXml).toContain(
    '<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>'
  );
  expect(savedDocumentXml).toContain(
    'xmlns:branch="http://schemas.microsoft.com/office/word/2010/wordml"'
  );
  expect(savedDocumentXml).toContain("☒");
  expect(savedDocumentXml.match(/<w14:checked w14:val="1"\/>/gu)).toHaveLength(
    1
  );
  expect(
    savedDocumentXml.match(/<branch:checked branch:val="1"\/>/gu)
  ).toHaveLength(1);
  expect(savedDocumentXml).toContain(
    '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2713" w14:font="Ticket Symbols"/>'
  );
  expect(savedDocumentXml).toContain(
    '<branch:checkbox><branch:checked branch:val="1"/><branch:checkedState branch:val="2713" branch:font="Ticket Symbols"/>'
  );
  expect(savedDocumentXml).toContain("✓");
  expect(savedDocumentXml).toContain("□");
  expect(
    savedDocumentXml.match(
      /<w:rFonts w:ascii="Ticket Symbols" w:hAnsi="Ticket Symbols"\/>/gu
    )
  ).toHaveLength(2);
  expect(savedDocumentXml).toContain('w:fullDate="2024-02-29T00:00:00Z"');
  expect(savedDocumentXml).toContain("Thursday, February 29, 2024");
  expect(savedDocumentXml).toContain("Thursday, February 29, 2024 d literal");
  expect(savedDocumentXml).not.toContain('w:fullDate="1999-12-31T00:00:00Z"');
  expect(savedDocumentXml).toContain('w:lastValue="stored_value"');
  expect(savedDocumentXml).toContain('w:lastValue=""');
  expect(savedDocumentXml).not.toContain('w:lastValue="old_value"');
  expect(savedDocumentXml).toContain("Empty option label");
  expect(savedDocumentXml).toContain("Friendly label");
  expect(savedDocumentXml).toContain("Suggested label");
  expect(savedDocumentXml).toContain('w:value="stored_value"');
  let pdfConversions = 0;
  const pdfApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () => {
        pdfConversions += 1;
        return Promise.resolve(new TextEncoder().encode("%PDF-native"));
      },
      forceSave: () => Promise.resolve(false),
    },
    prefillReturnUrl: "https://source.example.test/forms/return",
  });
  expect(pdfConversions).toBe(0);
  const pdfResponse = await pdfApp.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/pdf`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(pdfResponse.status).toBe(200);
  expect(await pdfResponse.text()).toBe("%PDF-native");
  expect(pdfConversions).toBe(1);
  const resumeResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  const resumeBody = (await resumeResponse.json()) as {
    editorConfigUrl?: string;
    fillMethod?: string;
    response?: { id?: string };
  };
  expect(resumeBody).toMatchObject({
    fillMethod: "native",
    response: { id: responseId },
  });
  if (!resumeBody.editorConfigUrl) {
    throw new Error("The saved Ticket 06 response did not reopen");
  }
  nativeConfig = await getNativeConfig(app, editorConfigUrl, userBearer);
  expect(nativeConfig).toMatchObject({
    data: {
      category: "stored_value",
      cleared_category: null,
      cleared_date: null,
      comments: "Saved native answer",
      custom_category: "suggested",
      empty_category: "",
      enabled: true,
      full_name: "Trusted\nPrefill",
      start_date: "2024-02-29",
      state_checkbox: true,
      value_checkbox: false,
    },
    lockedFields: prefillLocks,
  });
  return { draftData, nativeConfig, savedDraftDocumentKey };
};
