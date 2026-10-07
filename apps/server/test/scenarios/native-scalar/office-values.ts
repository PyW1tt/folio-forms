// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync, strToU8, zipSync } from "fflate";

import {
  pluginGuid,
  createDocumentAccessToken,
  createOnlyOfficeAuthorization,
} from "../../../src/onlyoffice";
import { readObject, putObject, DOCX_CONTENT_TYPE } from "../../../src/storage";
import {
  waitForOperation,
  onlyOfficeSaveCapabilityFor,
} from "../../fixtures/http";
import type { NativeDraftOutput } from "./draft";
import {
  capabilityHeaders,
  patchFillMethod,
  nativeDraftRequest,
  onlyOfficeDraftRequest,
  refreshOnlyOfficeEditorConfig,
} from "./helpers";
import type {
  ScenarioApp,
  NativeResponseConfig,
  OnlyOfficeResponseConfig,
} from "./helpers";
import type { NativeSetupOutput } from "./setup";

export interface OfficeCanonicalValuesInput {
  app: ScenarioApp;
  adminBearer: string;
  publicId: string;
  responseId: string;
  userBearer: string;
  nativeConfig: NativeResponseConfig;
  draftData: NativeDraftOutput["draftData"];
  savedDraftDocumentKey: string;
  prefillValues: NativeSetupOutput["prefillValues"];
}

export interface OfficeCanonicalValuesOutput {
  onlyOfficeCanonicalData: {
    category: string;
    cleared_category: null;
    cleared_date: null;
    comments: string;
    custom_category: string;
    empty_category: string;
    enabled: boolean;
    full_name: string;
    nullable_category: string;
    start_date: string;
    state_checkbox: boolean;
    value_checkbox: boolean;
  };
}

export const runOfficeCanonicalValues = async (
  input: OfficeCanonicalValuesInput
): Promise<OfficeCanonicalValuesOutput> => {
  const {
    app,
    adminBearer,
    publicId,
    responseId,
    userBearer,
    nativeConfig,
    draftData,
    savedDraftDocumentKey,
    prefillValues,
  } = input;

  const onlyOfficeMethodResponse = await patchFillMethod(
    app,
    publicId,
    adminBearer,
    "onlyoffice"
  );
  expect(onlyOfficeMethodResponse.status).toBe(200);
  const staleNativeSave = await nativeDraftRequest(
    app,
    publicId,
    responseId,
    nativeConfig,
    draftData
  );
  expect(staleNativeSave.status).toBe(409);
  expect(await staleNativeSave.json()).toMatchObject({
    error: "fill_method_changed",
  });
  const onlyOfficeStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  const onlyOfficeStart = (await onlyOfficeStartResponse.json()) as {
    editorConfigUrl?: string;
    fillMethod?: string;
  };
  expect(onlyOfficeStart.fillMethod).toBe("onlyoffice");
  if (!onlyOfficeStart.editorConfigUrl) {
    throw new Error("The switched Ticket 06 response did not reopen");
  }
  const onlyOfficeConfigResponse = await app.handle(
    new Request(
      new URL(onlyOfficeStart.editorConfigUrl, "http://test.local").toString(),
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  const onlyOfficeConfig =
    (await onlyOfficeConfigResponse.json()) as OnlyOfficeResponseConfig;
  expect(onlyOfficeConfig.fillMethod).toBe("onlyoffice");
  expect(onlyOfficeConfig.config?.document?.key).toBe(savedDraftDocumentKey);
  expect(
    onlyOfficeConfig.config?.editorConfig?.plugins?.options?.[pluginGuid]
      ?.prefill
  ).toMatchObject({
    data: prefillValues,
    editableFields: { full_name: false },
  });
  expect(
    onlyOfficeConfig.config?.editorConfig?.plugins?.options?.[pluginGuid]
      ?.tagAliases
  ).toEqual({ metadata: "full_name" });
  const onlyOfficeSubmitCapability =
    onlyOfficeConfig.bridge?.capabilities?.submit;
  if (!onlyOfficeSubmitCapability) {
    throw new Error("The ONLYOFFICE submit capability is missing");
  }
  let onlyOfficeSaveCapability = onlyOfficeSaveCapabilityFor(onlyOfficeConfig);
  let onlyOfficeDocumentKey = onlyOfficeConfig.config?.document?.key;
  if (!onlyOfficeDocumentKey) {
    throw new Error("The switched ONLYOFFICE document key is missing");
  }
  const onlyOfficeSubmitRequest = (data: Record<string, unknown>) =>
    app.handle(
      new Request(`http://test.local/api/forms/${publicId}/submit`, {
        body: JSON.stringify({
          data,
          documentKey: onlyOfficeDocumentKey,
          responseId,
        }),
        headers: capabilityHeaders(onlyOfficeSubmitCapability),
        method: "POST",
      })
    );
  const operationsBeforeInvalidOnlyOfficeSubmit = await prisma.operation.count({
    where: { responseId },
  });
  const unlistedWhitespaceOnlyOfficeSubmit = await onlyOfficeSubmitRequest({
    custom_category: "  ",
    empty_category: "",
    enabled: true,
    full_name: "Trusted\nPrefill",
  });
  expect(unlistedWhitespaceOnlyOfficeSubmit.status).toBe(422);
  expect(await unlistedWhitespaceOnlyOfficeSubmit.json()).toMatchObject({
    error: "invalid_response_data",
  });
  expect(await prisma.operation.count({ where: { responseId } })).toBe(
    operationsBeforeInvalidOnlyOfficeSubmit
  );
  const onlyOfficeOperationsBeforeInvalidInput = await prisma.operation.count({
    where: { responseId },
  });
  for (const invalidData of [
    { enabled: "invalid checkbox value" },
    { category: "Unknown option label" },
    { start_date: "Thursday, February 30, 2024 d literal" },
    { cleared_date: "31/12/99" },
    { comments: "x".repeat(10_001) },
    { custom_category: "x".repeat(10_001) },
    { full_name: "Tampered locked Prefill" },
  ]) {
    const invalidOnlyOfficeResponse = await onlyOfficeDraftRequest(
      app,
      publicId,
      responseId,
      onlyOfficeDocumentKey,
      onlyOfficeSaveCapability,
      invalidData
    );
    expect(invalidOnlyOfficeResponse.status).toBe(422);
    expect(await invalidOnlyOfficeResponse.json()).toMatchObject({
      error: "invalid_response_data",
    });
  }
  expect(await prisma.operation.count({ where: { responseId } })).toBe(
    onlyOfficeOperationsBeforeInvalidInput
  );
  const canonicalIsoDateResponse = await onlyOfficeDraftRequest(
    app,
    publicId,
    responseId,
    onlyOfficeDocumentKey,
    onlyOfficeSaveCapability,
    {
      cleared_date: "1999-12-31",
    }
  );
  expect(canonicalIsoDateResponse.status).toBe(202);
  const canonicalIsoDateOperation = (await canonicalIsoDateResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !canonicalIsoDateOperation.operationCapability ||
    !canonicalIsoDateOperation.operationId
  ) {
    throw new Error("The canonical ISO date draft operation was not created");
  }
  expect(
    await waitForOperation(app, canonicalIsoDateOperation.operationId, {
      "X-Editor-Capability": canonicalIsoDateOperation.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  ({ onlyOfficeSaveCapability, onlyOfficeDocumentKey } =
    await refreshOnlyOfficeEditorConfig(app, publicId, userBearer));
  const onlyOfficeDisplayData = {
    category: "Friendly label",
    cleared_category: "",
    cleared_date: "",
    comments: "",
    custom_category: "Whitespace combo label",
    empty_category: "Whitespace option label",
    enabled: "☒",
    full_name: "Trusted\nPrefill",
    nullable_category: "",
    start_date: "Thursday, February 29, 2024 d literal",
    value_checkbox: "□",
  };
  const callbackSource = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: responseId },
  });
  if (!callbackSource.draftObjectKey) {
    throw new Error("The ONLYOFFICE callback source document is missing");
  }
  const callbackArchive = unzipSync(
    await readObject(callbackSource.draftObjectKey)
  );
  const callbackXml = new TextDecoder().decode(
    callbackArchive["word/document.xml"]
  );
  if (!callbackXml.includes("</w:body>")) {
    throw new Error("The ONLYOFFICE callback source XML has no body");
  }
  callbackArchive["word/document.xml"] = strToU8(
    callbackXml
      .replace("<w:t>Trusted</w:t>", "<w:t>Wrong</w:t><w:tab/>")
      .replace(
        '<w:rPr><w:rFonts w:ascii="Ticket Symbols" w:hAnsi="Ticket Symbols"/></w:rPr>',
        '<w:rPr xmlns:fontAlias="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:b/><w:rFonts fontAlias:ascii="Old Ticket Font" fontAlias:hAnsi="Old Ticket Font" fontAlias:asciiTheme="majorAscii" fontAlias:hAnsiTheme="majorHAnsi"/></w:rPr>'
      )
      .replace(
        /(?<run><w:r><w:rPr xmlns:fontAlias="http:\/\/schemas\.openxmlformats\.org\/wordprocessingml\/2006\/main"><w:b\/><w:rFonts fontAlias:ascii="Old Ticket Font" fontAlias:hAnsi="Old Ticket Font" fontAlias:asciiTheme="majorAscii" fontAlias:hAnsiTheme="majorHAnsi"\/><\/w:rPr><w:t>[^<]*<\/w:t><\/w:r>)/u,
        '$1<w:r><w:rPr xmlns:fontAlias="urn:fixture"><w:rFonts fontAlias:asciiTheme="keep"/></w:rPr><w:t>extension marker</w:t></w:r>'
      )
      .replace(
        '<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>',
        '<w:rPr><w:i/><w:color w:val="00FF00"/></w:rPr>'
      )
      .replace("<w:t>Suggested label</w:t>", "<w:t>custom-browser-value</w:t>")
      .replace(
        "<w:t>Saved native answer</w:t>",
        "<w:t>Saved</w:t><w:br/><w:t>native</w:t><w:tab/><w:t>answer</w:t>"
      )
      .replace(
        "</w:body>",
        "<w:p><w:r><w:t>OnlyOffice-only edit</w:t></w:r></w:p></w:body>"
      )
  );

  await putObject(
    callbackSource.draftObjectKey,
    zipSync(callbackArchive),
    DOCX_CONTENT_TYPE
  );
  const onlyOfficeDraftResponse = await onlyOfficeDraftRequest(
    app,
    publicId,
    responseId,
    onlyOfficeDocumentKey,
    onlyOfficeSaveCapability,
    onlyOfficeDisplayData
  );
  expect(onlyOfficeDraftResponse.status).toBe(202);
  const onlyOfficeDraftBody = (await onlyOfficeDraftResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !onlyOfficeDraftBody.operationCapability ||
    !onlyOfficeDraftBody.operationId
  ) {
    throw new Error("The ONLYOFFICE draft operation was not created");
  }
  expect(
    await waitForOperation(app, onlyOfficeDraftBody.operationId, {
      "X-Editor-Capability": onlyOfficeDraftBody.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const onlyOfficeCanonicalData = {
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
  };
  const onlyOfficeDraft = await prisma.response.findUniqueOrThrow({
    select: { draftData: true, draftObjectKey: true },
    where: { id: responseId },
  });
  expect(onlyOfficeDraft.draftData).toMatchObject({
    custom_category: " ",
    empty_category: " ",
    nullable_category: null,
  });
  expect(onlyOfficeDraft.draftData).toMatchObject({
    comments: "Saved\nnative\tanswer",
  });
  if (!onlyOfficeDraft.draftObjectKey) {
    throw new Error("The ONLYOFFICE draft document was not stored");
  }
  const onlyOfficeDraftXml = new TextDecoder().decode(
    unzipSync(await readObject(onlyOfficeDraft.draftObjectKey))[
      "word/document.xml"
    ]
  );
  expect(onlyOfficeDraftXml).toContain("OnlyOffice-only edit");
  expect(onlyOfficeDraftXml).toContain("Friendly label");
  expect(onlyOfficeDraftXml).toContain("custom-browser-value");
  expect(onlyOfficeDraftXml).toContain("Whitespace option label");
  expect(onlyOfficeDraftXml).toContain("Whitespace combo label");
  expect(onlyOfficeDraftXml).toContain("☒");
  expect(onlyOfficeDraftXml).toContain("✓");
  expect(onlyOfficeDraftXml).toContain("□");
  expect(onlyOfficeDraftXml).toContain("Static layout");
  expect(onlyOfficeDraftXml).toContain(
    '<w:rPr><w:b/></w:rPr><w:t>Trust</w:t><w:t xml:space="preserve">ed</w:t>'
  );
  expect(onlyOfficeDraftXml).toContain(
    '<w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Prefill</w:t>'
  );
  expect(onlyOfficeDraftXml).not.toContain("Wrong");
  const onlyOfficeFullNameContent = onlyOfficeDraftXml.match(
    /<w:tag\b[^>]*\bw:val="full_name"\/>[\s\S]*?<w:sdtContent>(?<content>[\s\S]*?)<\/w:sdtContent>/u
  )?.groups?.content;
  expect(onlyOfficeFullNameContent).not.toContain("<w:tab/>");

  expect(onlyOfficeDraftXml).toContain(
    '<w:rPr xmlns:fontAlias="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:b/><w:rFonts fontAlias:ascii="Ticket Symbols" fontAlias:hAnsi="Ticket Symbols"/></w:rPr>'
  );
  expect(onlyOfficeDraftXml).not.toContain('fontAlias:asciiTheme="majorAscii"');
  expect(onlyOfficeDraftXml).not.toContain('fontAlias:hAnsiTheme="majorHAnsi"');
  expect(onlyOfficeDraftXml).toContain('fontAlias:asciiTheme="keep"');
  const onlyOfficeCommentsContent = onlyOfficeDraftXml.match(
    /<w:tag w:val="comments"\/>[\s\S]*?<w:sdtContent>(?<content>[\s\S]*?)<\/w:sdtContent>/u
  )?.groups?.content;
  expect(onlyOfficeCommentsContent).toContain("<w:br/>");
  expect(onlyOfficeCommentsContent).toContain("<w:tab/>");

  expect(onlyOfficeDraftXml).toContain(
    '<w:rPr><w:i/><w:color w:val="00FF00"/></w:rPr>'
  );
  expect(onlyOfficeDraftXml).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>');
  expect(onlyOfficeDraftXml).toContain(
    'xmlns:branch="http://schemas.microsoft.com/office/word/2010/wordml"'
  );
  expect(onlyOfficeDraftXml).toContain("Empty option label");
  expect(onlyOfficeDraftXml).toContain("Nullable empty option label");
  const callbackNullableCategoryContent = onlyOfficeDraftXml.match(
    /<w:tag w:val="nullable_category"\/>[\s\S]*?<w:sdtContent>(?<content>[\s\S]*?)<\/w:sdtContent>/u
  )?.[1];
  expect(callbackNullableCategoryContent).toBe("<w:r><w:t></w:t></w:r>");
  expect(onlyOfficeDraftXml).toContain('w:fullDate="2024-02-29T00:00:00Z"');
  expect(onlyOfficeDraftXml).not.toContain("AlternateContent");
  ({ onlyOfficeSaveCapability, onlyOfficeDocumentKey } =
    await refreshOnlyOfficeEditorConfig(app, publicId, userBearer));
  const emptyCheckboxResponse = await onlyOfficeDraftRequest(
    app,
    publicId,
    responseId,
    onlyOfficeDocumentKey,
    onlyOfficeSaveCapability,
    {
      category: "",
      custom_category: "",
      enabled: "",
      start_date: "",
    }
  );
  expect(emptyCheckboxResponse.status).toBe(202);
  const emptyCheckboxOperation = (await emptyCheckboxResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !emptyCheckboxOperation.operationCapability ||
    !emptyCheckboxOperation.operationId
  ) {
    throw new Error("The empty checkbox draft operation was not created");
  }
  expect(
    await waitForOperation(app, emptyCheckboxOperation.operationId, {
      "X-Editor-Capability": emptyCheckboxOperation.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const emptyScalarDraft = await prisma.response.findUniqueOrThrow({
    select: {
      draftData: true,
      draftDocumentKey: true,
      draftObjectKey: true,
    },
    where: { id: responseId },
  });
  const { draftObjectKey } = emptyScalarDraft;
  const { draftDocumentKey } = emptyScalarDraft;
  if (!draftObjectKey || !draftDocumentKey) {
    throw new Error("The saved scalar draft document is missing");
  }
  const emptyScalarBytes = await readObject(draftObjectKey);
  const emptyScalarArchive = unzipSync(emptyScalarBytes);
  const emptyScalarXml = new TextDecoder().decode(
    emptyScalarArchive["word/document.xml"]
  );
  expect(emptyScalarXml).toContain("<w14:checkbox");
  expect(emptyScalarXml).toContain("<w:date");
  expect(emptyScalarXml).toContain("<w:dropDownList");
  expect(emptyScalarXml).toContain("<w:comboBox");
  const nestedBlockControlXml =
    '<w:sdt><w:sdtPr><w:text/><w:tag w:val="nested_block"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Nested block</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  const pdfScalarFixtureWithNestedBlockXml = emptyScalarXml.replace(
    /(?<contentOpen><w:tag\b[^>]*w:val="enabled"\/>[\s\S]*?<w:sdtContent)(?<contentStart>>)[\s\S]*?(?<contentClose><\/w:sdtContent>)/u,
    `$<contentOpen> xmlns:pdfAlias="urn:ticket07:pdf"$<contentStart>${nestedBlockControlXml}$<contentClose>`
  );
  const customXmlBlockXml =
    "<w:customXml><w:p><w:r><w:t>Custom XML block</w:t></w:r></w:p></w:customXml>";
  const pdfScalarFixtureXml = pdfScalarFixtureWithNestedBlockXml.replace(
    /(?<contentOpen><w:tag\b[^>]*w:val="cleared_date"\/>[\s\S]*?<w:sdtContent>)[\s\S]*?(?<contentClose><\/w:sdtContent>)/u,
    `$<contentOpen>${customXmlBlockXml}$<contentClose>`
  );
  if (
    pdfScalarFixtureWithNestedBlockXml === emptyScalarXml ||
    pdfScalarFixtureXml === pdfScalarFixtureWithNestedBlockXml
  ) {
    throw new Error("The scalar block control fixtures are missing");
  }
  emptyScalarArchive["word/document.xml"] = strToU8(pdfScalarFixtureXml);
  await putObject(
    draftObjectKey,
    zipSync(emptyScalarArchive),
    DOCX_CONTENT_TYPE
  );
  try {
    const pdfDocumentUrl = new URL(
      `/onlyoffice/document/${encodeURIComponent(draftDocumentKey)}?token=${encodeURIComponent(createDocumentAccessToken(draftDocumentKey))}&pdf=1`,
      "http://test.local"
    );
    const pdfDocumentResponse = await app.handle(
      new Request(pdfDocumentUrl.toString(), {
        headers: {
          Authorization: createOnlyOfficeAuthorization({
            url: pdfDocumentUrl.toString(),
          }),
        },
      })
    );
    expect(pdfDocumentResponse.status).toBe(200);
    const pdfOnlyXml = new TextDecoder().decode(
      unzipSync(new Uint8Array(await pdfDocumentResponse.arrayBuffer()))[
        "word/document.xml"
      ]
    );
    expect(pdfOnlyXml).toContain("Thursday, February 29, 2024");
    expect(pdfOnlyXml).toContain("Friendly label");
    expect(pdfOnlyXml).toContain("custom-browser-value");
    expect(pdfOnlyXml).toContain("Nested block");
    expect(pdfOnlyXml).toContain("Custom XML block");
    expect(pdfOnlyXml).toContain('xmlns:pdfAlias="urn:ticket07:pdf"');
    expect(pdfOnlyXml).not.toContain("<w:p><w:sdt");
    expect(pdfOnlyXml).not.toContain("<w:p><w:customXml");
    const remainingPdfControlTags = [
      ...pdfOnlyXml.matchAll(/<w:tag\b[^>]*w:val="(?<tag>[^"]+)"/gu),
    ].map(({ groups }) => groups?.tag ?? "untagged");
    expect(remainingPdfControlTags).toEqual([
      "full_name",
      "comments",
      "nested_block",
    ]);
    expect(pdfOnlyXml).not.toContain("<w14:checkbox");
    expect(pdfOnlyXml).not.toContain("<w:date");
    expect(pdfOnlyXml).not.toContain("<w:dropDownList");
    expect(pdfOnlyXml).not.toContain("<w:comboBox");
  } finally {
    await putObject(draftObjectKey, emptyScalarBytes, DOCX_CONTENT_TYPE);
  }
  const unchangedDocxXml = new TextDecoder().decode(
    unzipSync(await readObject(draftObjectKey))["word/document.xml"]
  );
  expect(unchangedDocxXml).toContain("<w14:checkbox>");
  expect(unchangedDocxXml).toContain("Thursday, February 29, 2024");
  expect(emptyScalarXml).toContain("☒");
  expect(emptyScalarXml).toContain("Thursday, February 29, 2024");
  expect(emptyScalarXml).toContain("Friendly label");
  expect(emptyScalarXml).toContain("Whitespace combo label");
  expect(emptyScalarDraft.draftData).toMatchObject({
    category: "stored_value",
    custom_category: "custom-browser-value",
    enabled: true,
    start_date: "2024-02-29",
  });
  ({ onlyOfficeSaveCapability, onlyOfficeDocumentKey } =
    await refreshOnlyOfficeEditorConfig(app, publicId, userBearer));
  const checkedCheckboxResponse = await onlyOfficeDraftRequest(
    app,
    publicId,
    responseId,
    onlyOfficeDocumentKey,
    onlyOfficeSaveCapability,
    {
      enabled: true,
    }
  );
  expect(checkedCheckboxResponse.status).toBe(202);
  const checkedCheckboxOperation = (await checkedCheckboxResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !checkedCheckboxOperation.operationCapability ||
    !checkedCheckboxOperation.operationId
  ) {
    throw new Error("The checked checkbox draft operation was not created");
  }
  expect(
    await waitForOperation(app, checkedCheckboxOperation.operationId, {
      "X-Editor-Capability": checkedCheckboxOperation.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const checkedCheckboxDraft = await prisma.response.findUniqueOrThrow({
    select: { draftData: true },
    where: { id: responseId },
  });
  expect(checkedCheckboxDraft.draftData).toMatchObject({ enabled: true });
  ({ onlyOfficeSaveCapability, onlyOfficeDocumentKey } =
    await refreshOnlyOfficeEditorConfig(app, publicId, userBearer));
  const explicitEmptyOptionSource = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: responseId },
  });
  if (!explicitEmptyOptionSource.draftObjectKey) {
    throw new Error("The explicit empty option source document is missing");
  }
  const explicitEmptyOptionArchive = unzipSync(
    await readObject(explicitEmptyOptionSource.draftObjectKey)
  );
  const explicitEmptyOptionSourceXml = new TextDecoder().decode(
    explicitEmptyOptionArchive["word/document.xml"]
  );
  explicitEmptyOptionArchive["word/document.xml"] = strToU8(
    explicitEmptyOptionSourceXml.replace(
      /(?<opening><w:tag w:val="nullable_category"\/>[\s\S]*?<w:sdtContent>)[\s\S]*?(?<closing><\/w:sdtContent>)/u,
      "$1<w:r><w:t>Nullable empty option label</w:t></w:r>$2"
    )
  );
  await putObject(
    explicitEmptyOptionSource.draftObjectKey,
    zipSync(explicitEmptyOptionArchive),
    DOCX_CONTENT_TYPE
  );
  const explicitEmptyOptionResponse = await onlyOfficeDraftRequest(
    app,
    publicId,
    responseId,
    onlyOfficeDocumentKey,
    onlyOfficeSaveCapability,
    {
      ...onlyOfficeDisplayData,
      nullable_category: "Nullable empty option label",
    }
  );
  expect(explicitEmptyOptionResponse.status).toBe(202);
  const explicitEmptyOptionOperation =
    (await explicitEmptyOptionResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !explicitEmptyOptionOperation.operationCapability ||
    !explicitEmptyOptionOperation.operationId
  ) {
    throw new Error("The explicit empty option callback did not start");
  }
  expect(
    await waitForOperation(app, explicitEmptyOptionOperation.operationId, {
      "X-Editor-Capability": explicitEmptyOptionOperation.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const explicitEmptyOptionDraft = await prisma.response.findUniqueOrThrow({
    select: { draftData: true, draftObjectKey: true },
    where: { id: responseId },
  });
  expect(explicitEmptyOptionDraft.draftData).toMatchObject({
    nullable_category: "",
  });
  const explicitEmptyOptionObjectKey = explicitEmptyOptionDraft.draftObjectKey;
  if (!explicitEmptyOptionObjectKey) {
    throw new Error("The explicit empty option draft has no object key");
  }
  const explicitEmptyOptionXml = new TextDecoder().decode(
    unzipSync(await readObject(explicitEmptyOptionObjectKey))[
      "word/document.xml"
    ]
  );
  const explicitEmptyOptionContent = explicitEmptyOptionXml.match(
    /<w:tag w:val="nullable_category"\/>[\s\S]*?<w:sdtContent>(?<content>[\s\S]*?)<\/w:sdtContent>/u
  )?.[1];
  expect(explicitEmptyOptionContent).toContain(
    "<w:t>Nullable empty option label</w:t>"
  );
  return { onlyOfficeCanonicalData };
};
