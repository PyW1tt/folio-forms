// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

import { readObject } from "../../../src/storage";
import { docxXmlFixture } from "../../fixtures/documents";
import {
  jsonHeaders,
  formCreationRequest,
  waitForOperation,
} from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import { capabilityHeaders } from "./helpers";
import type { ScenarioApp } from "./helpers";

export interface CheckboxNamespacesInput {
  app: ScenarioApp;
  adminBearer: string;
  userBearer: string;
}

export const runCheckboxNamespaces = async (
  input: CheckboxNamespacesInput
): Promise<void> => {
  const { app, adminBearer, userBearer } = input;

  const checkboxTemplate = docxXmlFixture({
    document: `<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:word="urn:fixture"><body><sdt><sdtPr><alias w:val="Accept terms"/><tag w:val="accept_terms"/><checkbox xmlns="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checkedState w14:val="2713" w14:font="Ticket Symbols"/><w14:uncheckedState w14:val="25A1" w14:font="Ticket Symbols"/></checkbox></sdtPr><sdtContent><p><r><rPr><rFonts word:extension="keep"/></rPr><t>☐</t></r></p></sdtContent></sdt><sdt><sdtPr><alias w:val="Existing checked"/><tag w:val="existing_terms"/><checkbox xmlns="http://schemas.microsoft.com/office/word/2010/wordml"><checked val="0"/><w14:checkedState w14:val="2713" w14:font="Ticket Symbols"/><w14:uncheckedState w14:val="25A1" w14:font="Ticket Symbols"/></checkbox></sdtPr><sdtContent><p><r><t>☐</t></r></p></sdtContent></sdt><sdt><sdtPr><alias w:val="Self-closing checkbox"/><tag w:val="selfclosing_terms"/><checkbox xmlns="http://schemas.microsoft.com/office/word/2010/wordml"/></sdtPr><sdtContent><p><r><t>☐</t></r></p></sdtContent></sdt><sdt><sdtPr><alias w:val="Default date"/><tag w:val="default_date"/><date><dateFormat w:val="yyyy-MM-dd"/></date></sdtPr><sdtContent><p><r><t>2024-01-01</t></r></p></sdtContent></sdt><sdt><sdtPr><alias w:val="Default category"/><tag w:val="default_category"/><dropDownList><listItem w:displayText="Picked label" w:value="picked_value"/></dropDownList></sdtPr><sdtContent><p><r><t>Picked label</t></r></p></sdtContent></sdt><sdt xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><sdtPr><alias w:val="Strict terms"/><tag w:val="strict_terms"/><w14:checkbox><w14:checkedState w14:val="2713" w14:font="Ticket Symbols"/><w14:uncheckedState w14:val="25A1" w14:font="Ticket Symbols"/></w14:checkbox></sdtPr><sdtContent><p><r><rPr><rFonts word:extension="strict-keep"/></rPr><t>☐</t></r></p></sdtContent></sdt><sectPr/></body></document>`,
  });
  const checkboxCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes: checkboxTemplate, name: "checkbox.docx" },
      title: "Ticket 06 Checkbox Form",
    })
  );
  const checkboxCreated = (await checkboxCreateResponse.json()) as {
    form?: { publicId?: string };
  };
  const checkboxPublicId = checkboxCreated.form?.publicId;
  if (!checkboxPublicId) {
    throw new Error(
      `The Ticket 06 checkbox form was not created: ${JSON.stringify(checkboxCreated)} (HTTP ${checkboxCreateResponse.status})`
    );
  }
  const checkboxEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${checkboxPublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  const checkboxEditor =
    (await checkboxEditorResponse.json()) as EditorConfigBody;
  const checkboxPublishCapability = checkboxEditor.bridge.capabilities.publish;
  if (!checkboxPublishCapability) {
    throw new Error("The Ticket 06 checkbox publish capability is missing");
  }
  const checkboxPublishResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${checkboxPublicId}/publish`,
      {
        body: JSON.stringify({
          documentKey: checkboxEditor.config.document.key,
        }),
        headers: capabilityHeaders(checkboxPublishCapability),
        method: "POST",
      }
    )
  );
  const checkboxPublishBody = (await checkboxPublishResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !checkboxPublishBody.operationCapability ||
    !checkboxPublishBody.operationId
  ) {
    throw new Error("The Ticket 06 checkbox publish operation was not created");
  }
  expect(
    await waitForOperation(app, checkboxPublishBody.operationId, {
      "X-Editor-Capability": checkboxPublishBody.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const nativeCheckboxResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${checkboxPublicId}`, {
      body: JSON.stringify({ fillMethod: "native" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(nativeCheckboxResponse.status).toBe(200);
  const checkboxDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${checkboxPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(await checkboxDetailResponse.json()).toMatchObject({
    form: { fillMethod: "native", nativeFillAvailable: true },
  });
  const checkboxStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${checkboxPublicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  const checkboxStart = (await checkboxStartResponse.json()) as {
    editorConfigUrl?: string;
    response?: { id?: string };
  };
  if (!checkboxStart.editorConfigUrl || !checkboxStart.response?.id) {
    throw new Error("The default-namespace checkbox response did not start");
  }
  const checkboxNativeEditorResponse = await app.handle(
    new Request(
      new URL(checkboxStart.editorConfigUrl, "http://test.local").toString(),
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  const checkboxNativeEditor = (await checkboxNativeEditorResponse.json()) as {
    capabilities: Record<"save-draft", string>;
    documentKey: string;
  };
  const checkboxSaveResponse = await app.handle(
    new Request(`http://test.local/api/forms/${checkboxPublicId}/draft`, {
      body: JSON.stringify({
        data: {
          accept_terms: true,
          default_category: "picked_value",
          default_date: "2025-04-07",
          existing_terms: true,
          selfclosing_terms: true,
          strict_terms: true,
        },
        documentKey: checkboxNativeEditor.documentKey,
        fillMethod: "native",
        responseId: checkboxStart.response.id,
      }),
      headers: capabilityHeaders(
        checkboxNativeEditor.capabilities["save-draft"]
      ),
      method: "POST",
    })
  );
  const checkboxSave = (await checkboxSaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!checkboxSave.operationCapability || !checkboxSave.operationId) {
    throw new Error("The default-namespace checkbox save did not start");
  }
  expect(
    await waitForOperation(app, checkboxSave.operationId, {
      "X-Editor-Capability": checkboxSave.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const checkboxSaved = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: checkboxStart.response.id },
  });
  const checkboxSavedObjectKey = checkboxSaved.draftObjectKey;
  if (!checkboxSavedObjectKey) {
    throw new Error("The default-namespace checkbox draft has no object key");
  }
  const checkboxSavedXml = new TextDecoder().decode(
    unzipSync(await readObject(checkboxSavedObjectKey))["word/document.xml"]
  );
  expect(checkboxSavedXml).toContain(
    '<rFonts word:extension="keep" xmlns:word1="http://schemas.openxmlformats.org/wordprocessingml/2006/main" word1:ascii="Ticket Symbols" word1:hAnsi="Ticket Symbols"/>'
  );
  expect(checkboxSavedXml).toContain(
    '<rFonts word:extension="strict-keep" xmlns:word1="http://purl.oclc.org/ooxml/wordprocessingml/main" word1:ascii="Ticket Symbols" word1:hAnsi="Ticket Symbols"/>'
  );
  expect(checkboxSavedXml).toMatch(
    /<date[^>]*word:fullDate="2025-04-07T00:00:00Z"[^>]*xmlns:word="http:\/\/schemas\.openxmlformats\.org\/wordprocessingml\/2006\/main"/u
  );
  expect(checkboxSavedXml).toMatch(
    /<dropDownList[^>]*word:lastValue="picked_value"[^>]*xmlns:word="http:\/\/schemas\.openxmlformats\.org\/wordprocessingml\/2006\/main"/u
  );
  expect(checkboxSavedXml).toContain('xmlns:word="urn:fixture"');
  expect(checkboxSavedXml.match(/<checked w14:val="1"\/>/gu)).toHaveLength(3);
  expect(checkboxSavedXml).not.toMatch(/<checked\s+val=/u);
  expect(checkboxSavedXml).not.toMatch(
    /<(?:rFonts|date|dropDownList)[^>]*\s(?:ascii|hAnsi|fullDate|lastValue)=/u
  );
  const strictCheckboxTemplate = docxXmlFixture({
    document: `<document xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:word="urn:fixture"><body><sdt><sdtPr><alias w:val="Strict terms"/><tag w:val="strict_terms"/><w14:checkbox><w14:checkedState w14:val="2713" w14:font="Ticket Symbols"/><w14:uncheckedState w14:val="25A1" w14:font="Ticket Symbols"/></w14:checkbox></sdtPr><sdtContent/></sdt><sectPr/></body></document>`,
  });
  const strictCheckboxCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes: strictCheckboxTemplate, name: "strict-checkbox.docx" },
      title: "Ticket 07 Strict WordML Checkbox",
    })
  );
  const strictCheckboxCreated = (await strictCheckboxCreateResponse.json()) as {
    form?: { publicId?: string };
  };
  const strictCheckboxPublicId = strictCheckboxCreated.form?.publicId;
  if (!strictCheckboxPublicId) {
    throw new Error("The Strict WordML checkbox form was not created");
  }
  const strictCheckboxAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${strictCheckboxPublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  const strictCheckboxAdminEditor =
    (await strictCheckboxAdminEditorResponse.json()) as EditorConfigBody;
  const strictCheckboxPublishCapability =
    strictCheckboxAdminEditor.bridge.capabilities.publish;
  if (!strictCheckboxPublishCapability) {
    throw new Error("The Strict WordML checkbox publish capability is missing");
  }
  const strictCheckboxPublishResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${strictCheckboxPublicId}/publish`,
      {
        body: JSON.stringify({
          documentKey: strictCheckboxAdminEditor.config.document.key,
        }),
        headers: capabilityHeaders(strictCheckboxPublishCapability),
        method: "POST",
      }
    )
  );
  const strictCheckboxPublish =
    (await strictCheckboxPublishResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !strictCheckboxPublish.operationCapability ||
    !strictCheckboxPublish.operationId
  ) {
    throw new Error("The Strict WordML checkbox publish did not start");
  }
  expect(
    await waitForOperation(app, strictCheckboxPublish.operationId, {
      "X-Editor-Capability": strictCheckboxPublish.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const strictCheckboxNativeResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${strictCheckboxPublicId}`, {
      body: JSON.stringify({ fillMethod: "native" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(strictCheckboxNativeResponse.status).toBe(200);
  const strictCheckboxStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${strictCheckboxPublicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  const strictCheckboxStart = (await strictCheckboxStartResponse.json()) as {
    editorConfigUrl?: string;
    response?: { id?: string };
  };
  if (
    !strictCheckboxStart.editorConfigUrl ||
    !strictCheckboxStart.response?.id
  ) {
    throw new Error("The Strict WordML checkbox response did not start");
  }
  const strictCheckboxEditorResponse = await app.handle(
    new Request(
      new URL(
        strictCheckboxStart.editorConfigUrl,
        "http://test.local"
      ).toString(),
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  const strictCheckboxEditor = (await strictCheckboxEditorResponse.json()) as {
    capabilities: Record<"save-draft", string>;
    documentKey: string;
  };
  const strictCheckboxSaveResponse = await app.handle(
    new Request(`http://test.local/api/forms/${strictCheckboxPublicId}/draft`, {
      body: JSON.stringify({
        data: { strict_terms: true },
        documentKey: strictCheckboxEditor.documentKey,
        fillMethod: "native",
        responseId: strictCheckboxStart.response.id,
      }),
      headers: capabilityHeaders(
        strictCheckboxEditor.capabilities["save-draft"]
      ),
      method: "POST",
    })
  );
  const strictCheckboxSave = (await strictCheckboxSaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !strictCheckboxSave.operationCapability ||
    !strictCheckboxSave.operationId
  ) {
    throw new Error("The Strict WordML checkbox save did not start");
  }
  expect(
    await waitForOperation(app, strictCheckboxSave.operationId, {
      "X-Editor-Capability": strictCheckboxSave.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const strictCheckboxSaved = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: strictCheckboxStart.response.id },
  });
  const strictCheckboxSavedObjectKey = strictCheckboxSaved.draftObjectKey;
  if (!strictCheckboxSavedObjectKey) {
    throw new Error("The Strict WordML checkbox draft has no object key");
  }
  const strictCheckboxSavedXml = new TextDecoder().decode(
    unzipSync(await readObject(strictCheckboxSavedObjectKey))[
      "word/document.xml"
    ]
  );
  expect(strictCheckboxSavedXml).toContain(
    '<rFonts xmlns:word="http://purl.oclc.org/ooxml/wordprocessingml/main" word:ascii="Ticket Symbols" word:hAnsi="Ticket Symbols"/>'
  );
  expect(strictCheckboxSavedXml).toContain('xmlns:word="urn:fixture"');
};
