import { test, expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

import { createApp } from "../../src/app";
import { readObject } from "../../src/storage";
import { docxXmlFixture } from "../fixtures/documents";
import {
  createCredentialFixture,
  bearerFor,
  jsonHeaders,
  formCreationRequest,
  waitForOperation,
} from "../fixtures/http";
import type { EditorConfigBody } from "../fixtures/http";
import { runCheckboxNamespaces } from "../scenarios/native-scalar/checkboxes";
import { runNativeDraft } from "../scenarios/native-scalar/draft";
import { runNativeFallbacks } from "../scenarios/native-scalar/fallbacks";
import { runOfficeCanonicalValues } from "../scenarios/native-scalar/office-values";
import { runNativeSetup } from "../scenarios/native-scalar/setup";
import { runNativeResubmit } from "../scenarios/native-scalar/submission";

// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  throw new Error(
    "The HTTP application test requires DATABASE_URL for an isolated PostgreSQL database"
  );
}
const convertedDocumentKeys: string[] = [];
const app = createApp({
  legacySso: null,
  onlyOffice: {
    convertDocxToPdf: (documentKey) => {
      convertedDocumentKeys.push(documentKey);
      return Promise.resolve(new TextEncoder().encode("%PDF-test"));
    },
    forceSave: () => Promise.resolve(false),
  },
  prefillReturnUrl: "https://source.example.test/forms/return",
  requestIp: (request) => request.headers.get("x-test-ip"),
});

test("Ticket 07 native scalar forms preserve values, validation, and Fill Method", async () => {
  const setup = await runNativeSetup({ app });
  const draft = await runNativeDraft({
    app,
    editorConfigUrl: setup.editorConfigUrl,
    prefillLocks: setup.prefillLocks,
    prefillValues: setup.prefillValues,
    publicId: setup.publicId,
    responseId: setup.responseId,
    userBearer: setup.userBearer,
  });
  const officeValues = await runOfficeCanonicalValues({
    adminBearer: setup.adminBearer,
    app,
    draftData: draft.draftData,
    nativeConfig: draft.nativeConfig,
    prefillValues: setup.prefillValues,
    publicId: setup.publicId,
    responseId: setup.responseId,
    savedDraftDocumentKey: draft.savedDraftDocumentKey,
    userBearer: setup.userBearer,
  });
  await runNativeResubmit({
    adminBearer: setup.adminBearer,
    app,
    editorConfigUrl: setup.editorConfigUrl,
    onlyOfficeCanonicalData: officeValues.onlyOfficeCanonicalData,
    prefillLocks: setup.prefillLocks,
    publicId: setup.publicId,
    responseId: setup.responseId,
    userBearer: setup.userBearer,
  });
  await runCheckboxNamespaces({
    adminBearer: setup.adminBearer,
    app,
    userBearer: setup.userBearer,
  });
  await runNativeFallbacks({
    adminBearer: setup.adminBearer,
    app,
    userBearer: setup.userBearer,
  });
});

test("Ticket 07 preserves local namespaces when expanding empty content", async () => {
  const adminEmail = `ticket-07-local-namespace-admin-${crypto.randomUUID()}@example.com`;
  const userEmail = `ticket-07-local-namespace-user-${crypto.randomUUID()}@example.com`;
  const password = "Ticket07-local-namespace-password";
  await createCredentialFixture({
    email: adminEmail,
    name: "Ticket 07 Local Namespace Admin",
    password,
    role: "admin",
  });
  await createCredentialFixture({
    email: userEmail,
    name: "Ticket 07 Local Namespace User",
    password,
  });
  const adminBearer = await bearerFor(app, adminEmail, password);
  const userBearer = await bearerFor(app, userEmail, password);
  const wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const template = docxXmlFixture({
    document:
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<word:document xmlns:word="${wordNamespace}"><word:body>` +
      `<word:sdt><word:sdtPr><word:alias word:val="Local namespace text"/>` +
      `<word:tag word:val="local_namespace_text"/><word:text/></word:sdtPr>` +
      `<w:sdtContent xmlns:w="${wordNamespace}"/></word:sdt>` +
      `<word:sectPr/></word:body></word:document>`,
  });
  const createResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes: template, name: "local-namespace.docx" },
      title: "Ticket 07 Local Content Namespace",
    })
  );
  expect(createResponse.status).toBe(200);
  const created = (await createResponse.json()) as {
    form?: { publicId?: string };
  };
  const publicId = created.form?.publicId;
  if (!publicId) {
    throw new Error("The local namespace form was not created");
  }
  const editorResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/editor-config`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  const editor = (await editorResponse.json()) as EditorConfigBody;
  const publishCapability = editor.bridge.capabilities.publish;
  const templateDocumentKey = editor.config.document.key;
  if (!publishCapability || !templateDocumentKey) {
    throw new Error("The local namespace publish capability is missing");
  }
  const publishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/publish`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: {
        ...jsonHeaders,
        "X-Editor-Capability": publishCapability,
      },
      method: "POST",
    })
  );
  expect(publishResponse.status).toBe(202);
  const publishOperation = (await publishResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!publishOperation.operationCapability || !publishOperation.operationId) {
    throw new Error("The local namespace publish operation did not start");
  }
  expect(
    await waitForOperation(app, publishOperation.operationId, {
      "X-Editor-Capability": publishOperation.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const nativeMethodResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}`, {
      body: JSON.stringify({ fillMethod: "native" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(nativeMethodResponse.status).toBe(200);
  const startResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(startResponse.status).toBe(200);
  const start = (await startResponse.json()) as {
    editorConfigUrl?: string;
    response?: { id?: string };
  };
  const responseId = start.response?.id;
  if (!start.editorConfigUrl || !responseId) {
    throw new Error("The local namespace response did not start");
  }
  const responseEditor = await app.handle(
    new Request(
      new URL(start.editorConfigUrl, "http://test.local").toString(),
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(responseEditor.status).toBe(200);
  const responseConfig = (await responseEditor.json()) as {
    capabilities?: Partial<Record<"save-draft" | "submit", string>>;
    documentKey?: string;
  };
  const saveCapability = responseConfig.capabilities?.["save-draft"];
  const responseDocumentKey = responseConfig.documentKey;
  if (!saveCapability || !responseDocumentKey) {
    throw new Error("The local namespace save capability is missing");
  }
  const saveResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/draft`, {
      body: JSON.stringify({
        data: { local_namespace_text: "Saved text" },
        documentKey: responseDocumentKey,
        fillMethod: "native",
        responseId,
      }),
      headers: {
        ...jsonHeaders,
        "X-Editor-Capability": saveCapability,
      },
      method: "POST",
    })
  );
  expect(saveResponse.status).toBe(202);
  const saveOperation = (await saveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!saveOperation.operationCapability || !saveOperation.operationId) {
    throw new Error("The local namespace save operation did not start");
  }
  expect(
    await waitForOperation(app, saveOperation.operationId, {
      "X-Editor-Capability": saveOperation.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const savedDraft = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: responseId },
  });
  if (!savedDraft.draftObjectKey) {
    throw new Error("The local namespace draft was not stored");
  }
  const savedXml = new TextDecoder().decode(
    unzipSync(await readObject(savedDraft.draftObjectKey))["word/document.xml"]
  );
  expect(savedXml).toContain(
    `<w:sdtContent xmlns:w="${wordNamespace}"><w:p><w:r>`
  );
  expect(savedXml).toContain('<w:t xml:space="preserve">Saved text</w:t>');
});
