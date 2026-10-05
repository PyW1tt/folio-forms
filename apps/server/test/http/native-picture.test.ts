import { test, expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

import { createApp } from "../../src/app";
import { createOnlyOfficeAuthorization } from "../../src/onlyoffice";
import { readObject } from "../../src/storage";
import {
  pictureDocumentFixture,
  pngFixture,
  jpegFixture,
} from "../fixtures/documents";
import {
  createCredentialFixture,
  bearerFor,
  formCreationRequest,
  jsonHeaders,
  editorCapabilityHeaders,
  waitForOperation,
} from "../fixtures/http";
import type { EditorConfigBody } from "../fixtures/http";

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

test("Ticket 08 native Picture uploads survive response workflows", async () => {
  const adminPassword = "Ticket08-picture-admin-password";
  const userPassword = "Ticket08-picture-user-password";
  const adminEmail = `ticket-08-picture-admin-${crypto.randomUUID()}@example.com`;
  const userEmail = `ticket-08-picture-user-${crypto.randomUUID()}@example.com`;
  await createCredentialFixture({
    email: adminEmail,
    name: "Ticket 08 Picture Admin",
    password: adminPassword,
    role: "admin",
  });
  await createCredentialFixture({
    email: userEmail,
    name: "Ticket 08 Picture User",
    password: userPassword,
  });
  const adminBearer = await bearerFor(app, adminEmail, adminPassword);
  const userBearer = await bearerFor(app, userEmail, userPassword);
  const createResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: pictureDocumentFixture(),
        name: "ticket-08-picture.docx",
      },
      title: "Ticket 08 Native Picture",
    })
  );
  expect(createResponse.status).toBe(200);
  const created = (await createResponse.json()) as {
    form?: { publicId?: string };
  };
  const publicId = created.form?.publicId;
  if (!publicId) {
    throw new Error("The Ticket 08 Picture form was not created");
  }
  const adminEditorResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/editor-config`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  const adminEditor = (await adminEditorResponse.json()) as EditorConfigBody;
  const configureCapability =
    adminEditor.bridge.capabilities["configure-fields"];
  if (!configureCapability) {
    throw new Error("The Ticket 08 Picture configure capability is missing");
  }
  const fieldRuleResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/field-rules`, {
      body: JSON.stringify({
        documentKey: adminEditor.config.document.key,
        prefillPointer: null,
        prefillPolicy: "editable",
        previousTag: null,
        required: true,
        tag: "photo",
      }),
      headers: {
        ...jsonHeaders,
        ...editorCapabilityHeaders(configureCapability),
      },
      method: "PATCH",
    })
  );
  expect(fieldRuleResponse.status).toBe(200);
  const publishEditorResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/editor-config`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  const publishEditor =
    (await publishEditorResponse.json()) as EditorConfigBody;
  const publishCapability = publishEditor.bridge.capabilities.publish;
  if (!publishCapability) {
    throw new Error("The Ticket 08 Picture publish capability is missing");
  }
  const publishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/publish`, {
      body: JSON.stringify({
        documentKey: publishEditor.config.document.key,
      }),
      headers: {
        ...jsonHeaders,
        ...editorCapabilityHeaders(publishCapability),
      },
      method: "POST",
    })
  );
  expect(publishResponse.status).toBe(202);
  const publish = (await publishResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!publish.operationCapability || !publish.operationId) {
    throw new Error("The Ticket 08 Picture publish operation is missing");
  }
  expect(
    await waitForOperation(app, publish.operationId, {
      "X-Editor-Capability": publish.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const formDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(await formDetailResponse.json()).toMatchObject({
    form: { nativeFillAvailable: true },
  });
  const methodResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}`, {
      body: JSON.stringify({ fillMethod: "native" }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${adminBearer}`,
      },
      method: "PATCH",
    })
  );
  expect(methodResponse.status).toBe(200);
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
  const { editorConfigUrl } = start;
  if (!responseId || !editorConfigUrl) {
    throw new Error("The Ticket 08 Picture response did not start");
  }
  interface PictureNativeConfig {
    capabilities: { "save-draft": string; submit: string };
    documentKey: string;
    pictures: Record<string, boolean>;
    responseId: string;
  }
  const getNativeConfig = async () => {
    const response = await app.handle(
      new Request(new URL(editorConfigUrl, "http://test.local").toString(), {
        headers: { Authorization: `Bearer ${userBearer}` },
      })
    );
    expect(response.status).toBe(200);
    return (await response.json()) as PictureNativeConfig;
  };
  let nativeConfig = await getNativeConfig();
  expect(nativeConfig).toMatchObject({
    pictures: { photo: false },
    responseId,
  });
  const upload = (
    action: "draft" | "submit",
    image?: { bytes: Uint8Array; name: string; type: string }
  ) => {
    const body = new FormData();
    body.set(
      "payload",
      JSON.stringify({
        data: { photo: "must not enter scalar data" },
        documentKey: nativeConfig.documentKey,
        fillMethod: "native",
        responseId,
      })
    );
    if (image) {
      body.set(
        "picture:photo",
        new File([image.bytes], image.name, {
          type: image.type,
        })
      );
    }
    return app.handle(
      new Request(`http://test.local/api/forms/${publicId}/${action}`, {
        body,
        headers: editorCapabilityHeaders(
          nativeConfig.capabilities[
            action === "draft" ? "save-draft" : "submit"
          ]
        ),
        method: "POST",
      })
    );
  };
  const initialDraft = await prisma.response.findUniqueOrThrow({
    select: { draftData: true, draftObjectKey: true, status: true },
    where: { id: responseId },
  });
  const missingRequired = await upload("submit");
  expect(missingRequired.status).toBe(422);
  expect(
    await prisma.response.findUniqueOrThrow({
      select: { draftData: true, draftObjectKey: true, status: true },
      where: { id: responseId },
    })
  ).toEqual(initialDraft);
  const oversized = await upload("draft", {
    bytes: pngFixture(1, 1, 10 * 1024 * 1024 + 1),
    name: "oversized.png",
    type: "image/png",
  });
  expect(oversized.status).toBe(413);
  const wrongMime = await upload("draft", {
    bytes: pngFixture(),
    name: "wrong-type.gif",
    type: "image/gif",
  });
  expect(wrongMime.status).toBe(415);
  const oversizedDimensions = await upload("draft", {
    bytes: pngFixture(4097, 1),
    name: "too-wide.png",
    type: "image/png",
  });
  expect(oversizedDimensions.status).toBe(422);
  const png = pngFixture();
  const savePngResponse = await upload("draft", {
    bytes: png,
    name: "photo.png",
    type: "image/png",
  });
  expect(savePngResponse.status).toBe(202);
  const savePng = (await savePngResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!savePng.operationCapability || !savePng.operationId) {
    throw new Error("The Ticket 08 PNG draft operation is missing");
  }
  expect(
    await waitForOperation(app, savePng.operationId, {
      "X-Editor-Capability": savePng.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const savedDraft = await prisma.response.findUniqueOrThrow({
    select: {
      draftData: true,
      draftDocumentKey: true,
      draftObjectKey: true,
      status: true,
    },
    where: { id: responseId },
  });
  expect(savedDraft.draftData).toEqual({});
  expect(savedDraft.status).toBe("draft");
  if (!savedDraft.draftObjectKey || !savedDraft.draftDocumentKey) {
    throw new Error("The Ticket 08 PNG draft object is missing");
  }
  const pngDocx = await readObject(savedDraft.draftObjectKey);
  const pngArchive = unzipSync(pngDocx);
  expect(
    Object.entries(pngArchive).some(
      ([path, bytes]) =>
        path.startsWith("word/media/") &&
        Buffer.from(bytes).equals(Buffer.from(png))
    )
  ).toBe(true);
  nativeConfig = await getNativeConfig();
  expect(nativeConfig.pictures).toMatchObject({ photo: true });
  const resumedResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(await resumedResponse.json()).toMatchObject({
    response: { id: responseId, status: "draft" },
  });
  const persistedPngDraft = await prisma.response.findUniqueOrThrow({
    select: { draftData: true, draftObjectKey: true },
    where: { id: responseId },
  });
  const failedReplacement = await upload("draft", {
    bytes: png,
    name: "wrong-type.gif",
    type: "image/gif",
  });
  expect(failedReplacement.status).toBe(415);
  expect(
    await prisma.response.findUniqueOrThrow({
      select: { draftData: true, draftObjectKey: true },
      where: { id: responseId },
    })
  ).toEqual(persistedPngDraft);
  if (!persistedPngDraft.draftObjectKey) {
    throw new Error("The saved Ticket 08 PNG draft object is missing");
  }
  expect(await readObject(persistedPngDraft.draftObjectKey)).toEqual(pngDocx);
  const draftPdf = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/pdf`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftPdf.status).toBe(200);
  expect(draftPdf.headers.get("content-type")).toBe("application/pdf");
  expect(convertedDocumentKeys).toContain(savedDraft.draftDocumentKey);
  const onlyOfficeMethod = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}`, {
      body: JSON.stringify({ fillMethod: "onlyoffice" }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${adminBearer}`,
      },
      method: "PATCH",
    })
  );
  expect(onlyOfficeMethod.status).toBe(200);
  const onlyOfficeConfigResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${publicId}/editor-config?responseId=${responseId}&action=draft`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(onlyOfficeConfigResponse.status).toBe(200);
  const onlyOfficeConfig =
    (await onlyOfficeConfigResponse.json()) as EditorConfigBody;
  const docxUrl = onlyOfficeConfig.config.document.url;
  const docxResponse = await app.handle(
    new Request(docxUrl, {
      headers: {
        Authorization: createOnlyOfficeAuthorization({ url: docxUrl }),
      },
    })
  );
  expect(docxResponse.status).toBe(200);
  expect(
    Buffer.from(new Uint8Array(await docxResponse.arrayBuffer())).equals(
      Buffer.from(pngDocx)
    )
  ).toBe(true);
  const backToNative = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}`, {
      body: JSON.stringify({ fillMethod: "native" }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${adminBearer}`,
      },
      method: "PATCH",
    })
  );
  expect(backToNative.status).toBe(200);
  nativeConfig = await getNativeConfig();
  expect(nativeConfig.pictures).toMatchObject({ photo: true });
  const jpeg = jpegFixture();
  const saveJpegResponse = await upload("draft", {
    bytes: jpeg,
    name: "photo.jpg",
    type: "image/jpeg",
  });
  expect(saveJpegResponse.status).toBe(202);
  const saveJpeg = (await saveJpegResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!saveJpeg.operationCapability || !saveJpeg.operationId) {
    throw new Error("The Ticket 08 JPEG draft operation is missing");
  }
  expect(
    await waitForOperation(app, saveJpeg.operationId, {
      "X-Editor-Capability": saveJpeg.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  nativeConfig = await getNativeConfig();
  const savedJpegDraft = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: responseId },
  });
  if (!savedJpegDraft.draftObjectKey) {
    throw new Error("The saved Ticket 08 JPEG draft object is missing");
  }
  const pictureMedia = Object.entries(
    unzipSync(await readObject(savedJpegDraft.draftObjectKey))
  ).filter(([path]) => path.startsWith("word/media/picture-"));
  expect(pictureMedia).toHaveLength(1);
  expect(
    Buffer.from(pictureMedia[0]?.[1] ?? new Uint8Array()).equals(
      Buffer.from(jpeg)
    )
  ).toBe(true);
  const submitResponse = await upload("submit");
  expect(submitResponse.status).toBe(202);
  const submit = (await submitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!submit.operationCapability || !submit.operationId) {
    throw new Error("The Ticket 08 submission operation is missing");
  }
  expect(
    await waitForOperation(app, submit.operationId, {
      "X-Editor-Capability": submit.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const submission = await prisma.submission.findUniqueOrThrow({
    select: { data: true, objectKey: true },
    where: { responseId },
  });
  expect(submission.data).toEqual({});
  const submissionDocx = await readObject(submission.objectKey);
  expect(
    Object.entries(unzipSync(submissionDocx)).some(
      ([path, bytes]) =>
        path.startsWith("word/media/") &&
        Buffer.from(bytes).equals(Buffer.from(jpeg))
    )
  ).toBe(true);
});
