// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

import { readObject } from "../../../src/storage";
import {
  contentControl,
  docxXmlFixture,
  contentControlDocument,
} from "../../fixtures/documents";
import {
  jsonHeaders,
  formCreationRequest,
  waitForOperation,
} from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import { capabilityHeaders } from "./helpers";
import type { ScenarioApp } from "./helpers";

export interface NativeFallbacksInput {
  app: ScenarioApp;
  adminBearer: string;
  userBearer: string;
}

export const runNativeFallbacks = async (
  input: NativeFallbacksInput
): Promise<void> => {
  const { app, adminBearer, userBearer } = input;

  const assertDateTemplateFallsBack = async (
    template: Uint8Array,
    name: string,
    title: string
  ): Promise<string> => {
    const fallbackCreateResponse = await app.handle(
      formCreationRequest({
        authorization: adminBearer,
        source: "upload",
        template: { bytes: template, name },
        title,
      })
    );
    expect(fallbackCreateResponse.status).toBe(200);
    const created = (await fallbackCreateResponse.json()) as {
      form?: { publicId?: string };
    };
    const fallbackPublicId = created.form?.publicId;
    if (!fallbackPublicId) {
      throw new Error(`${title} form was not created`);
    }
    const editorResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/forms/${fallbackPublicId}/editor-config`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    const editor = (await editorResponse.json()) as EditorConfigBody;
    const fallbackPublishCapability = editor.bridge.capabilities.publish;
    if (!fallbackPublishCapability) {
      throw new Error(`${title} publish capability is missing`);
    }
    const fallbackPublishResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/forms/${fallbackPublicId}/publish`,
        {
          body: JSON.stringify({ documentKey: editor.config.document.key }),
          headers: capabilityHeaders(fallbackPublishCapability),
          method: "POST",
        }
      )
    );
    const publish = (await fallbackPublishResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
    if (!publish.operationCapability || !publish.operationId) {
      throw new Error(`${title} publish operation was not created`);
    }
    expect(
      await waitForOperation(app, publish.operationId, {
        "X-Editor-Capability": publish.operationCapability,
      })
    ).toMatchObject({ status: "completed" });
    const detailResponse = await app.handle(
      new Request(`http://test.local/api/admin/forms/${fallbackPublicId}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(await detailResponse.json()).toMatchObject({
      form: { fillMethod: "onlyoffice", nativeFillAvailable: false },
    });
    const fallbackNativeMethodResponse = await app.handle(
      new Request(`http://test.local/api/admin/forms/${fallbackPublicId}`, {
        body: JSON.stringify({ fillMethod: "native" }),
        headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
        method: "PATCH",
      })
    );
    expect(fallbackNativeMethodResponse.status).toBe(409);
    const fallbackForm = await prisma.form.findUniqueOrThrow({
      select: { id: true },
      where: { publicId: fallbackPublicId },
    });
    const fallbackPublishedTemplate =
      await prisma.publishedTemplate.findUniqueOrThrow({
        select: { objectKey: true },
        where: { formId: fallbackForm.id },
      });
    const document = new TextDecoder().decode(
      unzipSync(await readObject(fallbackPublishedTemplate.objectKey))[
        "word/document.xml"
      ]
    );
    expect(document).toContain('w:fullDate="2024-02-29T00:00:00Z"');
    return fallbackPublicId;
  };
  const thaiCalendarTemplate = docxXmlFixture({
    document: contentControlDocument(
      contentControl({
        alias: "Thai calendar date",
        tag: "thai_date",
        type: '<w:date w:fullDate="2024-02-29T00:00:00Z"><w:calendar w:val="thai"/><w:dateFormat w:val="yyyy-MM-dd"/></w:date>',
      }) +
        contentControl({
          alias: "Thai calendar display date",
          tag: "thai_display_date",
          type: '<w:date w:fullDate="2024-02-29T00:00:00Z"><w:calendar w:val="thai"/><w:dateFormat w:val="yyyy-MM-dd"/></w:date>',
        }) +
        contentControl({
          alias: "Thai calendar checkbox",
          tag: "thai_checkbox",
          type: "<w14:checkbox/>",
        })
    ),
  });
  const thaiCalendarPublicId = await assertDateTemplateFallsBack(
    thaiCalendarTemplate,
    "thai-calendar.docx",
    "Ticket 07 Thai Calendar Form"
  );
  const thaiStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${thaiCalendarPublicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  const thaiStart = (await thaiStartResponse.json()) as {
    editorConfigUrl?: string;
    response?: { id?: string };
  };
  if (!thaiStart.editorConfigUrl || !thaiStart.response?.id) {
    throw new Error("The Thai calendar response did not start");
  }
  const thaiEditorResponse = await app.handle(
    new Request(
      new URL(thaiStart.editorConfigUrl, "http://test.local").toString(),
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  const thaiEditor = (await thaiEditorResponse.json()) as EditorConfigBody;
  const thaiSaveCapability = thaiEditor.bridge.capabilities["save-draft"];
  const thaiDocumentKey = thaiEditor.config.document.key;
  if (!thaiSaveCapability || !thaiDocumentKey) {
    throw new Error("The Thai calendar save capability is missing");
  }
  const thaiDraftResponse = await app.handle(
    new Request(`http://test.local/api/forms/${thaiCalendarPublicId}/draft`, {
      body: JSON.stringify({
        canonicalDateFields: ["thai_date"],
        data: {
          thai_checkbox: "☒",
          thai_date: "2025-03-04",
          thai_display_date: "2568-03-04",
        },
        documentKey: thaiDocumentKey,
        responseId: thaiStart.response.id,
      }),
      headers: capabilityHeaders(thaiSaveCapability),
      method: "POST",
    })
  );
  expect(thaiDraftResponse.status).toBe(202);
  const thaiDraft = (await thaiDraftResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!thaiDraft.operationCapability || !thaiDraft.operationId) {
    throw new Error("The Thai calendar draft operation was not created");
  }
  expect(
    await waitForOperation(app, thaiDraft.operationId, {
      "X-Editor-Capability": thaiDraft.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const thaiSavedResponse = await prisma.response.findUniqueOrThrow({
    select: { draftData: true },
    where: { id: thaiStart.response.id },
  });
  expect(thaiSavedResponse.draftData).toMatchObject({
    thai_checkbox: true,
    thai_date: "2025-03-04",
    thai_display_date: "2025-03-04",
  });
  const frenchDateTemplate = docxXmlFixture({
    document: contentControlDocument(
      contentControl({
        alias: "French date language",
        tag: "french_date",
        type: '<w:date w:fullDate="2024-02-29T00:00:00Z"><w:lid w:val="fr-FR"/><w:dateFormat w:val="d MMMM yyyy"/></w:date>',
      })
    ),
  });
  await assertDateTemplateFallsBack(
    frenchDateTemplate,
    "french-date.docx",
    "Ticket 07 French Date Language Form"
  );
  const nestedTemplate = docxXmlFixture({
    document: contentControlDocument(
      `<w:sdt><w:sdtPr><w:alias w:val="Group"/><w:tag w:val="group"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Group</w:t></w:r>${contentControl(
        {
          alias: "Nested name",
          tag: "nested_name",
          type: "<w:text/>",
        }
      )}</w:sdtContent></w:sdt>`
    ),
  });
  const nestedCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes: nestedTemplate, name: "nested-text.docx" },
      title: "Ticket 06 Nested Text Form",
    })
  );
  const nestedCreated = (await nestedCreateResponse.json()) as {
    form?: { publicId?: string };
  };
  const nestedPublicId = nestedCreated.form?.publicId;
  if (!nestedPublicId) {
    throw new Error("The Ticket 06 nested text form was not created");
  }
  const nestedEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${nestedPublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  const nestedEditor = (await nestedEditorResponse.json()) as EditorConfigBody;
  const nestedPublishCapability = nestedEditor.bridge.capabilities.publish;
  if (!nestedPublishCapability) {
    throw new Error("The Ticket 06 nested publish capability is missing");
  }
  const nestedPublishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${nestedPublicId}/publish`, {
      body: JSON.stringify({
        documentKey: nestedEditor.config.document.key,
      }),
      headers: capabilityHeaders(nestedPublishCapability),
      method: "POST",
    })
  );
  const nestedPublish = (await nestedPublishResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!nestedPublish.operationCapability || !nestedPublish.operationId) {
    throw new Error("The Ticket 06 nested publish operation was not created");
  }
  expect(
    await waitForOperation(app, nestedPublish.operationId, {
      "X-Editor-Capability": nestedPublish.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const nestedDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${nestedPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(await nestedDetailResponse.json()).toMatchObject({
    form: { fillMethod: "onlyoffice", nativeFillAvailable: false },
  });
  const nestedNativeResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${nestedPublicId}`, {
      body: JSON.stringify({ fillMethod: "native" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(nestedNativeResponse.status).toBe(409);
  expect(await nestedNativeResponse.json()).toMatchObject({
    error: "native_fill_unsupported",
  });
};
