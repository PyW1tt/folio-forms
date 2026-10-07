// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

import { createApp } from "../../../src/app";
import { readObject, DOCX_CONTENT_TYPE, putObject } from "../../../src/storage";
import {
  pictureDocumentFixture,
  pngFixture,
  gifFixture,
  jpegFixture,
} from "../../fixtures/documents";
import { formCreationRequest, waitForOperation } from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import { capabilityHeaders } from "./helpers";
import type { BootstrapAndCreationOutput } from "./setup";

export interface OfficePictureResponseInput {
  app: ReturnType<typeof createApp>;
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
}

export interface OfficePictureResponseOutput {
  pictureResponseId: string;
  pictureCorrectionEditor: EditorConfigBody;
}

export const runOfficePictureResponse = async (
  input: OfficePictureResponseInput
): Promise<OfficePictureResponseOutput> => {
  const { app, adminBearer } = input;
  const requiredPictureCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: pictureDocumentFixture(),
        name: "required-picture.docx",
      },
      title: "Ticket 16 required picture",
    })
  );
  expect(requiredPictureCreateResponse.status).toBe(200);
  const requiredPictureCreateBody =
    (await requiredPictureCreateResponse.json()) as {
      form?: { publicId?: string };
    };
  const requiredPicturePublicId = requiredPictureCreateBody.form?.publicId;
  if (!requiredPicturePublicId) {
    throw new Error("The required picture form was not created");
  }
  const requiredPictureEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(requiredPictureEditorResponse.status).toBe(200);
  const requiredPictureEditor =
    (await requiredPictureEditorResponse.json()) as EditorConfigBody;
  const configurePictureCapability =
    requiredPictureEditor.bridge.capabilities["configure-fields"];
  if (!configurePictureCapability) {
    throw new Error(
      "The required picture configure capability was not returned"
    );
  }
  const requiredPictureRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: requiredPictureEditor.config.document.key,
          prefillPointer: null,
          prefillPolicy: "editable",
          previousTag: null,
          required: true,
          tag: "photo",
        }),
        headers: capabilityHeaders(configurePictureCapability),
        method: "PATCH",
      }
    )
  );
  expect(requiredPictureRuleResponse.status).toBe(200);
  const requiredPicturePublishEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  const requiredPicturePublishEditor =
    (await requiredPicturePublishEditorResponse.json()) as EditorConfigBody;
  const requiredPicturePublishCapability =
    requiredPicturePublishEditor.bridge.capabilities.publish;
  if (!requiredPicturePublishCapability) {
    throw new Error("The required picture publish capability was not returned");
  }
  const requiredPicturePublishResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/publish`,
      {
        body: JSON.stringify({
          documentKey: requiredPicturePublishEditor.config.document.key,
        }),
        headers: capabilityHeaders(requiredPicturePublishCapability),
        method: "POST",
      }
    )
  );
  expect(requiredPicturePublishResponse.status).toBe(202);
  const requiredPicturePublishBody =
    (await requiredPicturePublishResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !requiredPicturePublishBody.operationCapability ||
    !requiredPicturePublishBody.operationId
  ) {
    throw new Error("The required picture publish operation was not created");
  }
  const requiredPicturePublishOperation = await waitForOperation(
    app,
    requiredPicturePublishBody.operationId,
    {
      "X-Editor-Capability": requiredPicturePublishBody.operationCapability,
    }
  );
  expect(requiredPicturePublishOperation.status).toBe("completed");
  const requiredPictureForm = await prisma.form.findUniqueOrThrow({
    select: { id: true },
    where: { publicId: requiredPicturePublicId },
  });
  const requiredPictureManifest =
    await prisma.publishedTemplate.findUniqueOrThrow({
      include: { manifest: { include: { fields: true } } },
      where: { formId: requiredPictureForm.id },
    });
  expect(requiredPictureManifest.manifest?.fields).toContainEqual(
    expect.objectContaining({
      pictureMaxBytes: 10 * 1024 * 1024,
      pictureMaxHeight: 4096,
      pictureMaxWidth: 4096,
      required: true,
      tag: "photo",
      type: "picture",
    })
  );
  const pictureFormDetailResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(await pictureFormDetailResponse.json()).toMatchObject({
    form: { nativeFillAvailable: true },
  });
  let nextPictureDocument = pictureDocumentFixture({
    images: [{ bytes: pngFixture(), extension: "png" }],
    includeStaticImage: true,
  });
  const pictureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.resolve(new TextEncoder().encode("%PDF-picture")),
      forceSave: async (documentKey) => {
        const response = await prisma.response.findUnique({
          select: { draftObjectKey: true },
          where: { draftDocumentKey: documentKey },
        });
        if (!response?.draftObjectKey) {
          throw new Error("The picture response document was not found");
        }
        await putObject(
          response.draftObjectKey,
          nextPictureDocument,
          DOCX_CONTENT_TYPE
        );
        return false;
      },
    },
  });
  const pictureStartResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/forms/${requiredPicturePublicId}/start`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      }
    )
  );
  expect(pictureStartResponse.status).toBe(200);
  const pictureStartBody = (await pictureStartResponse.json()) as {
    response?: { id?: string };
  };
  const pictureResponseId = pictureStartBody.response?.id;
  if (!pictureResponseId) {
    throw new Error("The picture response was not started");
  }
  const pictureEditorResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/forms/${requiredPicturePublicId}/editor-config?responseId=${pictureResponseId}&action=draft`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(pictureEditorResponse.status).toBe(200);
  const pictureEditor =
    (await pictureEditorResponse.json()) as EditorConfigBody;
  let pictureDocumentKey = pictureEditor.config.document.key;
  let pictureSaveCapability: string =
    pictureEditor.bridge.capabilities["save-draft"] ?? "";
  let pictureSubmitCapability: string =
    pictureEditor.bridge.capabilities.submit ?? "";
  if (!pictureSaveCapability || !pictureSubmitCapability) {
    throw new Error("The picture response capabilities were not returned");
  }
  const refreshPictureEditor = async (): Promise<void> => {
    const refreshedResponse = await pictureApp.handle(
      new Request(
        `http://test.local/api/forms/${requiredPicturePublicId}/editor-config?responseId=${pictureResponseId}&action=draft`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    expect(refreshedResponse.status).toBe(200);
    const refreshedConfig =
      (await refreshedResponse.json()) as EditorConfigBody;
    const refreshedSaveCapability =
      refreshedConfig.bridge.capabilities["save-draft"];
    const refreshedSubmitCapability =
      refreshedConfig.bridge.capabilities.submit;
    if (!refreshedSaveCapability || !refreshedSubmitCapability) {
      throw new Error(
        "The refreshed picture response capabilities were not returned"
      );
    }
    pictureDocumentKey = refreshedConfig.config.document.key;
    pictureSaveCapability = refreshedSaveCapability;
    pictureSubmitCapability = refreshedSubmitCapability;
  };
  const pictureDraftRequest = (data: Record<string, unknown>) =>
    pictureApp.handle(
      new Request(
        `http://test.local/api/forms/${requiredPicturePublicId}/draft`,
        {
          body: JSON.stringify({
            data,
            documentKey: pictureDocumentKey,
            responseId: pictureResponseId,
          }),
          headers: capabilityHeaders(pictureSaveCapability),
          method: "POST",
        }
      )
    );
  const pictureSubmitRequest = (data: Record<string, unknown>) =>
    pictureApp.handle(
      new Request(
        `http://test.local/api/forms/${requiredPicturePublicId}/submit`,
        {
          body: JSON.stringify({
            data,
            documentKey: pictureDocumentKey,
            responseId: pictureResponseId,
          }),
          headers: capabilityHeaders(pictureSubmitCapability),
          method: "POST",
        }
      )
    );
  const savePicture = async (
    document: Uint8Array,
    data: Record<string, unknown>
  ) => {
    nextPictureDocument = document;
    const response = await pictureDraftRequest(data);
    expect(response.status).toBe(202);
    const body = (await response.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
    if (!body.operationCapability || !body.operationId) {
      throw new Error("The picture draft operation was not created");
    }
    const operation = await waitForOperation(app, body.operationId, {
      "X-Editor-Capability": body.operationCapability,
    });
    if (operation.status === "completed") {
      await refreshPictureEditor();
    }
    return operation;
  };
  const validPictureDocument = nextPictureDocument;
  const validPictureSave = await savePicture(validPictureDocument, {
    photo: "scalar data must be omitted",
  });
  expect(validPictureSave.status).toBe("completed");
  const savedPictureResponse = await prisma.response.findUniqueOrThrow({
    select: { draftData: true, draftObjectKey: true, status: true },
    where: { id: pictureResponseId },
  });
  expect(savedPictureResponse).toMatchObject({
    draftData: {},
    status: "draft",
  });
  if (!savedPictureResponse.draftObjectKey) {
    throw new Error("The saved picture draft object was not persisted");
  }
  const savedPictureArchive = unzipSync(
    await readObject(savedPictureResponse.draftObjectKey)
  );
  expect(savedPictureArchive["word/media/image1.png"]).toEqual(pngFixture());
  expect(savedPictureArchive["word/media/image2.png"]).toEqual(pngFixture());
  expect(
    new TextDecoder().decode(savedPictureArchive["word/document.xml"])
  ).toContain('r:embed="rIdPicture1"');
  expect(
    new TextDecoder().decode(savedPictureArchive["word/document.xml"])
  ).toContain('r:embed="rIdStatic"');
  const resumedPictureStartResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/forms/${requiredPicturePublicId}/start`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      }
    )
  );
  expect(await resumedPictureStartResponse.json()).toMatchObject({
    response: { id: pictureResponseId, status: "draft" },
  });
  const pictureDocxExport = await pictureApp.handle(
    new Request(
      `http://test.local/api/responses/${pictureResponseId}/draft/docx`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(pictureDocxExport.status).toBe(200);
  const exportedPictureArchive = unzipSync(
    new Uint8Array(await pictureDocxExport.arrayBuffer())
  );
  expect(exportedPictureArchive["word/media/image1.png"]).toEqual(pngFixture());
  expect(exportedPictureArchive["word/media/image2.png"]).toEqual(pngFixture());
  const picturePdfExport = await pictureApp.handle(
    new Request(
      `http://test.local/api/responses/${pictureResponseId}/draft/pdf`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(await picturePdfExport.text()).toBe("%PDF-picture");
  const missingPictureDraft = await savePicture(pictureDocumentFixture(), {});
  expect(missingPictureDraft.status).toBe("completed");
  expect(
    await prisma.submission.count({ where: { responseId: pictureResponseId } })
  ).toBe(0);
  const invalidPictureDocuments = [
    pictureDocumentFixture({
      images: [{ bytes: gifFixture, extension: "gif" }],
    }),
    pictureDocumentFixture({
      images: [
        { bytes: pngFixture(), extension: "png" },
        { bytes: jpegFixture(), extension: "jpg" },
      ],
    }),
    pictureDocumentFixture({
      images: [
        {
          bytes: pngFixture(1, 1, 10 * 1024 * 1024 + 1),
          extension: "png",
        },
      ],
    }),
    pictureDocumentFixture({
      images: [{ bytes: pngFixture(4097, 1), extension: "png" }],
    }),
  ];
  for (const invalidPictureDocument of invalidPictureDocuments) {
    const invalidPictureSave = await savePicture(invalidPictureDocument, {});
    expect(invalidPictureSave.status).toBe("failed");
    expect(
      await prisma.submission.count({
        where: { responseId: pictureResponseId },
      })
    ).toBe(0);
  }
  const placeholderPictureDocument = pictureDocumentFixture({
    images: [{ bytes: pngFixture(), extension: "png" }],
    showingPlaceholder: true,
  });
  const placeholderPictureDraft = await savePicture(
    placeholderPictureDocument,
    {}
  );
  expect(placeholderPictureDraft.status).toBe("completed");
  nextPictureDocument = placeholderPictureDocument;
  const placeholderPictureSubmitResponse = await pictureSubmitRequest({});
  expect(placeholderPictureSubmitResponse.status).toBe(202);
  const placeholderPictureSubmitBody =
    (await placeholderPictureSubmitResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !placeholderPictureSubmitBody.operationCapability ||
    !placeholderPictureSubmitBody.operationId
  ) {
    throw new Error("The placeholder picture submit operation was not created");
  }
  const placeholderPictureSubmitOperation = await waitForOperation(
    app,
    placeholderPictureSubmitBody.operationId,
    { "X-Editor-Capability": placeholderPictureSubmitBody.operationCapability }
  );
  expect(placeholderPictureSubmitOperation).toMatchObject({
    error: "invalid_template",
    status: "failed",
  });
  expect(
    await prisma.submission.count({ where: { responseId: pictureResponseId } })
  ).toBe(0);
  nextPictureDocument = pictureDocumentFixture();
  const missingPictureSubmitResponse = await pictureSubmitRequest({});
  expect(missingPictureSubmitResponse.status).toBe(202);
  const missingPictureSubmitBody =
    (await missingPictureSubmitResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !missingPictureSubmitBody.operationCapability ||
    !missingPictureSubmitBody.operationId
  ) {
    throw new Error("The missing picture submit operation was not created");
  }
  const missingPictureSubmitOperation = await waitForOperation(
    app,
    missingPictureSubmitBody.operationId,
    { "X-Editor-Capability": missingPictureSubmitBody.operationCapability }
  );
  expect(missingPictureSubmitOperation).toMatchObject({
    error: "invalid_template",
    status: "failed",
  });
  expect(
    await prisma.response.findUnique({
      select: { status: true },
      where: { id: pictureResponseId },
    })
  ).toMatchObject({ status: "draft" });
  expect(
    await prisma.submission.count({ where: { responseId: pictureResponseId } })
  ).toBe(0);
  nextPictureDocument = pictureDocumentFixture({
    images: [{ bytes: jpegFixture(), extension: "jpg" }],
  });
  const pictureSubmitResponse = await pictureSubmitRequest({
    photo: "scalar data must be omitted",
  });
  expect(pictureSubmitResponse.status).toBe(202);
  const pictureSubmitBody = (await pictureSubmitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !pictureSubmitBody.operationCapability ||
    !pictureSubmitBody.operationId
  ) {
    throw new Error("The picture submit operation was not created");
  }
  const pictureSubmitOperation = await waitForOperation(
    app,
    pictureSubmitBody.operationId,
    { "X-Editor-Capability": pictureSubmitBody.operationCapability }
  );
  expect(pictureSubmitOperation.status).toBe("completed");
  const pictureSubmission = await prisma.submission.findUniqueOrThrow({
    select: { data: true, objectKey: true, responseId: true },
    where: { responseId: pictureResponseId },
  });
  expect(pictureSubmission.data).toEqual({});
  expect(
    unzipSync(await readObject(pictureSubmission.objectKey))[
      "word/media/image1.jpg"
    ]
  ).toEqual(jpegFixture());
  const pictureCorrectionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${pictureResponseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(pictureCorrectionEditorResponse.status).toBe(200);
  const pictureCorrectionEditor =
    (await pictureCorrectionEditorResponse.json()) as EditorConfigBody;
  const pictureCorrectionCapability =
    pictureCorrectionEditor.bridge.capabilities["save-correction"];
  if (!pictureCorrectionCapability) {
    throw new Error("The picture correction capability was not returned");
  }
  const pictureCorrectionResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${pictureResponseId}/correction`,
      {
        body: JSON.stringify({
          data: {},
          documentKey: pictureCorrectionEditor.config.document.key,
          reason: "ตรวจสอบรูปภาพ",
        }),
        headers: capabilityHeaders(pictureCorrectionCapability),
        method: "POST",
      }
    )
  );
  expect(pictureCorrectionResponse.status).toBe(202);
  const pictureCorrectionBody = (await pictureCorrectionResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !pictureCorrectionBody.operationCapability ||
    !pictureCorrectionBody.operationId
  ) {
    throw new Error("The picture correction operation was not created");
  }
  const pictureCorrectionOperation = await waitForOperation(
    app,
    pictureCorrectionBody.operationId,
    {
      "X-Editor-Capability": pictureCorrectionBody.operationCapability,
    }
  );
  expect(pictureCorrectionOperation.status).toBe("completed");
  const pictureCorrection = await prisma.correction.findFirstOrThrow({
    select: { data: true, objectKey: true, revision: true },
    where: { responseId: pictureResponseId },
  });
  expect(pictureCorrection).toMatchObject({ data: {}, revision: 1 });
  expect(
    unzipSync(await readObject(pictureCorrection.objectKey))[
      "word/media/image1.jpg"
    ]
  ).toEqual(jpegFixture());
  await putObject(
    pictureCorrection.objectKey,
    pictureDocumentFixture(),
    DOCX_CONTENT_TYPE
  );
  const pictureRevisionsResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/responses/${pictureResponseId}/corrections`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(pictureRevisionsResponse.status).toBe(200);
  expect(await pictureRevisionsResponse.json()).toMatchObject({
    revisions: [
      { pictures: { photo: true }, revision: 0 },
      { pictures: { photo: false }, revision: 1 },
    ],
  });
  return {
    pictureCorrectionEditor,
    pictureResponseId,
  };
};
