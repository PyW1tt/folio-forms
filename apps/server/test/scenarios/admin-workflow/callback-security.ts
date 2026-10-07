import { expect } from "bun:test";
import { createHash } from "node:crypto";

import { prisma } from "@onlyoffice/db";
import { strToU8 } from "fflate";

import { createApp } from "../../../src/app";
import {
  createOnlyOfficeAuthorization,
  createCallbackUserdata,
  createOnlyOfficeBodyToken,
  verifyOnlyOfficeAuthorization,
} from "../../../src/onlyoffice";
import {
  readObject,
  objectExists,
  DOCX_CONTENT_TYPE,
  putObject,
  objectKey,
} from "../../../src/storage";
import { docxFixture, docxXmlFixture } from "../../fixtures/documents";
import {
  jsonHeaders,
  persistCallbackClaim,
  callbackPayload,
  tamperAuthorization,
} from "../../fixtures/http";
import type {
  EditorConfigBody,
  CallbackOperationFixture,
} from "../../fixtures/http";
import type { FieldConfigurationOutput } from "./editor-access";
import type { BootstrapAndCreationOutput } from "./setup";
import type { SubmissionExportsOutput } from "./submission";
import type { PrimaryPublicationOutput } from "./template-contract";

export interface CallbackSecurityInput {
  sourceDocument: SubmissionExportsOutput["sourceDocument"];
  adminId: BootstrapAndCreationOutput["adminId"];
  publishCapability: FieldConfigurationOutput["publishCapability"];
  formId: BootstrapAndCreationOutput["formId"];
  publishedManifestRecord: PrimaryPublicationOutput["publishedManifestRecord"];
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
}

export const runCallbackSecurity = async (
  input: CallbackSecurityInput
): Promise<void> => {
  const {
    sourceDocument,
    adminId,
    publishCapability,
    formId,
    publishedManifestRecord,
    formPublicId,
    adminBearer,
  } = input;

  const callbackFormId = crypto.randomUUID();
  const callbackFormPublicId = crypto.randomUUID().replaceAll("-", "");
  const callbackTemplateDocumentKey = `callback-template-${crypto.randomUUID()}`;
  const callbackTemplateObjectKey = objectKey(
    "forms",
    callbackFormId,
    "template-draft",
    "callback.docx"
  );
  await putObject(callbackTemplateObjectKey, sourceDocument, DOCX_CONTENT_TYPE);
  await prisma.form.create({
    data: {
      createdBy: adminId,
      description: "Callback operation test form",
      id: callbackFormId,
      publicId: callbackFormPublicId,
      templateDraft: {
        create: {
          contentHash: createHash("sha256")
            .update(sourceDocument)
            .digest("hex"),
          documentKey: callbackTemplateDocumentKey,
          id: crypto.randomUUID(),
          objectKey: callbackTemplateObjectKey,
        },
      },
      title: "Callback operation test form",
    },
  });
  const callbackTemplateDraft = await prisma.templateDraft.findUnique({
    where: { formId: callbackFormId },
  });
  if (!callbackTemplateDraft) {
    throw new Error("The callback template draft was not created");
  }

  const callbackDocument = docxFixture(
    `ticket-08-callback-${crypto.randomUUID()}`
  );
  const externalCallbackDocument = docxXmlFixture({
    additionalParts: {
      "word/_rels/document.xml.rels": strToU8(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="external" Target="http://127.0.0.1:80/" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/></Relationships>'
      ),
    },
    document:
      '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
  });
  const callbackMaximumBytes = Math.max(
    callbackDocument.byteLength,
    externalCallbackDocument.byteLength
  );
  const callbackDownloadPaths = new Set<string>();
  const callbackDocumentServer = Bun.serve({
    fetch(request) {
      const url = new URL(request.url);
      if (
        !verifyOnlyOfficeAuthorization(request.headers.get("authorization"), {
          url: request.url,
        })
      ) {
        return Response.json({ error: "unauthorized" }, { status: 401 });
      }
      callbackDownloadPaths.add(url.pathname);
      if (url.pathname === "/redirect.docx") {
        return new Response(null, {
          headers: { Location: `${url.origin}/ok.docx` },
          status: 302,
        });
      }
      if (url.pathname === "/large.docx") {
        return new Response(new Uint8Array(callbackMaximumBytes + 1));
      }
      if (url.pathname === "/malformed.docx") {
        return new Response("not a DOCX package");
      }
      if (url.pathname === "/external-relationship.docx") {
        return new Response(externalCallbackDocument, {
          headers: { "Content-Type": DOCX_CONTENT_TYPE },
        });
      }
      return new Response(callbackDocument, {
        headers: { "Content-Type": DOCX_CONTENT_TYPE },
      });
    },
    port: 0,
  });
  try {
    const callbackOrigin = callbackDocumentServer.url.origin;
    const callbackApp = createApp({
      onlyOffice: {
        convertDocxToPdf: () =>
          Promise.resolve(new TextEncoder().encode("%PDF-test")),
        forceSave: () => Promise.resolve(false),
      },
      onlyOfficeCallbackMaxBytes: callbackMaximumBytes,
      onlyOfficeCallbackOrigins: [callbackOrigin],
    });
    const callbackAppReplica = createApp({
      onlyOffice: {
        convertDocxToPdf: () =>
          Promise.resolve(new TextEncoder().encode("%PDF-test")),
        forceSave: () => Promise.resolve(false),
      },
      onlyOfficeCallbackMaxBytes: callbackMaximumBytes,
      onlyOfficeCallbackOrigins: [callbackOrigin],
    });
    const createCallbackOperation =
      async (): Promise<CallbackOperationFixture> => {
        const id = crypto.randomUUID();
        const stagedObjectKey = objectKey(
          "operations",
          id,
          "callback-staged.docx"
        );
        const finalObjectKey = objectKey(
          "operations",
          id,
          "callback-final.docx"
        );
        const userdata = createCallbackUserdata({
          documentKey: callbackTemplateDraft.documentKey,
          operationId: id,
          operationType: "save_template_draft",
        });
        await prisma.operation.create({
          data: {
            actorId: adminId,
            documentKey: callbackTemplateDraft.documentKey,
            errorCode: null,
            formId: callbackFormId,
            id,
            metadata: {
              action: "save-template",
              finalObjectKey,
              formId: callbackFormId,
              stagedObjectKey,
            },
            ownerUserId: adminId,
            stagingObjectKey: stagedObjectKey,
            status: "processing",
            targetId: id,
            targetType: "template_draft",
            type: "save_template_draft",
          },
        });
        await persistCallbackClaim({ operationId: id, userdata });
        return {
          documentKey: callbackTemplateDraft.documentKey,
          finalObjectKey,
          id,
          userdata,
        };
      };
    // oxlint-disable-next-line consistent-function-scoping -- Keep the original callback helper here; its default handler captures callbackApp.
    const postCallback = (
      payload: Record<string, unknown>,
      authorization = createOnlyOfficeAuthorization(payload),
      bodyToken = createOnlyOfficeBodyToken(payload),
      callbackHandler = callbackApp
    ): Promise<Response> =>
      callbackHandler.handle(
        new Request("http://test.local/onlyoffice/callback", {
          body: JSON.stringify({ ...payload, token: bodyToken }),
          headers: {
            Authorization: authorization,
            ...jsonHeaders,
          },
          method: "POST",
        })
      );

    const trustOperation = await createCallbackOperation();
    const persistedTrustClaim = await prisma.callbackClaim.findUnique({
      where: { operationId: trustOperation.id },
    });
    if (!persistedTrustClaim) {
      throw new Error("The callback claim was not persisted");
    }
    expect(persistedTrustClaim.tokenDigest).toMatch(/^[0-9a-f]{64}$/u);
    expect(persistedTrustClaim.tokenDigest).toBe(
      createHash("sha256").update(trustOperation.userdata).digest("hex")
    );
    const trustedPayload = callbackPayload(
      trustOperation,
      `${callbackOrigin}/ok.docx`
    );
    const invalidJwtResponse = await postCallback(
      trustedPayload,
      tamperAuthorization(createOnlyOfficeAuthorization(trustedPayload))
    );
    expect(invalidJwtResponse.status).toBe(401);
    const invalidBodyTokenResponse = await postCallback(
      trustedPayload,
      createOnlyOfficeAuthorization(trustedPayload),
      `${createOnlyOfficeBodyToken(trustedPayload)}x`
    );
    expect(invalidBodyTokenResponse.status).toBe(401);
    const ordinaryCallbackResponse = await postCallback({
      actions: [],
      key: trustOperation.documentKey,
      status: 1,
    });
    expect(await ordinaryCallbackResponse.json()).toEqual({ error: 0 });
    const browserSessionBoundaryResponse = await postCallback(
      trustedPayload,
      `Bearer ${publishCapability}`
    );
    expect(browserSessionBoundaryResponse.status).toBe(401);

    const wrongKeyPayload = {
      ...trustedPayload,
      key: `wrong-${trustOperation.documentKey}`,
    };
    const wrongKeyResponse = await postCallback(wrongKeyPayload);
    expect(await wrongKeyResponse.json()).toEqual({ error: 1 });
    const wrongOperationPayload = {
      ...trustedPayload,
      userdata: createCallbackUserdata({
        documentKey: trustOperation.documentKey,
        operationId: crypto.randomUUID(),
        operationType: "save_template_draft",
      }),
    };
    const wrongOperationResponse = await postCallback(wrongOperationPayload);
    expect(await wrongOperationResponse.json()).toEqual({ error: 1 });
    const wrongOperationTypePayload = {
      ...trustedPayload,
      userdata: createCallbackUserdata({
        documentKey: trustOperation.documentKey,
        operationId: trustOperation.id,
        operationType: "submit_response",
      }),
    };
    const wrongOperationTypeResponse = await postCallback(
      wrongOperationTypePayload
    );
    expect(await wrongOperationTypeResponse.json()).toEqual({ error: 1 });
    const expiredCallbackPayload = {
      ...trustedPayload,
      userdata: createCallbackUserdata({
        documentKey: trustOperation.documentKey,
        expiresAt: Math.floor(Date.now() / 1000) - 1,
        operationId: trustOperation.id,
        operationType: "save_template_draft",
      }),
    };
    const expiredCallbackResponse = await postCallback(expiredCallbackPayload);
    expect(await expiredCallbackResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: trustOperation.id },
      })
    ).toEqual({ status: "processing" });

    const oversizedCallbackPayload = {
      padding: "x".repeat(65 * 1024),
      status: 6,
    };
    const oversizedCallbackResponse = await postCallback(
      oversizedCallbackPayload
    );
    expect(oversizedCallbackResponse.status).toBe(413);

    const forbiddenOriginOperation = await createCallbackOperation();
    const forbiddenOriginResponse = await postCallback(
      callbackPayload(
        forbiddenOriginOperation,
        "https://attacker.example/document.docx"
      )
    );
    expect(await forbiddenOriginResponse.json()).toEqual({ error: 0 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: forbiddenOriginOperation.id },
      })
    ).toEqual({ status: "failed" });

    const redirectOperation = await createCallbackOperation();
    const redirectResponse = await postCallback(
      callbackPayload(redirectOperation, `${callbackOrigin}/redirect.docx`)
    );
    expect(await redirectResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: redirectOperation.id },
      })
    ).toEqual({ status: "failed" });

    const oversizedDocumentOperation = await createCallbackOperation();
    const oversizedDocumentResponse = await postCallback(
      callbackPayload(
        oversizedDocumentOperation,
        `${callbackOrigin}/large.docx`
      )
    );
    expect(await oversizedDocumentResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: oversizedDocumentOperation.id },
      })
    ).toEqual({ status: "failed" });
    const templateBeforeMalformedSave = await prisma.templateDraft.findUnique({
      select: { documentKey: true, objectKey: true },
      where: { formId: callbackFormId },
    });
    if (!templateBeforeMalformedSave) {
      throw new Error("The Template Draft rollback baseline was not found");
    }
    const templateBytesBeforeMalformedSave = await readObject(
      templateBeforeMalformedSave.objectKey
    );
    const malformedSaveOperation = await createCallbackOperation();
    const malformedSaveResponse = await postCallback(
      callbackPayload(
        malformedSaveOperation,
        `${callbackOrigin}/malformed.docx`
      )
    );
    expect(await malformedSaveResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: malformedSaveOperation.id },
      })
    ).toEqual({ status: "failed" });
    expect(
      await prisma.templateDraft.findUnique({
        select: { documentKey: true, objectKey: true },
        where: { formId: callbackFormId },
      })
    ).toEqual(templateBeforeMalformedSave);
    expect(await readObject(templateBeforeMalformedSave.objectKey)).toEqual(
      templateBytesBeforeMalformedSave
    );
    const externalResponseId = crypto.randomUUID();
    const externalResponseDocumentKey = `external-response-${crypto.randomUUID()}`;
    const externalResponseObjectKey = objectKey(
      "responses",
      externalResponseId,
      "draft",
      "baseline.docx"
    );
    await putObject(
      externalResponseObjectKey,
      callbackDocument,
      DOCX_CONTENT_TYPE
    );
    await prisma.response.create({
      data: {
        draftData: {},
        draftDocumentKey: externalResponseDocumentKey,
        draftObjectKey: externalResponseObjectKey,
        formId,
        id: externalResponseId,
        publishedTemplateId: publishedManifestRecord.id,
        publishedVersion: publishedManifestRecord.version,
        status: "draft",
        userId: adminId,
      },
    });
    const externalResponseOperationId = crypto.randomUUID();
    const externalResponseStagedObjectKey = objectKey(
      "operations",
      externalResponseOperationId,
      "external-response-staged.docx"
    );
    const externalResponseFinalObjectKey = objectKey(
      "operations",
      externalResponseOperationId,
      "external-response-final.docx"
    );
    const externalResponseUserdata = createCallbackUserdata({
      documentKey: externalResponseDocumentKey,
      operationId: externalResponseOperationId,
      operationType: "save_draft",
    });
    await prisma.operation.create({
      data: {
        actorId: adminId,
        documentKey: externalResponseDocumentKey,
        errorCode: null,
        formId,
        id: externalResponseOperationId,
        metadata: {
          action: "save-draft",
          data: {},
          finalObjectKey: externalResponseFinalObjectKey,
          formId,
          publicId: formPublicId,
          responseId: externalResponseId,
          stagedObjectKey: externalResponseStagedObjectKey,
        },
        ownerUserId: adminId,
        responseId: externalResponseId,
        stagingObjectKey: externalResponseStagedObjectKey,
        status: "processing",
        targetId: externalResponseId,
        targetType: "response",
        type: "save_draft",
      },
    });
    await persistCallbackClaim({
      operationId: externalResponseOperationId,
      userdata: externalResponseUserdata,
    });
    const externalResponseOperation: CallbackOperationFixture = {
      documentKey: externalResponseDocumentKey,
      finalObjectKey: externalResponseFinalObjectKey,
      id: externalResponseOperationId,
      userdata: externalResponseUserdata,
    };
    const externalResponseBefore = await prisma.response.findUnique({
      select: {
        draftData: true,
        draftDocumentKey: true,
        draftObjectKey: true,
        status: true,
      },
      where: { id: externalResponseId },
    });
    const externalResponseCallback = await postCallback(
      callbackPayload(
        externalResponseOperation,
        `${callbackOrigin}/external-relationship.docx`
      )
    );
    expect(await externalResponseCallback.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { errorCode: true, status: true },
        where: { id: externalResponseOperationId },
      })
    ).toEqual({
      errorCode: "invalid_template",
      status: "failed",
    });
    expect(
      await prisma.response.findUnique({
        select: {
          draftData: true,
          draftDocumentKey: true,
          draftObjectKey: true,
          status: true,
        },
        where: { id: externalResponseId },
      })
    ).toEqual(externalResponseBefore);
    expect(await objectExists(externalResponseObjectKey)).toBe(true);
    expect(await objectExists(externalResponseStagedObjectKey)).toBe(false);
    expect(await objectExists(externalResponseFinalObjectKey)).toBe(false);
    await prisma.operation.delete({
      where: { id: externalResponseOperationId },
    });
    await prisma.response.delete({ where: { id: externalResponseId } });

    const validCallbackResponse = await postCallback(trustedPayload);
    expect(await validCallbackResponse.json()).toEqual({ error: 0 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: trustOperation.id },
      })
    ).toEqual({ status: "completed" });
    expect(await readObject(trustOperation.finalObjectKey)).toEqual(
      callbackDocument
    );
    const consumedTrustClaim = await prisma.callbackClaim.findUnique({
      where: { operationId: trustOperation.id },
    });
    expect(consumedTrustClaim?.consumedAt).not.toBeNull();
    const changedEditorResponse = await callbackApp.handle(
      new Request(
        `http://test.local/api/admin/forms/${callbackFormPublicId}/editor-config`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    expect(changedEditorResponse.status).toBe(200);
    const changedEditor =
      (await changedEditorResponse.json()) as EditorConfigBody;
    expect(JSON.stringify(changedEditor)).not.toContain(callbackFormId);
    const changedDocumentUrl = changedEditor.config.document.url;
    const changedDocumentResponse = await callbackApp.handle(
      new Request(changedDocumentUrl, {
        headers: {
          Authorization: createOnlyOfficeAuthorization({
            url: changedDocumentUrl,
          }),
        },
      })
    );
    expect(changedDocumentResponse.status).toBe(200);
    expect(Buffer.from(await changedDocumentResponse.arrayBuffer())).toEqual(
      Buffer.from(callbackDocument)
    );

    const concurrentCallbackOperation = await createCallbackOperation();
    const concurrentCallbackPayload = callbackPayload(
      concurrentCallbackOperation,
      `${callbackOrigin}/ok.docx`
    );
    const [concurrentCallbackA, concurrentCallbackB] = await Promise.all([
      postCallback(concurrentCallbackPayload),
      postCallback(
        concurrentCallbackPayload,
        undefined,
        undefined,
        callbackAppReplica
      ),
    ]);
    expect(await concurrentCallbackA.json()).toEqual({ error: 0 });
    expect(await concurrentCallbackB.json()).toEqual({ error: 0 });
    const concurrentCallbackState = await prisma.operation.findUnique({
      select: { result: true, status: true },
      where: { id: concurrentCallbackOperation.id },
    });
    expect(concurrentCallbackState).toMatchObject({ status: "completed" });
    const concurrentCallbackClaim = await prisma.callbackClaim.findUnique({
      where: { operationId: concurrentCallbackOperation.id },
    });
    expect(concurrentCallbackClaim?.consumedAt).not.toBeNull();
    expect(
      await readObject(concurrentCallbackOperation.finalObjectKey)
    ).toEqual(callbackDocument);
    const replayedCallbackResponse = await postCallback(
      concurrentCallbackPayload
    );
    expect(await replayedCallbackResponse.json()).toEqual({ error: 0 });
    expect(
      await prisma.operation.findUnique({
        select: { result: true, status: true },
        where: { id: concurrentCallbackOperation.id },
      })
    ).toEqual(concurrentCallbackState);

    expect(callbackDownloadPaths).toEqual(
      new Set([
        "/external-relationship.docx",
        "/large.docx",
        "/malformed.docx",
        "/ok.docx",
        "/redirect.docx",
      ])
    );
  } finally {
    callbackDocumentServer.stop(true);
  }
};
