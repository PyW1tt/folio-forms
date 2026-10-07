import { test, expect, vi } from "bun:test";

import { AgentSession } from "@earendil-works/pi-coding-agent";
import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";

import { createApp } from "../../src/app";
import { DOCX_CONTENT_TYPE, objectExists, readObject } from "../../src/storage";
import {
  temporaryDirectories,
  jsonHeaders,
  createCredentialFixture,
  bearerFor,
  formCreationRequest,
} from "../fixtures/http";

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

test("Ticket 17/18 AI Authoring creates, revises, restores, and uploads a validated DOCX", async () => {
  const model = "ticket-17-local-model";
  const serviceKey = `ticket-17-local-only-${crypto.randomUUID()}`;
  const generatedTemplate = {
    description: "Request equipment for work.",
    fields: [
      {
        label: "Employee Name",
        placeholder: "Enter employee name",
        tag: "employee_name",
      },
    ],
    paragraphs: ["Complete each field."],
    title: "Equipment Request",
  };
  const revisedTemplate = {
    ...generatedTemplate,
    fields: [
      ...generatedTemplate.fields,
      {
        label: "Department",
        placeholder: "Enter department",
        tag: "department",
      },
    ],
    paragraphs: [...generatedTemplate.paragraphs, "Route through Facilities."],
  };
  const omittedTemplate = {
    ...revisedTemplate,
    paragraphs: ["Route through Facilities.", "Send to manager for approval."],
  };
  const revisionPrompt = "Add a Department field and a Facilities routing note";
  const failedRevisionPrompt =
    "Add an approval note below the existing paragraphs";
  let revisionCalls = 0;
  let upstreamCalls = 0;
  let forcedToolCalls: boolean[] = [];
  let omitAssistantSummary = false;
  let blockNextProviderRequest = false;
  const { promise: providerBlocked, resolve: notifyProviderBlocked } =
    Promise.withResolvers<undefined>();
  const { promise: providerResponseGate, resolve: releaseProviderResponse } =
    Promise.withResolvers<undefined>();
  const { promise: providerAborted, resolve: notifyProviderAborted } =
    Promise.withResolvers<undefined>();
  let blockNextPromptPreflight = false;
  const {
    promise: promptPreflightPaused,
    resolve: notifyPromptPreflightPaused,
  } = Promise.withResolvers<undefined>();
  const { promise: promptPreflightGate, resolve: releasePromptPreflight } =
    Promise.withResolvers<undefined>();
  const { promise: promptAbortRequested, resolve: notifyPromptAbortRequested } =
    Promise.withResolvers<undefined>();
  let restoreAgentSessionHooks: (() => void) | undefined;
  const upstream = Bun.serve({
    fetch: async (request) => {
      expect(request.method).toBe("POST");
      expect(new URL(request.url).pathname).toBe("/v1/chat/completions");
      expect(request.headers.get("authorization")).toBe(`Bearer ${serviceKey}`);
      const requestBody = (await request.json()) as {
        messages?: unknown;
        model?: unknown;
        tools?: { function?: { name?: unknown } }[];
      };
      const messages = JSON.stringify(requestBody.messages ?? []);
      const isRevision =
        messages.includes(revisionPrompt) ||
        messages.includes(failedRevisionPrompt);
      if (isRevision) {
        revisionCalls += 1;
        expect(messages).toContain("Create an equipment request form");
        expect(messages).toContain("Complete each field.");
        expect(messages).toContain("employee_name");
        expect(messages).toContain(
          "Created Equipment Request with 1 tagged fields."
        );
        expect(messages).toContain(
          "Created Equipment Request with one tagged field."
        );
        expect(messages).toContain("CURRENT document:");
        if (messages.includes(failedRevisionPrompt)) {
          expect(messages).toContain("Route through Facilities.");
          expect(messages).toContain("department");
        }
      } else {
        upstreamCalls += 1;
        expect(messages).toContain("Create an equipment request form");
      }
      expect(requestBody.model).toBe(model);
      expect(
        requestBody.tools?.some(
          (tool) => tool.function?.name === "create_template_docx"
        )
      ).toBe(true);
      if (blockNextProviderRequest) {
        blockNextProviderRequest = false;
        request.signal.addEventListener(
          "abort",
          () => notifyProviderAborted(),
          { once: true }
        );
        notifyProviderBlocked();
        await providerResponseGate;
      }
      const toolCall =
        forcedToolCalls.shift() ??
        (isRevision ? revisionCalls % 2 === 1 : upstreamCalls % 2 === 1);
      const template = isRevision ? revisedTemplate : generatedTemplate;
      const toolArguments = messages.includes(failedRevisionPrompt)
        ? omittedTemplate
        : template;
      let completionChunks;
      if (toolCall) {
        completionChunks = [
          {
            delta: {
              role: "assistant",
              tool_calls: [
                {
                  function: {
                    arguments: JSON.stringify(toolArguments),
                    name: "create_template_docx",
                  },
                  id: `call-${upstreamCalls}-${revisionCalls}`,
                  index: 0,
                  type: "function",
                },
              ],
            },
            finish_reason: null,
          },
          { delta: {}, finish_reason: "tool_calls" },
        ];
      } else if (omitAssistantSummary) {
        completionChunks = [
          { delta: { role: "assistant" }, finish_reason: "stop" },
        ];
      } else {
        completionChunks = [
          {
            delta: {
              content: isRevision
                ? "Updated Equipment Request with Department and Facilities routing."
                : "Created Equipment Request with one tagged field.",
              role: "assistant",
            },
            finish_reason: null,
          },
          { delta: {}, finish_reason: "stop" },
        ];
      }
      const createdAt = Math.floor(Date.now() / 1000);
      const stream = completionChunks
        .map((chunk) =>
          [
            "data: ",
            JSON.stringify({
              choices: [{ index: 0, ...chunk }],
              created: createdAt,
              id: `chatcmpl-${upstreamCalls}-${revisionCalls}`,
              model,
              object: "chat.completion.chunk",
            }),
            "\n\n",
          ].join("")
        )
        .join("");
      return new Response(`${stream}data: [DONE]\n\n`, {
        headers: { "Content-Type": "text/event-stream" },
      });
    },
    hostname: "127.0.0.1",
    port: 0,
  });
  let now = new Date();
  const aiApp = createApp({
    clock: () => new Date(now),
    legacySso: null,
    omniRoute: {
      apiKey: serviceKey,
      baseUrl: `http://127.0.0.1:${upstream.port}/v1`,
      model,
    },
    onlyOffice: {
      convertDocxToPdf: () => Promise.resolve(new Uint8Array()),
      forceSave: () => Promise.resolve(false),
    },
    prefillReturnUrl: "https://source.example.test/forms/return",
  });
  const disabledApp = createApp({ legacySso: null, omniRoute: null });
  const createPreview = async (authorization: string) => {
    const directoriesBefore = new Set(await temporaryDirectories());
    const response = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: true,
          prompt: "Create an equipment request form",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${authorization}`,
        },
        method: "POST",
      })
    );
    expect(response.status).toBe(200);
    const result = (await response.json()) as {
      session: {
        assistantMessage: string;
        description: string;
        downloadUrl: string;
        fields: { label: string; placeholder: string; tag: string }[];
        paragraphs: string[];
        turns: { assistantMessage: string; prompt: string }[];
        sessionId: string;
        title: string;
      };
    };
    const directoriesAfterCreate = await temporaryDirectories();
    const directoriesAdded = directoriesAfterCreate.filter(
      (directory) => !directoriesBefore.has(directory)
    );
    expect(directoriesAdded).toHaveLength(1);
    const [directory] = directoriesAdded;
    if (!directory) {
      throw new Error("AI Authoring did not create its temporary directory");
    }
    expect(result.session.sessionId).toMatch(
      /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/iu
    );
    expect(result.session.downloadUrl).toBe(
      `/api/admin/ai-authoring/sessions/${result.session.sessionId}/docx`
    );
    expect(Object.hasOwn(result.session, "prompt")).toBe(false);
    return { directory, session: result.session };
  };
  const adminPassword = "Ticket17-admin-password";
  const otherAdminPassword = "Ticket17-other-admin-password";
  const userPassword = "Ticket17-user-password";
  const delayedBodyPassword = "Ticket17-delayed-body-password";
  const changedDelayedBodyPassword = "Ticket17-delayed-body-updated-password";
  const changedOtherAdminPassword = "Ticket17-other-admin-updated-password";
  let adminBearer = "";
  let otherAdminBearer = "";
  let userBearer = "";
  let mutationActorBearer = "";
  let sameUserOtherBearer = "";
  let deletionActorBearer = "";
  let cleanupTargetBearer = "";
  let activeSessionId: string | undefined;
  let formPublicId: string | undefined;
  let signedOut = false;
  try {
    const admin = await createCredentialFixture({
      email: `ticket-17-admin-${crypto.randomUUID()}@example.com`,
      name: "Ticket 17 Admin",
      password: adminPassword,
      role: "admin",
    });
    const otherAdmin = await createCredentialFixture({
      email: `ticket-17-other-admin-${crypto.randomUUID()}@example.com`,
      name: "Ticket 17 Other Admin",
      password: otherAdminPassword,
      role: "admin",
    });
    const user = await createCredentialFixture({
      email: `ticket-17-user-${crypto.randomUUID()}@example.com`,
      name: "Ticket 17 User",
      password: userPassword,
    });

    const delayedBodyAdmin = await createCredentialFixture({
      email: `ticket-17-delayed-body-${crypto.randomUUID()}@example.com`,
      name: "Ticket 17 Delayed Body Admin",
      password: delayedBodyPassword,
      role: "admin",
    });
    adminBearer = await bearerFor(app, admin.email, adminPassword);
    otherAdminBearer = await bearerFor(
      app,
      otherAdmin.email,
      otherAdminPassword
    );
    userBearer = await bearerFor(app, user.email, userPassword);
    const delayedBodyBearer = await bearerFor(
      app,
      delayedBodyAdmin.email,
      delayedBodyPassword
    );
    const { promise: bodyRead, resolve: notifyBodyRead } =
      Promise.withResolvers<undefined>();
    const { promise: bodyGate, resolve: releaseBody } =
      Promise.withResolvers<undefined>();
    let bodyReadNotified = false;
    let delayedCreate: Promise<globalThis.Response> | undefined;
    try {
      const requestBody = new ReadableStream<Uint8Array>(
        {
          pull(controller) {
            if (!bodyReadNotified) {
              bodyReadNotified = true;
              notifyBodyRead();
            }
            return bodyGate.then(() => {
              controller.enqueue(
                new TextEncoder().encode(
                  JSON.stringify({
                    consent: true,
                    prompt: "Create an equipment request form",
                  })
                )
              );
              controller.close();
            });
          },
        },
        { highWaterMark: 0 }
      );
      const delayedRequestInit = {
        body: requestBody,
        duplex: "half" as const,
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${delayedBodyBearer}`,
        },
        method: "POST",
      };
      delayedCreate = aiApp.handle(
        new Request(
          "http://test.local/api/admin/ai-authoring/sessions",
          delayedRequestInit
        )
      );
      await bodyRead;
      const callsBeforeSignOut = upstreamCalls;
      const delayedSignOut = await aiApp.handle(
        new Request("http://test.local/api/auth/sign-out", {
          headers: { Authorization: `Bearer ${delayedBodyBearer}` },
          method: "POST",
        })
      );
      expect(delayedSignOut.status).toBe(200);
      releaseBody();
      const staleCreate = await delayedCreate;
      expect(staleCreate.status).toBe(404);
      expect(upstreamCalls).toBe(callsBeforeSignOut);
      const delayedBodySession = await aiApp.handle(
        new Request("http://test.local/api/session", {
          headers: { Authorization: `Bearer ${delayedBodyBearer}` },
        })
      );
      expect(delayedBodySession.status).toBe(401);
    } finally {
      await aiApp
        .handle(
          new Request("http://test.local/api/auth/sign-out", {
            headers: { Authorization: `Bearer ${delayedBodyBearer}` },
            method: "POST",
          })
        )
        .catch(() => {});
      releaseBody();
      await delayedCreate?.catch(() => {});
    }
    const unauthenticatedStatus = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring")
    );
    expect(unauthenticatedStatus.status).toBe(401);
    const userStatus = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring", {
        headers: { Authorization: `Bearer ${userBearer}` },
      })
    );
    expect(userStatus.status).toBe(403);
    const disabledStatus = await disabledApp.handle(
      new Request("http://test.local/api/admin/ai-authoring", {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(disabledStatus.status).toBe(200);
    expect(await disabledStatus.json()).toMatchObject({ enabled: false });
    const disabledCreate = await disabledApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: true,
          prompt: "Create an equipment request form",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    expect(disabledCreate.status).toBe(503);
    expect(upstreamCalls).toBe(0);
    const adminStatus = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring", {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(adminStatus.status).toBe(200);
    const statusBody = (await adminStatus.json()) as {
      disclosure: string;
      enabled: boolean;
    };
    expect(statusBody.enabled).toBe(true);
    expect(statusBody.disclosure).toContain(
      "cannot promise deletion by OmniRoute"
    );
    expect(JSON.stringify(statusBody)).not.toContain(serviceKey);
    const userCreate = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: true,
          prompt: "Create an equipment request form",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${userBearer}`,
        },
        method: "POST",
      })
    );
    expect(userCreate.status).toBe(403);
    const noConsent = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: false,
          prompt: "Create an equipment request form",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    expect(noConsent.status).toBe(428);
    const providerSelection = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: true,
          prompt: "Create an equipment request form",
          provider: "client-selected-provider",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    expect(providerSelection.status).toBe(400);
    expect(upstreamCalls).toBe(0);

    const firstDirectories = new Set(await temporaryDirectories());
    const first = await createPreview(adminBearer);
    activeSessionId = first.session.sessionId;
    expect(first.session.title).toBe(generatedTemplate.title);
    expect(first.session.description).toBe(generatedTemplate.description);
    expect(first.session.paragraphs).toEqual(generatedTemplate.paragraphs);
    expect(first.session.fields).toEqual(generatedTemplate.fields);
    expect(first.session.assistantMessage).toContain(
      "Created Equipment Request"
    );
    expect(await temporaryDirectories()).toContain(first.directory);
    expect(firstDirectories.has(first.directory)).toBe(false);
    expect(upstreamCalls).toBe(2);
    expect(first.session.turns).toEqual([
      {
        assistantMessage: first.session.assistantMessage,
        prompt: "Create an equipment request form",
      },
    ]);
    const currentUrl =
      "http://test.local/api/admin/ai-authoring/sessions/current";
    const revisionUrl = `http://test.local/api/admin/ai-authoring/sessions/${first.session.sessionId}/revisions`;
    const currentBeforeRevision = await aiApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(currentBeforeRevision.status).toBe(200);
    expect(await currentBeforeRevision.json()).toEqual({
      session: first.session,
    });
    sameUserOtherBearer = await bearerFor(app, admin.email, adminPassword);
    for (const authorization of [sameUserOtherBearer, otherAdminBearer]) {
      const otherCurrent = await aiApp.handle(
        new Request(currentUrl, {
          headers: { Authorization: `Bearer ${authorization}` },
        })
      );
      expect(otherCurrent.status).toBe(200);
      expect(await otherCurrent.json()).toEqual({ session: null });
      const otherRevision = await aiApp.handle(
        new Request(revisionUrl, {
          body: JSON.stringify({ consent: true, prompt: revisionPrompt }),
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${authorization}`,
          },
          method: "POST",
        })
      );
      expect(otherRevision.status).toBe(404);
    }
    const userCurrent = await aiApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${userBearer}` },
      })
    );
    expect(userCurrent.status).toBe(403);
    const anonymousCurrent = await aiApp.handle(new Request(currentUrl));
    expect(anonymousCurrent.status).toBe(401);
    const duplicateCreate = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: true,
          prompt: "Create another equipment request form",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    expect(duplicateCreate.status).toBe(409);
    expect(upstreamCalls).toBe(2);

    const unauthenticatedDownload = await aiApp.handle(
      new Request(`http://test.local${first.session.downloadUrl}`)
    );
    expect(unauthenticatedDownload.status).toBe(401);
    const userDownload = await aiApp.handle(
      new Request(`http://test.local${first.session.downloadUrl}`, {
        headers: { Authorization: `Bearer ${userBearer}` },
      })
    );
    expect(userDownload.status).toBe(403);
    const otherAdminDownload = await aiApp.handle(
      new Request(`http://test.local${first.session.downloadUrl}`, {
        headers: { Authorization: `Bearer ${otherAdminBearer}` },
      })
    );
    expect(otherAdminDownload.status).toBe(404);
    const otherAdminDelete = await aiApp.handle(
      new Request(
        `http://test.local/api/admin/ai-authoring/sessions/${first.session.sessionId}`,
        {
          headers: { Authorization: `Bearer ${otherAdminBearer}` },
          method: "DELETE",
        }
      )
    );
    expect(otherAdminDelete.status).toBe(404);
    const downloadResponse = await aiApp.handle(
      new Request(`http://test.local${first.session.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(downloadResponse.status).toBe(200);
    expect(downloadResponse.headers.get("content-type")).toBe(
      DOCX_CONTENT_TYPE
    );
    expect(downloadResponse.headers.get("cache-control")).toBe(
      "private, no-store"
    );
    const initialDocumentBytes = new Uint8Array(
      await downloadResponse.arrayBuffer()
    );
    const archive = unzipSync(initialDocumentBytes);
    const documentXmlBytes = archive["word/document.xml"];
    if (!documentXmlBytes) {
      throw new Error("AI Authoring did not return a DOCX document part");
    }
    const documentXml = new TextDecoder().decode(documentXmlBytes);
    expect(documentXml).toContain("Equipment Request");
    expect(documentXml).toContain("Request equipment for work.");
    expect(documentXml).toContain("Complete each field.");
    expect(documentXml).toContain("<w:sdt>");
    expect(documentXml).toContain('<w:tag w:val="employee_name"/>');
    expect(documentXml).toContain("Enter employee name");
    const noRevisionConsent = await aiApp.handle(
      new Request(revisionUrl, {
        body: JSON.stringify({ consent: false, prompt: revisionPrompt }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    expect(noRevisionConsent.status).toBe(428);
    const revisedResponse = await aiApp.handle(
      new Request(revisionUrl, {
        body: JSON.stringify({ consent: true, prompt: revisionPrompt }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    expect(revisedResponse.status).toBe(200);
    const { session: revised } = (await revisedResponse.json()) as {
      session: typeof first.session;
    };
    expect(revisionCalls).toBe(2);
    expect(revised.sessionId).toBe(first.session.sessionId);
    expect(revised.downloadUrl).toBe(first.session.downloadUrl);
    expect(revised.fields).toEqual(revisedTemplate.fields);
    expect(revised.paragraphs).toEqual(revisedTemplate.paragraphs);
    expect(revised.turns).toEqual([
      {
        assistantMessage: first.session.assistantMessage,
        prompt: "Create an equipment request form",
      },
      {
        assistantMessage:
          "Updated Equipment Request with Department and Facilities routing.",
        prompt: revisionPrompt,
      },
    ]);
    const restoredResponse = await aiApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(restoredResponse.status).toBe(200);
    expect(await restoredResponse.json()).toEqual({ session: revised });
    expect(await temporaryDirectories()).toContain(first.directory);
    const revisedDownload = await aiApp.handle(
      new Request(`http://test.local${revised.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(revisedDownload.status).toBe(200);
    const documentBytes = new Uint8Array(await revisedDownload.arrayBuffer());
    expect(documentBytes).not.toEqual(initialDocumentBytes);
    const revisedXmlBytes = unzipSync(documentBytes)["word/document.xml"];
    if (!revisedXmlBytes) {
      throw new Error("Revised AI Authoring DOCX has no document part");
    }
    const revisedXml = new TextDecoder().decode(revisedXmlBytes);
    expect(revisedXml).toContain("Complete each field.");
    expect(revisedXml).toContain('<w:tag w:val="employee_name"/>');
    expect(revisedXml).toContain("Route through Facilities.");
    expect(revisedXml).toContain('<w:tag w:val="department"/>');
    expect(revisedXml).toContain("Enter department");
    const failedRevision = await aiApp.handle(
      new Request(revisionUrl, {
        body: JSON.stringify({ consent: true, prompt: failedRevisionPrompt }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    expect(failedRevision.status).toBe(502);
    const afterFailedRevision = await aiApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(await afterFailedRevision.json()).toEqual({ session: revised });
    const retainedDownload = await aiApp.handle(
      new Request(`http://test.local${revised.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(retainedDownload.status).toBe(200);
    expect(new Uint8Array(await retainedDownload.arrayBuffer())).toEqual(
      documentBytes
    );
    expect(await temporaryDirectories()).toContain(first.directory);

    const uploadResponse = await aiApp.handle(
      formCreationRequest({
        authorization: adminBearer,
        description: first.session.description,
        source: "upload",
        template: {
          bytes: documentBytes,
          name: "equipment-request.docx",
        },
        title: first.session.title,
      })
    );
    expect(uploadResponse.status).toBe(200);
    const uploadBody = (await uploadResponse.json()) as {
      form: { hasTemplateDraft: boolean; publicId: string; title: string };
    };
    formPublicId = uploadBody.form.publicId;
    expect(uploadBody.form.title).toBe(generatedTemplate.title);
    expect(uploadBody.form.hasTemplateDraft).toBe(true);
    const storedForm = await prisma.form.findUnique({
      include: { templateDraft: true },
      where: { publicId: uploadBody.form.publicId },
    });
    if (!storedForm?.templateDraft) {
      throw new Error("The upload did not create a normal Template Draft");
    }
    expect(await objectExists(storedForm.templateDraft.objectKey)).toBe(true);
    expect(await readObject(storedForm.templateDraft.objectKey)).toEqual(
      documentBytes
    );
    const deleteForm = await aiApp.handle(
      new Request(
        `http://test.local/api/admin/forms/${uploadBody.form.publicId}`,
        {
          headers: { Authorization: `Bearer ${adminBearer}` },
          method: "DELETE",
        }
      )
    );
    expect(deleteForm.status).toBe(200);
    formPublicId = undefined;
    const endFirstSession = await aiApp.handle(
      new Request(
        `http://test.local/api/admin/ai-authoring/sessions/${first.session.sessionId}`,
        {
          headers: { Authorization: `Bearer ${adminBearer}` },
          method: "DELETE",
        }
      )
    );
    expect(endFirstSession.status).toBe(200);
    activeSessionId = undefined;
    expect(await temporaryDirectories()).not.toContain(first.directory);
    const endedDownload = await aiApp.handle(
      new Request(`http://test.local${first.session.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(endedDownload.status).toBe(404);
    const endedCurrent = await aiApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(await endedCurrent.json()).toEqual({ session: null });

    omitAssistantSummary = true;
    const second = await createPreview(adminBearer);
    omitAssistantSummary = false;
    expect(second.session.assistantMessage).toBe(
      "Created Equipment Request with 1 tagged field."
    );
    activeSessionId = second.session.sessionId;
    now = new Date(now.getTime() + 2 * 60 * 60 * 1000 - 1);
    const sessionStillActive = await aiApp.handle(
      new Request(`http://test.local${second.session.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(sessionStillActive.status).toBe(200);
    await sessionStillActive.arrayBuffer();
    now = new Date(now.getTime() + 2 * 60 * 60 * 1000);
    const expiredSession = await aiApp.handle(
      new Request(`http://test.local${second.session.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(expiredSession.status).toBe(404);
    activeSessionId = undefined;
    expect(await temporaryDirectories()).not.toContain(second.directory);
    const expiredCurrent = await aiApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(await expiredCurrent.json()).toEqual({ session: null });

    const third = await createPreview(adminBearer);
    activeSessionId = third.session.sessionId;
    const signOutWorkspace = await aiApp.handle(
      new Request("http://test.local/api/auth/sign-out", {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      })
    );
    expect(signOutWorkspace.status).toBe(200);
    signedOut = true;
    activeSessionId = undefined;
    expect(await temporaryDirectories()).not.toContain(third.directory);
    const signedOutCurrent = await aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions/current", {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(signedOutCurrent.status).toBe(401);
    adminBearer = await bearerFor(app, admin.email, adminPassword);
    signedOut = false;
    const newAuthSessionCurrent = await aiApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(await newAuthSessionCurrent.json()).toEqual({ session: null });
    const agentPrototype = AgentSession.prototype as unknown as Record<
      string,
      unknown
    >;
    const originalNormalizePromptImages = agentPrototype._normalizePromptImages;
    if (typeof originalNormalizePromptImages !== "function") {
      throw new TypeError("Pi AgentSession lacks prompt image normalization");
    }
    const normalizePromptImages = originalNormalizePromptImages as (
      this: AgentSession,
      images: unknown
    ) => Promise<unknown>;
    const originalAbort = AgentSession.prototype.abort;
    agentPrototype._normalizePromptImages =
      async function _normalizePromptImages(
        this: AgentSession,
        images: unknown
      ): Promise<unknown> {
        if (blockNextPromptPreflight) {
          blockNextPromptPreflight = false;
          notifyPromptPreflightPaused();
          await promptPreflightGate;
        }
        return await normalizePromptImages.call(this, images);
      };
    AgentSession.prototype.abort = async function abort(
      this: AgentSession
    ): Promise<void> {
      notifyPromptAbortRequested();
      await originalAbort.call(this);
    };
    restoreAgentSessionHooks = () => {
      agentPrototype._normalizePromptImages = originalNormalizePromptImages;
      AgentSession.prototype.abort = originalAbort;
    };

    const directoriesBeforePreflight = new Set(await temporaryDirectories());
    blockNextPromptPreflight = true;
    const preflightCreate = aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: true,
          prompt: "Create an equipment request form",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${adminBearer}`,
        },
        method: "POST",
      })
    );
    await promptPreflightPaused;
    const directoriesAfterPreflight = await temporaryDirectories();
    const preflightDirectories = directoriesAfterPreflight.filter(
      (directory) => !directoriesBeforePreflight.has(directory)
    );
    expect(preflightDirectories).toHaveLength(1);
    const [preflightDirectory] = preflightDirectories;
    if (!preflightDirectory) {
      throw new Error(
        "Preflight AI Authoring did not create its temporary directory"
      );
    }
    const signOutPreflight = aiApp.handle(
      new Request("http://test.local/api/auth/sign-out", {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      })
    );
    await promptAbortRequested;
    releasePromptPreflight();
    const [preflightResponse, signOutResponse] = await Promise.all([
      preflightCreate,
      signOutPreflight,
    ]);
    expect(signOutResponse.status).toBe(200);
    expect(preflightResponse.status).toBe(502);
    signedOut = true;
    activeSessionId = undefined;
    expect(await temporaryDirectories()).not.toContain(third.directory);
    expect(await temporaryDirectories()).not.toContain(preflightDirectory);
    expect(upstreamCalls).toBe(6);
    restoreAgentSessionHooks?.();
    restoreAgentSessionHooks = undefined;

    forcedToolCalls = [true, false];
    const laterSession = await createPreview(otherAdminBearer);
    const endLaterSession = await aiApp.handle(
      new Request(
        `http://test.local/api/admin/ai-authoring/sessions/${laterSession.session.sessionId}`,
        {
          headers: { Authorization: `Bearer ${otherAdminBearer}` },
          method: "DELETE",
        }
      )
    );
    expect(endLaterSession.status).toBe(200);
    expect(await temporaryDirectories()).not.toContain(laterSession.directory);
    const directoriesBeforePending = new Set(await temporaryDirectories());
    blockNextProviderRequest = true;
    const pendingCreate = aiApp.handle(
      new Request("http://test.local/api/admin/ai-authoring/sessions", {
        body: JSON.stringify({
          consent: true,
          prompt: "Create an equipment request form",
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${otherAdminBearer}`,
        },
        method: "POST",
      })
    );
    await providerBlocked;
    const directoriesAfterPending = await temporaryDirectories();
    const pendingDirectories = directoriesAfterPending.filter(
      (directory) => !directoriesBeforePending.has(directory)
    );
    expect(pendingDirectories).toHaveLength(1);
    const [pendingDirectory] = pendingDirectories;
    if (!pendingDirectory) {
      throw new Error(
        "Pending AI Authoring did not create its temporary directory"
      );
    }
    const pendingDisposeError = new Error(
      "Pending AI authoring cleanup failed"
    );
    const originalDispose = AgentSession.prototype.dispose;
    let pendingDisposeFailed = false;
    AgentSession.prototype.dispose = function dispose(
      this: AgentSession
    ): void {
      if (!pendingDisposeFailed) {
        pendingDisposeFailed = true;
        throw pendingDisposeError;
      }
      originalDispose.call(this);
    };
    let pendingResponses: [Response, Response];
    try {
      const signOutPending = aiApp.handle(
        new Request("http://test.local/api/auth/sign-out", {
          headers: { Authorization: `Bearer ${otherAdminBearer}` },
          method: "POST",
        })
      );
      await providerAborted;
      releaseProviderResponse();
      pendingResponses = await Promise.all([pendingCreate, signOutPending]);
    } finally {
      AgentSession.prototype.dispose = originalDispose;
    }
    const [pendingResponse, pendingSignOutResponse] = pendingResponses;
    expect(pendingSignOutResponse.status).toBe(500);
    expect(await pendingSignOutResponse.json()).toMatchObject({
      error: "sign_out_cleanup_failed",
      sessionRevoked: true,
    });
    expect(pendingDisposeFailed).toBe(true);
    expect(pendingResponse.status).toBe(502);
    expect(await temporaryDirectories()).not.toContain(pendingDirectory);
    expect(await temporaryDirectories()).not.toContain(laterSession.directory);
    expect(upstreamCalls).toBe(9);

    mutationActorBearer = await bearerFor(
      app,
      delayedBodyAdmin.email,
      delayedBodyPassword
    );
    cleanupTargetBearer = await bearerFor(
      app,
      otherAdmin.email,
      otherAdminPassword
    );
    forcedToolCalls = [true, false];
    const roleSession = await createPreview(cleanupTargetBearer);
    const demoteTarget = await aiApp.handle(
      new Request(`http://test.local/api/admin/users/${otherAdmin.id}`, {
        body: JSON.stringify({ role: "user" }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${mutationActorBearer}`,
        },
        method: "PATCH",
      })
    );
    expect(demoteTarget.status).toBe(200);
    expect(await temporaryDirectories()).not.toContain(roleSession.directory);
    const demotedSessionStatus = await aiApp.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${cleanupTargetBearer}` },
      })
    );
    expect(demotedSessionStatus.status).toBe(401);

    const promoteTarget = await aiApp.handle(
      new Request(`http://test.local/api/admin/users/${otherAdmin.id}`, {
        body: JSON.stringify({ role: "admin" }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${mutationActorBearer}`,
        },
        method: "PATCH",
      })
    );
    expect(promoteTarget.status).toBe(200);
    cleanupTargetBearer = await bearerFor(
      app,
      otherAdmin.email,
      otherAdminPassword
    );
    forcedToolCalls = [true, false];
    const resetSession = await createPreview(cleanupTargetBearer);
    const resetTargetPassword = await aiApp.handle(
      new Request(
        `http://test.local/api/admin/users/${otherAdmin.id}/password-reset`,
        {
          headers: { Authorization: `Bearer ${mutationActorBearer}` },
          method: "POST",
        }
      )
    );
    expect(resetTargetPassword.status).toBe(200);
    expect(await temporaryDirectories()).not.toContain(resetSession.directory);
    const resetBody = (await resetTargetPassword.json()) as {
      temporaryPassword: string;
    };
    const resetSessionStatus = await aiApp.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${cleanupTargetBearer}` },
      })
    );
    expect(resetSessionStatus.status).toBe(401);

    cleanupTargetBearer = await bearerFor(
      app,
      otherAdmin.email,
      resetBody.temporaryPassword
    );
    const changeTargetPassword = await aiApp.handle(
      new Request("http://test.local/api/account/password", {
        body: JSON.stringify({
          currentPassword: resetBody.temporaryPassword,
          newPassword: changedOtherAdminPassword,
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${cleanupTargetBearer}`,
        },
        method: "POST",
      })
    );
    expect(changeTargetPassword.status).toBe(200);
    const changedPasswordSessionStatus = await aiApp.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${cleanupTargetBearer}` },
      })
    );
    expect(changedPasswordSessionStatus.status).toBe(401);
    cleanupTargetBearer = await bearerFor(
      app,
      otherAdmin.email,
      changedOtherAdminPassword
    );
    forcedToolCalls = [true, false];
    const deleteSession = await createPreview(cleanupTargetBearer);
    deletionActorBearer = await bearerFor(app, admin.email, adminPassword);
    const deleteTarget = await aiApp.handle(
      new Request(`http://test.local/api/admin/users/${otherAdmin.id}`, {
        body: JSON.stringify({ confirm: true }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${deletionActorBearer}`,
        },
        method: "DELETE",
      })
    );
    expect(deleteTarget.status).toBe(200);
    expect(await temporaryDirectories()).not.toContain(deleteSession.directory);
    const deletedSessionStatus = await aiApp.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${cleanupTargetBearer}` },
      })
    );
    expect(deletedSessionStatus.status).toBe(401);

    const directoriesBeforeFailedGeneration = new Set(
      await temporaryDirectories()
    );
    const originalGetLastAssistantText =
      AgentSession.prototype.getLastAssistantText;
    const originalFailedGenerationDispose = AgentSession.prototype.dispose;
    const originalUint8ArrayFill = Uint8Array.prototype.fill;
    const uint8ArrayFillSpy = vi.spyOn(Uint8Array.prototype, "fill");
    let generatedDocumentZeroed = false;
    AgentSession.prototype.getLastAssistantText = () => {
      throw new Error("generation failed after DOCX creation");
    };
    AgentSession.prototype.dispose = () => {
      throw new Error("Pi disposal failed");
    };
    uint8ArrayFillSpy.mockImplementation(function fill(
      this: Uint8Array,
      value: number,
      start?: number,
      end?: number
    ): Uint8Array {
      const isGeneratedDocument =
        value === 0 && this[0] === 0x50 && this[1] === 0x4b;
      const result = originalUint8ArrayFill.call(this, value, start, end);
      if (isGeneratedDocument) {
        generatedDocumentZeroed = this.every((byte) => byte === 0);
      }
      return result;
    });
    try {
      forcedToolCalls = [true, false];
      const failedGeneration = await aiApp.handle(
        new Request("http://test.local/api/admin/ai-authoring/sessions", {
          body: JSON.stringify({
            consent: true,
            prompt: "Create an equipment request form",
          }),
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${mutationActorBearer}`,
          },
          method: "POST",
        })
      );
      expect(failedGeneration.status).toBe(502);
      const failedGenerationBody = (await failedGeneration.json()) as {
        error?: string;
      };
      expect(failedGenerationBody.error).toBe("ai_authoring_failed");
      expect(generatedDocumentZeroed).toBe(true);
      const directoriesAfterFailedGeneration = await temporaryDirectories();
      expect(
        directoriesAfterFailedGeneration.filter(
          (directory) => !directoriesBeforeFailedGeneration.has(directory)
        )
      ).toHaveLength(0);
    } finally {
      AgentSession.prototype.getLastAssistantText =
        originalGetLastAssistantText;
      AgentSession.prototype.dispose = originalFailedGenerationDispose;
      uint8ArrayFillSpy.mockRestore();
    }

    forcedToolCalls = [true, false];
    const actorSession = await createPreview(mutationActorBearer);
    const changeActorPassword = await aiApp.handle(
      new Request("http://test.local/api/account/password", {
        body: JSON.stringify({
          currentPassword: delayedBodyPassword,
          newPassword: changedDelayedBodyPassword,
        }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${mutationActorBearer}`,
        },
        method: "POST",
      })
    );
    expect(changeActorPassword.status).toBe(200);
    expect(await temporaryDirectories()).not.toContain(actorSession.directory);
    const changedActorSessionStatus = await aiApp.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${mutationActorBearer}` },
      })
    );
    expect(changedActorSessionStatus.status).toBe(401);
  } finally {
    restoreAgentSessionHooks?.();
    releasePromptPreflight();
    if (formPublicId && adminBearer && !signedOut) {
      await aiApp
        .handle(
          new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
            headers: { Authorization: `Bearer ${adminBearer}` },
            method: "DELETE",
          })
        )
        .catch(() => {});
    }
    if (activeSessionId && adminBearer && !signedOut) {
      await aiApp
        .handle(
          new Request(
            `http://test.local/api/admin/ai-authoring/sessions/${activeSessionId}`,
            {
              headers: { Authorization: `Bearer ${adminBearer}` },
              method: "DELETE",
            }
          )
        )
        .catch(() => {});
    }
    if (adminBearer && !signedOut) {
      await aiApp
        .handle(
          new Request("http://test.local/api/auth/sign-out", {
            headers: { Authorization: `Bearer ${adminBearer}` },
            method: "POST",
          })
        )
        .catch(() => {});
    }
    for (const bearer of [
      otherAdminBearer,
      sameUserOtherBearer,
      userBearer,
      mutationActorBearer,
      deletionActorBearer,
      cleanupTargetBearer,
    ]) {
      if (bearer) {
        await aiApp
          .handle(
            new Request("http://test.local/api/auth/sign-out", {
              headers: { Authorization: `Bearer ${bearer}` },
              method: "POST",
            })
          )
          .catch(() => {});
      }
    }
    upstream.stop(true);
  }
});
