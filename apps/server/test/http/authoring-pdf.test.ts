import { test, expect } from "bun:test";

import { unzipSync } from "fflate";

import { createApp } from "../../src/app";
import {
  createCredentialFixture,
  bearerFor,
  jsonHeaders,
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

test("Ticket 19 PDF authoring re-inspects original bytes when revising a retained session", async () => {
  const model = "ticket-19-pdf-model";
  const serviceKey = `ticket-19-${crypto.randomUUID()}`;
  const originalPdf = Uint8Array.from([
    ...new TextEncoder().encode(
      "%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>\nendobj\n% source bytes "
    ),
    0,
    255,
    ...new TextEncoder().encode("\n%%EOF\n"),
  ]);
  const firstPrompt = "Create an equipment intake form from the attached PDF";
  const revisionPrompt = "Add a serial number field from the same PDF";
  const instructionsRevisionPrompt =
    "Shorten the instructions to match the attached PDF";
  const observations = [
    "PDF observation: the intake requires an employee name.",
    "PDF observation: the intake also requires an equipment serial number.",
    "PDF observation: enter the employee name and equipment serial number.",
  ];
  const receivedPdfs: Uint8Array[] = [];
  const piMessages: string[] = [];
  let piCalls = 0;
  let invalidEditDeclaration = false;
  let blockInspection = false;
  const { promise: inspectionStarted, resolve: notifyInspectionStarted } =
    Promise.withResolvers<undefined>();
  const { promise: inspectionGate, resolve: releaseInspection } =
    Promise.withResolvers<undefined>();
  const upstream = Bun.serve({
    fetch: async (request) => {
      const body = (await request.json()) as {
        messages?: {
          content?:
            | string
            | {
                file?: { file_data?: string; filename?: string };
                text?: string;
                type?: string;
              }[];
        }[];
        model?: string;
        stream?: boolean;
      };
      const messages = body.messages ?? [];
      if (body.stream === false) {
        const content = messages.at(-1)?.content;
        const file = Array.isArray(content)
          ? content.find((part) => part.type === "file")?.file
          : undefined;
        const encoded = file?.file_data?.match(
          /^data:application\/pdf;base64,(?<encoded>[A-Za-z0-9+/=]+)$/u
        )?.[1];
        receivedPdfs.push(
          encoded
            ? Uint8Array.from(Buffer.from(encoded, "base64"))
            : new Uint8Array()
        );
        if (blockInspection) {
          blockInspection = false;
          notifyInspectionStarted();
          await inspectionGate;
        }
        const observation = observations[receivedPdfs.length - 1];
        return Response.json({
          choices: [
            {
              finish_reason: "stop",
              index: 0,
              message: {
                content: observation ?? "PDF observation: intake form.",
                role: "assistant",
              },
            },
          ],
          created: Math.floor(Date.now() / 1000),
          id: `pdf-${receivedPdfs.length}`,
          model,
          object: "chat.completion",
        });
      }
      const text = JSON.stringify(messages);
      piMessages.push(text);
      piCalls += 1;
      const revisingInstructions = text.includes(instructionsRevisionPrompt);
      const revising = text.includes(revisionPrompt) || revisingInstructions;
      const toolCall = piCalls % 2 === 1;
      const generated = {
        description: "Request equipment for work.",
        ...(revisingInstructions
          ? {
              editedParagraphs: [
                invalidEditDeclaration
                  ? "This paragraph does not exist in the current document."
                  : "Complete each field.",
              ],
            }
          : {}),
        fields: [
          {
            label: "Employee Name",
            placeholder: "Enter employee name",
            tag: "employee_name",
          },
          ...(revising
            ? [
                {
                  label: "Serial Number",
                  placeholder: "Enter serial number",
                  tag: "serial_number",
                },
              ]
            : []),
        ],
        paragraphs: [
          revisingInstructions ? "Enter name, serial." : "Complete each field.",
          ...(invalidEditDeclaration
            ? ["Route the intake to Facilities."]
            : []),
        ],
        title: "Equipment Intake",
      };
      const chunks = toolCall
        ? [
            {
              delta: {
                role: "assistant",
                tool_calls: [
                  {
                    function: {
                      arguments: JSON.stringify(generated),
                      name: "create_template_docx",
                    },
                    id: `tool-${piCalls}`,
                    index: 0,
                    type: "function",
                  },
                ],
              },
              finish_reason: null,
            },
            { delta: {}, finish_reason: "tool_calls" },
          ]
        : [
            {
              delta: {
                content: revising
                  ? "Revised the intake form."
                  : "Created the intake form.",
                role: "assistant",
              },
              finish_reason: null,
            },
            { delta: {}, finish_reason: "stop" },
          ];
      const stream = chunks
        .map(
          (chunk) =>
            `data: ${JSON.stringify({
              choices: [{ index: 0, ...chunk }],
              created: Math.floor(Date.now() / 1000),
              id: `chat-${piCalls}`,
              model,
              object: "chat.completion.chunk",
            })}\n\n`
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
  const pdfApp = createApp({
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
  const sessionUrl = "http://test.local/api/admin/ai-authoring/sessions";
  const currentUrl = `${sessionUrl}/current`;
  const password = "Ticket19-pdf-password";
  let adminBearer = "";
  let userBearer = "";
  let otherAdminBearer = "";
  let activeSessionId: string | undefined;
  const createRequest = (
    authorization?: string,
    pdf = originalPdf,
    mime = "application/pdf",
    consent = "true",
    filename = "source.pdf"
  ) => {
    const body = new FormData();
    body.set("prompt", firstPrompt);
    body.set("consent", consent);
    body.set("pdf", new File([pdf], filename, { type: mime }));
    return new Request(sessionUrl, {
      body,
      headers: authorization
        ? { Authorization: `Bearer ${authorization}` }
        : {},
      method: "POST",
    });
  };
  const getCurrent = (authorization: string) =>
    pdfApp.handle(
      new Request(currentUrl, {
        headers: { Authorization: `Bearer ${authorization}` },
      })
    );
  try {
    const admin = await createCredentialFixture({
      email: `ticket-19-admin-${crypto.randomUUID()}@example.com`,
      name: "Ticket 19 Admin",
      password,
      role: "admin",
    });
    const otherAdmin = await createCredentialFixture({
      email: `ticket-19-other-${crypto.randomUUID()}@example.com`,
      name: "Ticket 19 Other Admin",
      password,
      role: "admin",
    });
    const user = await createCredentialFixture({
      email: `ticket-19-user-${crypto.randomUUID()}@example.com`,
      name: "Ticket 19 User",
      password,
    });
    adminBearer = await bearerFor(app, admin.email, password);
    otherAdminBearer = await bearerFor(app, otherAdmin.email, password);
    userBearer = await bearerFor(app, user.email, password);
    const anonymousCreateResponse = await pdfApp.handle(createRequest());
    expect(anonymousCreateResponse.status).toBe(401);
    const userCreateResponse = await pdfApp.handle(createRequest(userBearer));
    expect(userCreateResponse.status).toBe(403);
    const disabledCreateResponse = await disabledApp.handle(
      createRequest(adminBearer)
    );
    expect(disabledCreateResponse.status).toBe(503);
    const missingConsentResponse = await pdfApp.handle(
      createRequest(adminBearer, originalPdf, "application/pdf", "false")
    );
    expect(missingConsentResponse.status).toBe(428);
    const wrongContentTypeResponse = await pdfApp.handle(
      createRequest(
        adminBearer,
        originalPdf,
        "text/plain",
        "true",
        "source.txt"
      )
    );
    expect(wrongContentTypeResponse.status).toBe(415);
    const invalidPdfResponse = await pdfApp.handle(
      createRequest(adminBearer, new TextEncoder().encode("not a PDF"))
    );
    expect(invalidPdfResponse.status).toBe(415);
    const oversizedPdf = new Uint8Array(10 * 1024 * 1024 + 1);
    oversizedPdf.set(originalPdf);
    const oversizedPdfResponse = await pdfApp.handle(
      createRequest(adminBearer, oversizedPdf)
    );
    expect(oversizedPdfResponse.status).toBe(413);
    const oversizedBody = new FormData();
    oversizedBody.set("prompt", firstPrompt);
    oversizedBody.set("consent", "true");
    oversizedBody.set(
      "pdf",
      new File([originalPdf], "source.pdf", { type: "application/pdf" })
    );
    oversizedBody.set("padding", "x".repeat(11 * 1024 * 1024));
    const oversizedBodyResponse = await pdfApp.handle(
      new Request(sessionUrl, {
        body: oversizedBody,
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      })
    );
    expect(oversizedBodyResponse.status).toBe(413);
    expect(receivedPdfs).toHaveLength(0);

    const createdResponse = await pdfApp.handle(createRequest(adminBearer));
    expect(createdResponse.status).toBe(200);
    const { session: created } = (await createdResponse.json()) as {
      session: {
        downloadUrl: string;
        hasSourcePdf: boolean;
        sessionId: string;
      };
    };
    activeSessionId = created.sessionId;
    expect(created.hasSourcePdf).toBe(true);
    expect(receivedPdfs).toEqual([originalPdf]);
    expect(piMessages[0]).toContain(observations[0]);
    const initialDownload = await pdfApp.handle(
      new Request(`http://test.local${created.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(initialDownload.status).toBe(200);
    const initialXml = unzipSync(
      new Uint8Array(await initialDownload.arrayBuffer())
    )["word/document.xml"];
    expect(new TextDecoder().decode(initialXml)).toContain(
      '<w:tag w:val="employee_name"/>'
    );

    const refreshed = await getCurrent(adminBearer);
    expect(refreshed.status).toBe(200);
    expect(await refreshed.json()).toEqual({ session: created });
    const otherAdminCurrentResponse = await getCurrent(otherAdminBearer);
    expect(await otherAdminCurrentResponse.json()).toEqual({
      session: null,
    });
    const revisionUrl = `${sessionUrl}/${created.sessionId}/revisions`;
    const revisionRequest = (
      authorization: string,
      consent = true,
      prompt = revisionPrompt
    ) =>
      new Request(revisionUrl, {
        body: JSON.stringify({ consent, prompt }),
        headers: { ...jsonHeaders, Authorization: `Bearer ${authorization}` },
        method: "POST",
      });
    const otherAdminRevisionResponse = await pdfApp.handle(
      revisionRequest(otherAdminBearer)
    );
    expect(otherAdminRevisionResponse.status).toBe(404);
    const missingRevisionConsentResponse = await pdfApp.handle(
      revisionRequest(adminBearer, false)
    );
    expect(missingRevisionConsentResponse.status).toBe(428);
    expect(receivedPdfs).toHaveLength(1);
    const revisedResponse = await pdfApp.handle(revisionRequest(adminBearer));
    expect(revisedResponse.status).toBe(200);
    const { session: revised } = (await revisedResponse.json()) as {
      session: typeof created;
    };
    expect(revised.hasSourcePdf).toBe(true);
    expect(receivedPdfs).toEqual([originalPdf, originalPdf]);
    expect(piMessages[2]).toContain(observations[1]);
    const revisedDownload = await pdfApp.handle(
      new Request(`http://test.local${revised.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(revisedDownload.status).toBe(200);
    const revisedXml = unzipSync(
      new Uint8Array(await revisedDownload.arrayBuffer())
    )["word/document.xml"];
    expect(new TextDecoder().decode(revisedXml)).toContain(
      '<w:tag w:val="serial_number"/>'
    );
    const currentAfterRevision = await getCurrent(adminBearer);
    expect(await currentAfterRevision.json()).toEqual({
      session: revised,
    });

    const instructionsResponse = await pdfApp.handle(
      revisionRequest(adminBearer, true, instructionsRevisionPrompt)
    );
    expect(instructionsResponse.status).toBe(200);
    const { session: instructionsRevised } =
      (await instructionsResponse.json()) as {
        session: typeof created;
      };
    expect(instructionsRevised.hasSourcePdf).toBe(true);
    expect(receivedPdfs).toEqual([originalPdf, originalPdf, originalPdf]);
    expect(piMessages[4]).toContain(observations[2]);
    const instructionsDownload = await pdfApp.handle(
      new Request(`http://test.local${instructionsRevised.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(instructionsDownload.status).toBe(200);
    const instructionsDocument = new Uint8Array(
      await instructionsDownload.arrayBuffer()
    );
    const instructionsXml =
      unzipSync(instructionsDocument)["word/document.xml"];
    const instructionsText = new TextDecoder().decode(instructionsXml);
    expect(instructionsText).toContain("Enter name, serial.");
    expect(instructionsText).not.toContain("Complete each field.");
    expect(instructionsText).toContain('<w:tag w:val="employee_name"/>');
    expect(instructionsText).toContain('<w:tag w:val="serial_number"/>');
    const currentAfterInstructions = await getCurrent(adminBearer);
    expect(await currentAfterInstructions.json()).toEqual({
      session: instructionsRevised,
    });

    invalidEditDeclaration = true;
    const invalidRevisionResponse = await pdfApp.handle(
      revisionRequest(
        adminBearer,
        true,
        "Add a PDF routing note after the instructions"
      )
    );
    invalidEditDeclaration = false;
    expect(invalidRevisionResponse.status).toBe(502);
    expect(receivedPdfs).toEqual([
      originalPdf,
      originalPdf,
      originalPdf,
      originalPdf,
    ]);
    const currentAfterInvalidRevision = await getCurrent(adminBearer);
    expect(await currentAfterInvalidRevision.json()).toEqual({
      session: instructionsRevised,
    });
    const unchangedDownload = await pdfApp.handle(
      new Request(`http://test.local${instructionsRevised.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(unchangedDownload.status).toBe(200);
    expect(new Uint8Array(await unchangedDownload.arrayBuffer())).toEqual(
      instructionsDocument
    );
    const deleteSessionResponse = await pdfApp.handle(
      new Request(`${sessionUrl}/${created.sessionId}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "DELETE",
      })
    );
    expect(deleteSessionResponse.status).toBe(200);
    activeSessionId = undefined;
    const currentAfterDelete = await getCurrent(adminBearer);
    expect(await currentAfterDelete.json()).toEqual({
      session: null,
    });
    const deletedDownloadResponse = await pdfApp.handle(
      new Request(`http://test.local${created.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(deletedDownloadResponse.status).toBe(404);
    const logoutCreate = await pdfApp.handle(
      createRequest(adminBearer, originalPdf, "", "true", "source.bin")
    );
    expect(logoutCreate.status).toBe(200);
    const { session: logoutSession } = (await logoutCreate.json()) as {
      session: typeof created;
    };
    activeSessionId = logoutSession.sessionId;
    expect(logoutSession.hasSourcePdf).toBe(true);
    const logoutResponse = await pdfApp.handle(
      new Request("http://test.local/api/auth/sign-out", {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      })
    );
    expect(logoutResponse.status).toBe(200);
    activeSessionId = undefined;
    adminBearer = await bearerFor(app, admin.email, password);
    const currentAfterLogout = await getCurrent(adminBearer);
    expect(await currentAfterLogout.json()).toEqual({
      session: null,
    });
    const logoutDownloadResponse = await pdfApp.handle(
      new Request(`http://test.local${logoutSession.downloadUrl}`, {
        headers: { Authorization: `Bearer ${adminBearer}` },
      })
    );
    expect(logoutDownloadResponse.status).toBe(404);

    const expiryCreate = await pdfApp.handle(createRequest(adminBearer));
    expect(expiryCreate.status).toBe(200);
    const { session: expirySession } = (await expiryCreate.json()) as {
      session: typeof created;
    };
    activeSessionId = expirySession.sessionId;
    now = new Date(now.getTime() + 2 * 60 * 60 * 1000);
    const currentAfterExpiry = await getCurrent(adminBearer);
    expect(await currentAfterExpiry.json()).toEqual({
      session: null,
    });
    activeSessionId = undefined;
    const expiredRevisionResponse = await pdfApp.handle(
      revisionRequest(adminBearer)
    );
    expect(expiredRevisionResponse.status).toBe(404);
    expect(receivedPdfs).toEqual([
      originalPdf,
      originalPdf,
      originalPdf,
      originalPdf,
      originalPdf,
      originalPdf,
    ]);

    blockInspection = true;
    const blockedCreate = pdfApp.handle(createRequest(adminBearer));
    await inspectionStarted;
    const logoutWhileBlocked = pdfApp.handle(
      new Request("http://test.local/api/auth/sign-out", {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      })
    );
    // Network abort uses real I/O; fake timers cannot drive this gateway cancellation.
    const completedBeforeGatewayRelease = await Promise.race([
      logoutWhileBlocked.then(() => true),
      Bun.sleep(1500).then(() => false),
    ]);
    releaseInspection();
    expect(completedBeforeGatewayRelease).toBe(true);
    const logoutWhileBlockedResponse = await logoutWhileBlocked;
    expect(logoutWhileBlockedResponse.status).toBe(200);
    const blockedCreateResponse = await blockedCreate;
    expect(blockedCreateResponse.status).not.toBe(200);
    adminBearer = await bearerFor(app, admin.email, password);
    const currentAfterBlockedLogout = await getCurrent(adminBearer);
    expect(await currentAfterBlockedLogout.json()).toEqual({
      session: null,
    });
  } finally {
    releaseInspection();
    if (activeSessionId && adminBearer) {
      await pdfApp
        .handle(
          new Request(`${sessionUrl}/${activeSessionId}`, {
            headers: { Authorization: `Bearer ${adminBearer}` },
            method: "DELETE",
          })
        )
        .catch(() => {});
    }
    for (const bearer of [adminBearer, otherAdminBearer, userBearer]) {
      if (bearer) {
        await pdfApp
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
