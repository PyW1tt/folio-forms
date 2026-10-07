import { test, expect } from "bun:test";

import { unzipSync } from "fflate";

import { createApp } from "../../src/app";
import { createOnlyOfficeAuthorization } from "../../src/onlyoffice";
import { docxFixture } from "../fixtures/documents";
import {
  createCredentialFixture,
  bearerFor,
  formCreationRequest,
  jsonHeaders,
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

const documentWorkerTestUrl = process.env.DOCUMENT_WORKER_URL;

if (documentWorkerTestUrl) {
  test("Ticket 20 Python document worker generates a Template Draft and contains hostile code", async () => {
    const workerUrl = documentWorkerTestUrl;
    const health = await fetch(new URL("/health", workerUrl), {
      signal: AbortSignal.timeout(2000),
    }).catch(() => null);
    if (!health?.ok) {
      throw new Error(
        `Document worker unavailable at ${workerUrl}; start the Docker worker before running this test`
      );
    }
    const model = "ticket-20-python-model";
    const serviceKey = `ticket-20-${crypto.randomUUID()}`;
    const marker = `Python-generated-${crypto.randomUUID()}`;
    const template = {
      description: "Request equipment with a generated document.",
      fields: [
        {
          label: "Employee Name",
          placeholder: "Enter employee name",
          tag: "employee_name",
        },
      ],
      paragraphs: ["Complete the equipment request.", "Use R&D <equipment>."],
      title: `Python Equipment Request ${marker}`,
    };
    const benignSource = `
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

title = ${JSON.stringify(template.title)}
description = ${JSON.stringify(template.description)}
document = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body><w:p><w:r><w:t>' + escape(title) + '</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>' + escape(description) + '</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Complete the equipment request.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>' + escape('Use R&') + '</w:t></w:r>'
    '<w:r><w:t>' + escape('D <equipment>.') + '</w:t></w:r></w:p>'
    '<w:p><w:sdt><w:sdtPr><w:alias w:val="Employee Name"/>'
    '<w:tag w:val="employee_name"/><w:text/><w:showingPlcHdr/></w:sdtPr>'
    '<w:sdtContent><w:r><w:t>Enter employee name</w:t></w:r></w:sdtContent>'
    '</w:sdt></w:p><w:sectPr/></w:body></w:document>'
)
with ZipFile('output.docx', 'w', ZIP_DEFLATED) as archive:
    archive.writestr('[Content_Types].xml', '''<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
    archive.writestr('_rels/.rels', '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
    archive.writestr('word/document.xml', document)
`;
    const revisionPrompt =
      "A warmer welcome and a requester identity suit this form better.";
    const paragraphRevision = {
      ...template,
      paragraphs: [
        "Welcome! Tell us which equipment you need.",
        "Use R&D <equipment>.",
      ],
      source: benignSource.replace(
        "Complete the equipment request.",
        "Welcome! Tell us which equipment you need."
      ),
    };
    const fieldRevision = {
      ...template,
      fields: [
        {
          label: "Requester Name",
          placeholder: "Enter requester name",
          tag: "requester_name",
        },
      ],
      source: benignSource
        .replace("Employee Name", "Requester Name")
        .replace("Enter employee name", "Enter requester name")
        .replace("employee_name", "requester_name"),
    };
    const revisedTemplate = {
      ...paragraphRevision,
      editedParagraphs: ["Complete the equipment request."],
      fields: fieldRevision.fields,
      removedFieldTags: ["employee_name"],
      source: fieldRevision.source
        .replace(
          "Complete the equipment request.",
          "Welcome! Tell us which equipment you need."
        )
        .replace("<w:showingPlcHdr/>", ""),
    };
    const rejectedOutputs = {
      "bogus document output": {
        ...template,
        source: "open('output.docx', 'wb').write(b'not a DOCX')",
      },
      "field revision with stale declaration": {
        ...fieldRevision,
        removedFieldTags: ["unknown_field"],
      },
      "field revision without declaration": fieldRevision,
      "mismatched control output": {
        ...template,
        source: benignSource.replace("employee_name", "undeclared_name"),
      },
      "paragraph revision with stale declaration": {
        ...paragraphRevision,
        editedParagraphs: ["This paragraph is not in the current document."],
      },
      "paragraph revision without declaration": paragraphRevision,
      "source-only description change": {
        ...template,
        source: benignSource.replace(
          "escape(description)",
          "escape('Incorrect description')"
        ),
      },
      "source-only field label change": {
        ...template,
        source: benignSource.replace(
          'w:alias w:val="Employee Name"',
          'w:alias w:val="Incorrect label"'
        ),
      },
      "source-only field placeholder change": {
        ...template,
        source: benignSource.replace(
          "Enter employee name",
          "Incorrect placeholder"
        ),
      },
      "source-only paragraph change": {
        ...template,
        source: benignSource.replace(
          "Complete the equipment request.",
          "Original instructions were silently replaced."
        ),
      },
      "source-only paragraph removal": {
        ...template,
        source: benignSource.replace(
          "'<w:p><w:r><w:t>Complete the equipment request.</w:t></w:r></w:p>'",
          "''"
        ),
      },
      "source-only title change": {
        ...template,
        source: benignSource.replace(
          "escape(title)",
          "escape('Incorrect title')"
        ),
      },
      "unsupported control output": {
        ...template,
        source: benignSource.replace("<w:text/>", "<w:comboBox/>"),
      },
    };
    const hostileSources: Record<string, string> = {
      "contact internet":
        "import socket; socket.create_connection(('1.1.1.1', 443), timeout=1).close()",
      "execute shell":
        "import subprocess; subprocess.run(['/bin/sh', '-c', 'echo breached'], check=True)",
      "exhaust CPU": "while True: pass",
      "exhaust memory": "allocation = bytearray(512 * 1024 * 1024)",
      "read private server file": "open('/app/.env').read()",
      "write outside workspace":
        "open('/tmp/folio-worker-escape', 'w').write('breached')",
    };
    const upstream = Bun.serve({
      fetch: async (request) => {
        const body = (await request.json()) as {
          messages?: { content?: unknown; role?: string }[];
          tools?: { function?: { name?: string } }[];
        };
        const messages = body.messages ?? [];
        const latestUserContent = messages
          .toReversed()
          .find((message) => message.role === "user")?.content;
        const promptText = JSON.stringify(latestUserContent ?? "");
        const attemptedAttack = Object.entries(hostileSources).find(
          ([prompt]) => promptText.includes(prompt)
        );
        const requestedOutput = Object.entries({
          ...rejectedOutputs,
          [revisionPrompt]: revisedTemplate,
        }).find(([prompt]) => promptText.includes(prompt))?.[1];
        const generated = attemptedAttack
          ? { ...template, source: `${attemptedAttack[1]}\n${benignSource}` }
          : (requestedOutput ?? { ...template, source: benignSource });
        // Pi also sends tool-free requests when compacting the long attack history.
        const shouldSummarize = messages.at(-1)?.role === "tool" || !body.tools;
        const chunks = shouldSummarize
          ? [
              {
                delta: { content: "Document prepared.", role: "assistant" },
                finish_reason: null,
              },
              { delta: {}, finish_reason: "stop" },
            ]
          : [
              {
                delta: {
                  role: "assistant",
                  tool_calls: [
                    {
                      function: {
                        arguments: JSON.stringify(generated),
                        name: "create_template_docx_python",
                      },
                      id: `call-${crypto.randomUUID()}`,
                      index: 0,
                      type: "function",
                    },
                  ],
                },
                finish_reason: null,
              },
              { delta: {}, finish_reason: "tool_calls" },
            ];
        const stream = chunks
          .map(
            (chunk) =>
              `data: ${JSON.stringify({
                choices: [{ index: 0, ...chunk }],
                created: Math.floor(Date.now() / 1000),
                id: `chatcmpl-${crypto.randomUUID()}`,
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
    const workerApp = createApp({
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
    const sessionUrl = "http://test.local/api/admin/ai-authoring/sessions";
    const currentUrl = `${sessionUrl}/current`;
    const password = "Ticket20-worker-password";
    let adminBearer = "";
    let otherAdminBearer = "";
    let userBearer = "";
    let sessionId = "";
    let otherSessionId = "";
    let baselinePublicId = "";
    let generatedPublicId = "";
    try {
      const admin = await createCredentialFixture({
        email: `ticket-20-admin-${crypto.randomUUID()}@example.com`,
        name: "Ticket 20 Admin",
        password,
        role: "admin",
      });
      const otherAdmin = await createCredentialFixture({
        email: `ticket-20-other-${crypto.randomUUID()}@example.com`,
        name: "Ticket 20 Other Admin",
        password,
        role: "admin",
      });
      const user = await createCredentialFixture({
        email: `ticket-20-user-${crypto.randomUUID()}@example.com`,
        name: "Ticket 20 User",
        password,
      });
      adminBearer = await bearerFor(app, admin.email, password);
      otherAdminBearer = await bearerFor(app, otherAdmin.email, password);
      userBearer = await bearerFor(app, user.email, password);
      const templateDocument = async (
        bearer: string,
        publicId: string
      ): Promise<Uint8Array> => {
        const editorResponse = await workerApp.handle(
          new Request(
            `http://test.local/api/admin/forms/${publicId}/editor-config`,
            { headers: { Authorization: `Bearer ${bearer}` } }
          )
        );
        expect(editorResponse.status).toBe(200);
        const editor = (await editorResponse.json()) as EditorConfigBody;
        const { url } = editor.config.document;
        const response = await workerApp.handle(
          new Request(url, {
            headers: { Authorization: createOnlyOfficeAuthorization({ url }) },
          })
        );
        expect(response.status).toBe(200);
        return new Uint8Array(await response.arrayBuffer());
      };
      const existingDocument = docxFixture(`untouched-${crypto.randomUUID()}`);
      const baselineResponse = await workerApp.handle(
        formCreationRequest({
          authorization: otherAdminBearer,
          source: "upload",
          template: { bytes: existingDocument, name: "existing.docx" },
          title: "Existing Form",
        })
      );
      expect(baselineResponse.status).toBe(200);
      const baseline = (await baselineResponse.json()) as {
        form: { publicId: string };
      };
      baselinePublicId = baseline.form.publicId;

      const create = (bearer: string) =>
        workerApp.handle(
          new Request(sessionUrl, {
            body: JSON.stringify({
              consent: true,
              prompt: "Create Python equipment request",
            }),
            headers: { ...jsonHeaders, Authorization: `Bearer ${bearer}` },
            method: "POST",
          })
        );
      const createdResponse = await create(adminBearer);
      expect(createdResponse.status).toBe(200);
      const { session } = (await createdResponse.json()) as {
        session: {
          description: string;
          downloadUrl: string;
          fields: typeof template.fields;
          sessionId: string;
          title: string;
        };
      };
      ({ sessionId } = session);
      expect(session.title).toBe(template.title);
      expect(session.description).toBe(template.description);
      expect(session.fields).toEqual(template.fields);
      const download = (bearer: string, url: string) =>
        workerApp.handle(
          new Request(`http://test.local${url}`, {
            headers: { Authorization: `Bearer ${bearer}` },
          })
        );
      const documentResponse = await download(adminBearer, session.downloadUrl);
      expect(documentResponse.status).toBe(200);
      const generatedDocument = new Uint8Array(
        await documentResponse.arrayBuffer()
      );
      const documentPart = unzipSync(generatedDocument)["word/document.xml"];
      if (!documentPart) {
        throw new Error("Python output has no document part");
      }
      const xml = new TextDecoder().decode(documentPart);
      expect(xml).toContain(marker);
      expect(xml).toContain("Complete the equipment request.");
      expect(xml).toContain('<w:tag w:val="employee_name"/>');

      const otherResponse = await create(otherAdminBearer);
      expect(otherResponse.status).toBe(200);
      const { session: otherSession } = (await otherResponse.json()) as {
        session: { downloadUrl: string; sessionId: string };
      };
      otherSessionId = otherSession.sessionId;
      const otherDownload = await download(
        otherAdminBearer,
        otherSession.downloadUrl
      );
      expect(otherDownload.status).toBe(200);
      const otherDocument = new Uint8Array(await otherDownload.arrayBuffer());

      for (const prompt of [
        ...Object.keys(rejectedOutputs),
        ...Object.keys(hostileSources),
      ]) {
        const rejected = await workerApp.handle(
          new Request(`${sessionUrl}/${sessionId}/revisions`, {
            body: JSON.stringify({ consent: true, prompt }),
            headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
            method: "POST",
          })
        );
        expect(rejected.status).toBe(502);
        const currentAfterRejectedRevision = await workerApp.handle(
          new Request(currentUrl, {
            headers: { Authorization: `Bearer ${adminBearer}` },
          })
        );
        expect(await currentAfterRejectedRevision.json()).toEqual({ session });
        const retained = await download(adminBearer, session.downloadUrl);
        expect(retained.status).toBe(200);
        expect(new Uint8Array(await retained.arrayBuffer())).toEqual(
          generatedDocument
        );
        const otherRetained = await download(
          otherAdminBearer,
          otherSession.downloadUrl
        );
        expect(otherRetained.status).toBe(200);
        expect(new Uint8Array(await otherRetained.arrayBuffer())).toEqual(
          otherDocument
        );
        const userSessionResponse = await workerApp.handle(
          new Request("http://test.local/api/session", {
            headers: { Authorization: `Bearer ${userBearer}` },
          })
        );
        expect(userSessionResponse.status).toBe(200);
        const baselineStillExists = await workerApp.handle(
          new Request(`http://test.local/api/admin/forms/${baselinePublicId}`, {
            headers: { Authorization: `Bearer ${otherAdminBearer}` },
          })
        );
        expect(baselineStillExists.status).toBe(200);
        expect(
          (await baselineStillExists.json()) as { form: { title: string } }
        ).toMatchObject({
          form: { title: "Existing Form" },
        });
        expect(
          await templateDocument(otherAdminBearer, baselinePublicId)
        ).toEqual(existingDocument);
      }

      const revisedResponse = await workerApp.handle(
        new Request(`${sessionUrl}/${sessionId}/revisions`, {
          body: JSON.stringify({ consent: true, prompt: revisionPrompt }),
          headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
          method: "POST",
        })
      );
      expect(revisedResponse.status).toBe(200);
      const { session: revisedSession } = (await revisedResponse.json()) as {
        session: typeof session;
      };
      expect(revisedSession.fields).toEqual(fieldRevision.fields);
      const currentAfterRevision = await workerApp.handle(
        new Request(currentUrl, {
          headers: { Authorization: `Bearer ${adminBearer}` },
        })
      );
      expect(await currentAfterRevision.json()).toEqual({
        session: revisedSession,
      });
      const revisedDownload = await download(
        adminBearer,
        revisedSession.downloadUrl
      );
      expect(revisedDownload.status).toBe(200);
      const revisedDocument = new Uint8Array(
        await revisedDownload.arrayBuffer()
      );
      const revisedXml = new TextDecoder().decode(
        unzipSync(revisedDocument)["word/document.xml"]
      );
      expect(revisedXml).toContain(marker);
      expect(revisedXml).toContain(
        "Welcome! Tell us which equipment you need."
      );
      expect(revisedXml).not.toContain("Complete the equipment request.");
      expect(revisedXml).toContain('<w:tag w:val="requester_name"/>');
      expect(revisedXml).not.toContain('<w:tag w:val="employee_name"/>');

      const uploadResponse = await workerApp.handle(
        formCreationRequest({
          authorization: adminBearer,
          description: revisedSession.description,
          source: "upload",
          template: { bytes: revisedDocument, name: "python-equipment.docx" },
          title: revisedSession.title,
        })
      );
      expect(uploadResponse.status).toBe(200);
      const { form } = (await uploadResponse.json()) as {
        form: { hasTemplateDraft: boolean; publicId: string };
      };
      generatedPublicId = form.publicId;
      expect(form.hasTemplateDraft).toBe(true);
      expect(await templateDocument(adminBearer, generatedPublicId)).toEqual(
        revisedDocument
      );
      expect(
        await templateDocument(otherAdminBearer, baselinePublicId)
      ).toEqual(existingDocument);
    } finally {
      for (const [bearer, id] of [
        [adminBearer, sessionId],
        [otherAdminBearer, otherSessionId],
      ]) {
        if (bearer && id) {
          await workerApp
            .handle(
              new Request(`${sessionUrl}/${id}`, {
                headers: { Authorization: `Bearer ${bearer}` },
                method: "DELETE",
              })
            )
            .catch(() => {});
        }
      }
      for (const [bearer, publicId] of [
        [adminBearer, generatedPublicId],
        [otherAdminBearer, baselinePublicId],
      ]) {
        if (bearer && publicId) {
          await workerApp
            .handle(
              new Request(`http://test.local/api/admin/forms/${publicId}`, {
                headers: { Authorization: `Bearer ${bearer}` },
                method: "DELETE",
              })
            )
            .catch(() => {});
        }
      }
      for (const bearer of [adminBearer, otherAdminBearer, userBearer]) {
        if (bearer) {
          await workerApp
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
}
