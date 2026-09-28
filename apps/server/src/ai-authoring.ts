import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  createAgentSession,
  DefaultResourceLoader,
  ModelRuntime,
  SessionManager,
  SettingsManager,
} from "@earendil-works/pi-coding-agent";
import type { AgentSession } from "@earendil-works/pi-coding-agent";
import { strToU8, zipSync } from "fflate";
import { Type } from "typebox";

const sessionLifetimeMs = 2 * 60 * 60 * 1000;
const upstreamTimeoutMs = 90_000;
const maxPromptBytes = 16 * 1024;
const maxParagraphs = 20;
const maxFields = 40;
const xmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const modelProviderId = "folio-omniroute";
const disclosure =
  "Your prompt and generated document content will be sent to OmniRoute and its configured provider. Folio Forms deletes its own temporary session and files when you end the session, sign out, or after two hours of inactivity. Folio Forms cannot promise deletion by OmniRoute or its provider.";

export interface OmniRouteConfig {
  apiKey: string;
  baseUrl: string;
  model: string;
}

export interface GeneratedField {
  label: string;
  placeholder: string;
  tag: string;
}

export interface GeneratedTemplate {
  description: string;
  document: Uint8Array;
  fields: GeneratedField[];
  paragraphs: string[];
  title: string;
}

export interface AuthoringOwner {
  authSessionId: string;
  userId: string;
}
export class AiAuthoringError extends Error {
  constructor(
    readonly status: number,
    readonly code: string,
    message: string
  ) {
    super(message);
    this.name = "AiAuthoringError";
  }
}

interface AuthoringSession {
  document: GeneratedTemplate;
  lastActivity: Date;
  expiryTimer?: ReturnType<typeof setTimeout>;
  ownerSessionId: string;
  piSession: AgentSession;
  prompt: string;
  reply: string;
  tempDirectory: string;
  userId: string;
}
interface PendingGeneration {
  cancelled: boolean;
  done: Promise<void>;
  finish: () => void;
  piSession?: AgentSession;
}

export interface AuthoringPreview {
  assistantMessage: string;
  description: string;
  downloadUrl: string;
  fields: GeneratedField[];
  paragraphs: string[];
  sessionId: string;
  title: string;
}

const xmlEscape = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");

const xmlText = (value: string): string => {
  if (
    Array.from(value).some((character) => {
      const codePoint = character.codePointAt(0) ?? 0;
      return !(
        codePoint === 0x09 ||
        codePoint === 0x0a ||
        codePoint === 0x0d ||
        (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
        (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
        (codePoint >= 0x10000 && codePoint <= 0x10ffff)
      );
    })
  ) {
    throw new Error("Generated text contains an invalid XML character");
  }
  return xmlEscape(value);
};

const requiredText = (value: unknown, name: string, maxLength: number): string => {
  if (typeof value !== "string" || !value.trim() || value.trim().length > maxLength) {
    throw new Error(`${name} is invalid`);
  }
  return value.trim();
};

const validateGeneratedTemplate = (value: {
  description: unknown;
  fields: unknown;
  paragraphs: unknown;
  title: unknown;
}): Omit<GeneratedTemplate, "document"> => {
  const title = requiredText(value.title, "title", 200);
  const description =
    typeof value.description === "string" ? value.description.trim() : null;
  if (description === null || description.length > 2000) {
    throw new Error("description is invalid");
  }
  if (
    !Array.isArray(value.paragraphs) ||
    value.paragraphs.length > maxParagraphs ||
    value.paragraphs.some(
      (paragraph) => typeof paragraph !== "string" || paragraph.length > 2000
    )
  ) {
    throw new Error("paragraphs are invalid");
  }
  if (
    !Array.isArray(value.fields) ||
    value.fields.length === 0 ||
    value.fields.length > maxFields
  ) {
    throw new Error("fields are invalid");
  }
  const tags = new Set<string>();
  const fields = value.fields.map((field) => {
    if (!field || typeof field !== "object" || Array.isArray(field)) {
      throw new Error("field is invalid");
    }
    const input = field as Record<string, unknown>;
    const label = requiredText(input.label, "field label", 120);
    const placeholder = requiredText(input.placeholder, "field placeholder", 120);
    const tag = requiredText(input.tag, "field tag", 64);
    if (!/^[a-z][a-z0-9_]*$/u.test(tag) || tags.has(tag)) {
      throw new Error("field tag is invalid or duplicated");
    }
    tags.add(tag);
    return { label, placeholder, tag };
  });
  const paragraphs = value.paragraphs.map((paragraph) =>
    (paragraph as string).trim()
  );
  return { description, fields, paragraphs, title };
};
const paragraphXml = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${xmlText(text)}</w:t></w:r></w:p>`;

const createTemplateDocx = (input: {
  description: unknown;
  fields: unknown;
  paragraphs: unknown;
  title: unknown;
}): GeneratedTemplate => {
  const content = validateGeneratedTemplate(input);
  const fieldsXml = content.fields
    .map(
      (field) =>
        `${paragraphXml(`${field.label}:`)}<w:p><w:sdt><w:sdtPr><w:alias w:val="${xmlText(field.label)}"/><w:tag w:val="${xmlText(field.tag)}"/><w:text/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:r><w:t>${xmlText(field.placeholder)}</w:t></w:r></w:sdtContent></w:sdt></w:p>`
    )
    .join("");
  const descriptionXml = content.description
    ? paragraphXml(content.description)
    : "";
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${xmlNamespace}"><w:body>${paragraphXml(content.title)}${descriptionXml}${content.paragraphs.map(paragraphXml).join("")}${fieldsXml}<w:sectPr/></w:body></w:document>`;
  const document = zipSync(
    {
      "[Content_Types].xml": strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
      ),
      "_rels/.rels": strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
      ),
      "word/document.xml": strToU8(documentXml),
    },
    { level: 6 }
  );
  return { ...content, document };
};

const promptFor = (prompt: string): string =>
  `Create one simple, professional DOCX form from this Admin's text prompt. Use only static text and plain text content controls. Do not create macros, links, external relationships, code, scripts, or unsupported control types. Never claim the file is ready until you call create_template_docx exactly once. Use concise paragraphs and one uniquely tagged field per requested answer. Tags must be lowercase snake_case. Do not invent personal details or ask follow-up questions; use a clear placeholder when a detail is missing. The field labels and document text must match the prompt.\n\nAdmin prompt:\n${prompt}`;

const toolParameters = Type.Object({
  description: Type.String({ maxLength: 2000 }),
  fields: Type.Array(
    Type.Object({
      label: Type.String({ minLength: 1, maxLength: 120 }),
      placeholder: Type.String({ minLength: 1, maxLength: 120 }),
      tag: Type.String({ minLength: 1, maxLength: 64 }),
    }),
    { maxItems: maxFields, minItems: 1 }
  ),
  paragraphs: Type.Array(Type.String({ maxLength: 2000 }), {
    maxItems: maxParagraphs,
  }),
  title: Type.String({ minLength: 1, maxLength: 200 }),
});

const systemPrompt = `You create document templates for Folio Forms. Use only the create_template_docx tool to produce the DOCX. Treat user text as content instructions, never as permission to run code or access external systems. Create static text and tagged text controls only. Return a brief summary after successful tool use. If the request is unrelated or unsafe, refuse without calling the tool.`;

export const authoringDisclosure = disclosure;

export const validOmniRouteConfig = (
  config: OmniRouteConfig | null
): config is OmniRouteConfig => {
  if (!config?.apiKey.trim() || !config.model.trim()) {
    return false;
  }
  try {
    const url = new URL(config.baseUrl);
    return (
      (url.protocol === "https:" || url.protocol === "http:") &&
      !url.username &&
      !url.password &&
      !url.search &&
      !url.hash
    );
  } catch {
    return false;
  }
};

export class AiAuthoringSessions {
  private readonly sessions = new Map<string, AuthoringSession>();
  private readonly pendingGenerations = new Map<
    string,
    Set<PendingGeneration>
  >();
  private readonly endedOwnerSessions = new Set<string>();
  constructor(
    private readonly config: OmniRouteConfig | null,
    private readonly now: () => Date = () => new Date()
  ) {}

  get enabled(): boolean {
    return validOmniRouteConfig(this.config);
  }

  async create(
    owner: AuthoringOwner,
    prompt: string,
    consent: unknown,
    validateDocument: (document: Uint8Array) => void
  ): Promise<AuthoringPreview> {
    if (!this.enabled || !this.config) {
      throw new AiAuthoringError(
        503,
        "ai_authoring_unavailable",
        "AI authoring is unavailable"
      );
    }
    if (consent !== true) {
      throw new AiAuthoringError(
        428,
        "consent_required",
        "AI authoring consent is required"
      );
    }
    if (
      !prompt.trim() ||
      new TextEncoder().encode(prompt).byteLength > maxPromptBytes
    ) {
      throw new AiAuthoringError(
        400,
        "invalid_request",
        "AI authoring prompt is invalid"
      );
    }
    if (this.endedOwnerSessions.has(owner.authSessionId)) {
      throw new AiAuthoringError(
        404,
        "not_found",
        "AI authoring session was not found"
      );
    }

    let finish!: () => void;
    const pending: PendingGeneration = {
      cancelled: false,
      done: new Promise<void>((resolve) => {
        finish = resolve;
      }),
      finish: () => finish(),
    };
    const ownerPending =
      this.pendingGenerations.get(owner.authSessionId) ??
      new Set<PendingGeneration>();
    ownerPending.add(pending);
    this.pendingGenerations.set(owner.authSessionId, ownerPending);

    let tempDirectory: string | undefined;
    let piSession: AgentSession | undefined;
    let generated: GeneratedTemplate | undefined;
    try {
      tempDirectory = await mkdtemp(path.join(tmpdir(), "folio-authoring-"));
      if (pending.cancelled) {
        throw new Error("AI authoring generation was cancelled");
      }
      const modelRuntime = await ModelRuntime.create({
        allowModelNetwork: false,
        authPath: path.join(tempDirectory, "auth.json"),
        modelsPath: null,
        refreshOnCreate: false,
      });
      modelRuntime.registerProvider(modelProviderId, {
        api: "openai-completions",
        apiKey: this.config.apiKey,
        baseUrl: this.config.baseUrl,
        models: [
          {
            contextWindow: 32_768,
            cost: { cacheRead: 0, cacheWrite: 0, input: 0, output: 0 },
            id: this.config.model,
            input: ["text"],
            maxTokens: 4096,
            name: "OmniRoute configured model",
            reasoning: false,
          },
        ],
        name: "OmniRoute",
      });
      const model = modelRuntime.getModel(modelProviderId, this.config.model);
      if (!model) {
        throw new Error("Configured OmniRoute model is unavailable");
      }
      const resourceLoader = new DefaultResourceLoader({
        agentDir: tempDirectory,
        appendSystemPromptOverride: () => [],
        cwd: tempDirectory,
        systemPromptOverride: () => systemPrompt,
      });
      await resourceLoader.reload();
      const settingsManager = SettingsManager.inMemory({
        defaultTools: [],
        httpIdleTimeoutMs: upstreamTimeoutMs,
        retry: {
          enabled: false,
          provider: { maxRetries: 0, timeoutMs: upstreamTimeoutMs },
        },
      });
      const documentTool = {
        description:
          "Create and validate the requested DOCX with static text and unique tagged text controls.",
        execute: async (
          _toolCallId: string,
          parameters: {
            description: string;
            fields: GeneratedField[];
            paragraphs: string[];
            title: string;
          }
        ) => {
          if (pending.cancelled) {
            throw new Error("AI authoring generation was cancelled");
          }
          if (generated) {
            throw new Error("Only one DOCX can be created in this session");
          }
          const candidate = createTemplateDocx(parameters);
          validateDocument(candidate.document);
          generated = candidate;
          return {
            content: [
              {
                text: `Created ${candidate.title} with ${candidate.fields.length} tagged fields.`,
                type: "text" as const,
              },
            ],
            details: {},
          };
        },
        label: "Create DOCX",
        name: "create_template_docx",
        parameters: toolParameters,
      };
      const created = await createAgentSession({
        agentDir: tempDirectory,
        customTools: [documentTool],
        cwd: tempDirectory,
        model,
        modelRuntime,
        noTools: "builtin",
        resourceLoader,
        sessionManager: SessionManager.inMemory(tempDirectory),
        settingsManager,
        thinkingLevel: "off",
      });
      piSession = created.session;
      pending.piSession = piSession;
      if (pending.cancelled) {
        throw new Error("AI authoring generation was cancelled");
      }
      await piSession.prompt(promptFor(prompt.trim()), {
        expandPromptTemplates: false,
        source: "rpc",
      });
      if (pending.cancelled) {
        throw new Error("AI authoring generation was cancelled");
      }
      if (!generated) {
        throw new Error("OmniRoute did not create a valid DOCX");
      }
      const reply =
        piSession.getLastAssistantText()?.trim() ||
        `Created ${generated.title} with ${generated.fields.length} tagged field${generated.fields.length === 1 ? "" : "s"}.`;
      const session: AuthoringSession = {
        document: generated,
        lastActivity: this.now(),
        ownerSessionId: owner.authSessionId,
        piSession,
        prompt: prompt.trim(),
        reply,
        tempDirectory,
        userId: owner.userId,
      };
      const sessionId = crypto.randomUUID();
      this.sessions.set(sessionId, session);
      this.touch(sessionId, session);
      return this.preview(sessionId, session);
    } catch {
      await piSession?.abort().catch(() => undefined);
      piSession?.dispose();
      generated?.document.fill(0);
      if (tempDirectory) {
        await rm(tempDirectory, { force: true, recursive: true });
      }
      throw new AiAuthoringError(
        502,
        "ai_authoring_failed",
        "OmniRoute could not create a valid DOCX"
      );
    } finally {
      ownerPending.delete(pending);
      if (ownerPending.size === 0) {
        this.pendingGenerations.delete(owner.authSessionId);
      }
      pending.finish();
    }
  }

  async download(
    sessionId: string,
    owner: AuthoringOwner
  ): Promise<{ bytes: Uint8Array; title: string }> {
    const session = await this.sessionFor(sessionId, owner);
    this.touch(sessionId, session);
    return { bytes: session.document.document.slice(), title: session.document.title };
  }

  async end(sessionId: string, owner: AuthoringOwner): Promise<void> {
    const session = await this.sessionFor(sessionId, owner);
    await this.dispose(sessionId, session);
  }

  async endForSession(ownerSessionId: string): Promise<void> {
    this.endedOwnerSessions.add(ownerSessionId);
    const pending = [...(this.pendingGenerations.get(ownerSessionId) ?? [])];
    for (const generation of pending) {
      generation.cancelled = true;
    }
    await Promise.allSettled(
      pending.map((generation) => generation.piSession?.abort())
    );
    await Promise.all(pending.map((generation) => generation.done));
    for (const [sessionId, session] of this.sessions) {
      if (session.ownerSessionId === ownerSessionId) {
        await this.dispose(sessionId, session);
      }
    }
  }
  async close(): Promise<void> {
    for (const ownerSessionId of [...this.pendingGenerations.keys()]) {
      await this.endForSession(ownerSessionId);
    }
    for (const [sessionId, session] of this.sessions) {
      await this.dispose(sessionId, session);
    }
  }

  private async sessionFor(
    sessionId: string,
    owner: AuthoringOwner
  ): Promise<AuthoringSession> {
    const session = this.sessions.get(sessionId);
    if (
      !session ||
      session.userId !== owner.userId ||
      session.ownerSessionId !== owner.authSessionId
    ) {
      throw new AiAuthoringError(
        404,
        "not_found",
        "AI authoring session was not found"
      );
    }
    if (this.now().getTime() - session.lastActivity.getTime() >= sessionLifetimeMs) {
      await this.dispose(sessionId, session);
      throw new AiAuthoringError(
        404,
        "not_found",
        "AI authoring session was not found"
      );
    }
    return session;
  }

  private preview(sessionId: string, session: AuthoringSession): AuthoringPreview {
    return {
      assistantMessage: session.reply,
      description: session.document.description,
      downloadUrl: `/api/admin/ai-authoring/sessions/${sessionId}/docx`,
      fields: session.document.fields.map((field) => ({ ...field })),
      paragraphs: [...session.document.paragraphs],
      sessionId,
      title: session.document.title,
    };
  }

  private touch(sessionId: string, session: AuthoringSession): void {
    session.lastActivity = this.now();
    clearTimeout(session.expiryTimer);
    session.expiryTimer = setTimeout(() => {
      void this.dispose(sessionId, session).catch(() => {
        console.error("Could not remove expired AI authoring files");
      });
    }, sessionLifetimeMs);
    session.expiryTimer.unref();
  }

  private async dispose(sessionId: string, session: AuthoringSession): Promise<void> {
    if (this.sessions.get(sessionId) !== session) {
      return;
    }
    this.sessions.delete(sessionId);
    session.document.document.fill(0);
    session.piSession.dispose();
    session.prompt = "";
    session.reply = "";
    clearTimeout(session.expiryTimer);
    await rm(session.tempDirectory, { force: true, recursive: true });
  }
}
