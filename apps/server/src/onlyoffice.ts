// oxlint-disable func-style no-await-in-loop -- Preserve function contracts; bounded streams must be consumed sequentially.
import { createHmac, timingSafeEqual } from "node:crypto";
import { isDeepStrictEqual } from "node:util";

import { env } from "@onlyoffice/env/server";

export const pluginGuid = "asc.{E0B14962-3B9D-4E82-923E-89A93E8A1A51}";

const tokenLifetimeSeconds = 5 * 60;
const operationCapabilityLifetimeSeconds = tokenLifetimeSeconds + 60;
const maxOnlyOfficeDownloadBytes = 25 * 1024 * 1024;
const jwtHeader = Buffer.from(
  JSON.stringify({ alg: "HS256", typ: "JWT" })
).toString("base64url");
const uuidPattern = /^[0-9a-f-]{36}$/iu;
const editorCapabilityActions = new Set<unknown>([
  "save-template",
  "publish",
  "save-draft",
  "submit",
  "save-correction",
  "configure-fields",
  "poll-operation",
]);
const editorCapabilityTargets = new Set<unknown>([
  "template-draft",
  "response",
  "correction",
]);
const userRoles = new Set<unknown>(["admin", "user"]);
const capabilityStringFields = [
  "actorId",
  "documentKey",
  "formId",
  "targetId",
] as const;
const trustBoundarySecrets = [
  env.BETTER_AUTH_SECRET,
  env.EDITOR_CAPABILITY_SECRET,
  env.PREFILL_HANDOFF_SECRET,
  env.ONLYOFFICE_JWT_SECRET,
] as const;
if (new Set(trustBoundarySecrets).size !== trustBoundarySecrets.length) {
  throw new Error(
    "BETTER_AUTH_SECRET, EDITOR_CAPABILITY_SECRET, PREFILL_HANDOFF_SECRET, and ONLYOFFICE_JWT_SECRET must be distinct"
  );
}
export type EditorCapabilityAction =
  | "save-template"
  | "publish"
  | "save-draft"
  | "submit"
  | "save-correction"
  | "configure-fields"
  | "poll-operation";
export type EditorCapabilityTarget =
  | "template-draft"
  | "response"
  | "correction";
type PluginAction =
  | "template-edit"
  | "fill"
  | "draft"
  | "submit"
  | "correction";
type UserRole = "admin" | "user";
type JsonRecord = Record<string, unknown>;

export interface EditorCapabilityClaims {
  action: EditorCapabilityAction;
  actorId: string;
  documentKey: string;
  expiresAt: number;
  formId: string;
  issuedAt: number;
  isSso?: boolean;
  sessionExpiresAt?: number;
  kind: "editor-capability";
  leaseId?: string;
  leaseProof?: string;
  operationId?: string;
  role: UserRole;
  targetId: string;
  targetType: EditorCapabilityTarget;
}

export interface EditorCapabilityInput {
  action: EditorCapabilityAction;
  actorId: string;
  documentKey: string;
  expiresAt?: number;
  formId: string;
  isSso?: boolean;
  sessionExpiresAt?: number;
  leaseId?: string;
  leaseProof?: string;
  operationId?: string;
  role: UserRole;
  targetId: string;
  targetType: EditorCapabilityTarget;
}

export interface CallbackClaim {
  documentKey: string;
  expiresAt: number;
  kind: "onlyoffice-callback";
  operationId: string;
  operationType: string;
}

export interface CallbackClaimInput {
  documentKey: string;
  expiresAt?: number;
  operationId: string;
  operationType: string;
}

export interface EditorOptions {
  action: PluginAction;
  capabilities: Partial<Record<EditorCapabilityAction, string>>;
  lease: {
    expiresAt: string;
    id: string;
    releaseUrl: string;
    renewUrl: string;
  };
  documentKey: string;
  operationId?: string;
  prefill?: {
    data: Record<string, unknown>;
    editableFields: Record<string, unknown>;
  };
  publicId: string;
  responseId?: string;
}

interface ConverterResponse {
  error?: number;
  endConvert?: boolean;
  fileUrl?: string;
  url?: string;
}

function trimOrigin(origin: string): string {
  return origin.replace(/\/+$/u, "");
}

function exactOrigin(value: string): string {
  return new URL(value).origin;
}

function hmac(value: string, secret: string): Buffer {
  return createHmac("sha256", secret).update(value).digest();
}

function signToken(payload: JsonRecord, secret: string): string {
  const payloadSegment = Buffer.from(JSON.stringify(payload)).toString(
    "base64url"
  );
  const unsignedToken = `${jwtHeader}.${payloadSegment}`;
  return `${unsignedToken}.${hmac(unsignedToken, secret).toString("base64url")}`;
}

function verifiedTokenPayload(
  token: string,
  secret: string
): JsonRecord | null {
  const [headerSegment, payloadSegment, signatureSegment, extra] =
    token.split(".");
  if (!headerSegment || !payloadSegment || !signatureSegment || extra) {
    return null;
  }
  const unsignedToken = `${headerSegment}.${payloadSegment}`;
  const expected = hmac(unsignedToken, secret);
  let provided: Buffer;
  try {
    provided = Buffer.from(signatureSegment, "base64url");
  } catch {
    return null;
  }
  if (
    provided.length !== expected.length ||
    !timingSafeEqual(provided, expected)
  ) {
    return null;
  }
  try {
    const header: unknown = JSON.parse(
      Buffer.from(headerSegment, "base64url").toString("utf-8")
    );
    if (
      !header ||
      typeof header !== "object" ||
      Array.isArray(header) ||
      !("alg" in header) ||
      header.alg !== "HS256" ||
      ("typ" in header && header.typ !== "JWT")
    ) {
      return null;
    }
    const payload: unknown = JSON.parse(
      Buffer.from(payloadSegment, "base64url").toString("utf-8")
    );
    return payload && typeof payload === "object" && !Array.isArray(payload)
      ? (payload as JsonRecord)
      : null;
  } catch {
    return null;
  }
}

function authorizationToken(value: string | null): string | null {
  if (!value) {
    return null;
  }
  return /^Bearer\s+(?<token>[^\s]+)$/iu.exec(value)?.groups?.token ?? null;
}

function expiresIn(seconds: number): number {
  return Math.floor(Date.now() / 1000) + seconds;
}

function liveExpiry(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value > Date.now() / 1000
  );
}
function onlyOfficeTokenTimes(): { exp: number; iat: number } {
  const issuedAt = Math.floor(Date.now() / 1000);
  return { exp: issuedAt + tokenLifetimeSeconds, iat: issuedAt };
}

function onlyOfficePayload(value: JsonRecord): JsonRecord {
  const payload = { ...value };
  delete payload.token;
  return payload;
}

function onlyOfficeSemanticClaims(value: JsonRecord): JsonRecord {
  const payload = { ...value };
  delete payload.exp;
  delete payload.iat;
  return payload;
}

function validOnlyOfficeTokenTimes(claims: JsonRecord): boolean {
  return (
    typeof claims.iat === "number" &&
    Number.isInteger(claims.iat) &&
    claims.iat <= Date.now() / 1000 + 30 &&
    liveExpiry(claims.exp) &&
    claims.exp > claims.iat &&
    claims.exp <= claims.iat + tokenLifetimeSeconds
  );
}

function capabilityLifetimeSeconds(action: unknown): number {
  return action === "poll-operation"
    ? operationCapabilityLifetimeSeconds
    : tokenLifetimeSeconds;
}

function validCapabilityTimes(claims: JsonRecord): boolean {
  return (
    typeof claims.issuedAt === "number" &&
    Number.isInteger(claims.issuedAt) &&
    claims.issuedAt <= Date.now() / 1000 + 30 &&
    liveExpiry(claims.expiresAt) &&
    claims.expiresAt > claims.issuedAt &&
    claims.expiresAt <=
      claims.issuedAt + capabilityLifetimeSeconds(claims.action)
  );
}

function validOperationScope(claims: JsonRecord): boolean {
  if (claims.action === "poll-operation") {
    return (
      typeof claims.operationId === "string" &&
      uuidPattern.test(claims.operationId) &&
      claims.leaseId === undefined &&
      claims.leaseProof === undefined
    );
  }
  return (
    claims.operationId === undefined &&
    typeof claims.leaseId === "string" &&
    uuidPattern.test(claims.leaseId) &&
    typeof claims.leaseProof === "string" &&
    claims.leaseProof.length > 0
  );
}

export function createEditorCapability(input: EditorCapabilityInput): string {
  const now = Math.floor(Date.now() / 1000);
  const maximumExpiry = now + capabilityLifetimeSeconds(input.action);
  return signToken(
    {
      ...input,
      expiresAt:
        input.isSso === true
          ? Math.min(
              input.expiresAt ?? maximumExpiry,
              maximumExpiry,
              input.sessionExpiresAt ?? maximumExpiry
            )
          : (input.expiresAt ?? maximumExpiry),
      isSso: input.isSso === undefined ? false : input.isSso,
      issuedAt: now,
      kind: "editor-capability",
    },
    env.EDITOR_CAPABILITY_SECRET
  );
}

export function verifyEditorCapability(
  token: string
): EditorCapabilityClaims | null {
  const claims = verifiedTokenPayload(token, env.EDITOR_CAPABILITY_SECRET);
  const validSsoMetadata =
    claims?.isSso === true
      ? typeof claims.issuedAt === "number" &&
        typeof claims.expiresAt === "number" &&
        typeof claims.sessionExpiresAt === "number" &&
        Number.isInteger(claims.sessionExpiresAt) &&
        claims.sessionExpiresAt > claims.issuedAt &&
        claims.expiresAt <= claims.sessionExpiresAt
      : claims?.sessionExpiresAt === undefined;
  if (
    !claims ||
    claims.kind !== "editor-capability" ||
    !editorCapabilityActions.has(claims.action) ||
    !userRoles.has(claims.role) ||
    !editorCapabilityTargets.has(claims.targetType) ||
    (claims.isSso !== undefined && typeof claims.isSso !== "boolean") ||
    !capabilityStringFields.every(
      (field) => typeof claims[field] === "string"
    ) ||
    !validCapabilityTimes(claims) ||
    !validSsoMetadata ||
    !validOperationScope(claims)
  ) {
    return null;
  }
  return claims as unknown as EditorCapabilityClaims;
}

export function createDocumentAccessToken(
  documentKey: string,
  expiresAt = expiresIn(tokenLifetimeSeconds)
): string {
  return signToken(
    { documentKey, expiresAt, kind: "onlyoffice-document" },
    env.ONLYOFFICE_JWT_SECRET
  );
}

export function verifyDocumentAccessToken(
  token: string,
  documentKey: string
): boolean {
  const claims = verifiedTokenPayload(token, env.ONLYOFFICE_JWT_SECRET);
  return (
    claims?.kind === "onlyoffice-document" &&
    claims.documentKey === documentKey &&
    liveExpiry(claims.expiresAt)
  );
}

export function createCallbackUserdata(input: CallbackClaimInput): string {
  return signToken(
    {
      documentKey: input.documentKey,
      expiresAt: input.expiresAt ?? expiresIn(tokenLifetimeSeconds),
      kind: "onlyoffice-callback",
      operationId: input.operationId,
      operationType: input.operationType,
    },
    env.ONLYOFFICE_JWT_SECRET
  );
}

export function callbackClaim(value: unknown): CallbackClaim | null {
  if (typeof value !== "string") {
    return null;
  }
  const claims = verifiedTokenPayload(value, env.ONLYOFFICE_JWT_SECRET);
  if (
    claims?.kind !== "onlyoffice-callback" ||
    typeof claims.documentKey !== "string" ||
    typeof claims.operationId !== "string" ||
    !uuidPattern.test(claims.operationId) ||
    typeof claims.operationType !== "string" ||
    !liveExpiry(claims.expiresAt)
  ) {
    return null;
  }
  return claims as unknown as CallbackClaim;
}

export function createOnlyOfficeBodyToken(payload: JsonRecord): string {
  return signToken(
    { ...onlyOfficePayload(payload), ...onlyOfficeTokenTimes() },
    env.ONLYOFFICE_JWT_SECRET
  );
}

export function createOnlyOfficeAuthorization(payload: JsonRecord): string {
  return `Bearer ${signToken(
    {
      payload: onlyOfficePayload(payload),
      ...onlyOfficeTokenTimes(),
    },
    env.ONLYOFFICE_JWT_SECRET
  )}`;
}

export function verifyOnlyOfficeAuthorization(
  authorization: string | null,
  payload: JsonRecord
): boolean {
  const token = authorizationToken(authorization);
  if (!token) {
    return false;
  }
  const claims = verifiedTokenPayload(token, env.ONLYOFFICE_JWT_SECRET);
  const semanticPayload = onlyOfficePayload(payload);
  if (
    !claims ||
    !validOnlyOfficeTokenTimes(claims) ||
    !isDeepStrictEqual(claims.payload, semanticPayload)
  ) {
    return false;
  }
  if (payload.token === undefined) {
    return true;
  }
  if (typeof payload.token !== "string") {
    return false;
  }
  const bodyClaims = verifiedTokenPayload(
    payload.token,
    env.ONLYOFFICE_JWT_SECRET
  );
  return Boolean(
    bodyClaims &&
    validOnlyOfficeTokenTimes(bodyClaims) &&
    isDeepStrictEqual(onlyOfficeSemanticClaims(bodyClaims), semanticPayload)
  );
}

export function documentUrl(documentKey: string): string {
  const token = createDocumentAccessToken(documentKey);
  return `${trimOrigin(env.ONLYOFFICE_DOCUMENT_BASE_URL)}/onlyoffice/document/${encodeURIComponent(documentKey)}?token=${encodeURIComponent(token)}`;
}

export function editorConfig(
  options: EditorOptions,
  user: { id: string; name: string }
): Record<string, unknown> {
  // DOCXF enables ONLYOFFICE's native fixed-form interactions in fill mode.
  const documentFileType =
    options.action === "template-edit" ? "docx" : "docxf";
  const { capabilities, lease, ...pluginOptions } = options;
  const officeServerOrigin = trimOrigin(env.ONLYOFFICE_DOCUMENT_BASE_URL);
  const browserServerOrigin = trimOrigin(env.API_BASE);
  const bridgeId = crypto.randomUUID();
  const config = {
    document: {
      fileType: documentFileType,
      key: options.documentKey,
      permissions: {
        comment: false,
        download: false,
        // Prefill uses the Office API once the document is open. The plugin
        // switches user sessions to the `forms` restriction immediately after
        // applying the server-provided values.
        edit: true,
        fillForms: options.action !== "template-edit",
        review: false,
      },
      title: `${options.action}-${options.publicId}.${documentFileType}`,
      url: documentUrl(options.documentKey),
    },
    documentType: "word",
    editorConfig: {
      ...(options.action === "template-edit"
        ? {}
        : {
            customization: {
              compactToolbar: true,
              hideRightMenu: true,
            },
          }),
      callbackUrl: `${officeServerOrigin}/onlyoffice/callback`,
      mode: "edit",
      plugins: {
        autostart: [pluginGuid],
        disable: ["asc.{9DC93CDB-B576-4F0C-B55E-FCC9C48DD007}"],
        options: {
          [pluginGuid]: {
            ...pluginOptions,
            apiBase: browserServerOrigin,
            bridgeId,
            parentOrigin: exactOrigin(env.CORS_ORIGIN),
          },
        },
        pluginsData: [`${browserServerOrigin}/onlyoffice-plugin/config.json`],
      },
      user: {
        id: user.id,
        name: user.name,
      },
    },
    height: "100%",
    width: "100%",
  };
  return {
    apiUrl: trimOrigin(env.ONLYOFFICE_URL),
    bridge: {
      capabilities,
      id: bridgeId,
      lease,
      pluginOrigin: exactOrigin(env.API_BASE),
    },
    config: { ...config, token: signToken(config, env.ONLYOFFICE_JWT_SECRET) },
  };
}

export function readOnlyViewerConfig(
  documentKey: string,
  publicId: string,
  user: { id: string; name: string }
): Record<string, unknown> {
  const config = {
    document: {
      fileType: "docx",
      key: documentKey,
      permissions: {
        comment: false,
        download: false,
        edit: false,
        fillForms: false,
        review: false,
      },
      title: `submission-${publicId}.docx`,
      url: documentUrl(documentKey),
    },
    documentType: "word",
    editorConfig: {
      customization: { compactToolbar: true, hideRightMenu: true },
      mode: "view",
      user,
    },
    height: "100%",
    width: "100%",
  };
  return {
    apiUrl: trimOrigin(env.ONLYOFFICE_URL),
    config: { ...config, token: signToken(config, env.ONLYOFFICE_JWT_SECRET) },
  };
}


export interface OnlyOfficeClient {
  forceSave: (documentKey: string, userdata: string) => Promise<boolean>;
  convertDocxToPdf: (documentKey: string) => Promise<Uint8Array>;
}

export interface OnlyOfficeClientOptions {
  documentBaseUrl?: string;
  fetch?: typeof fetch;
  internalUrl?: string;
  maxDownloadBytes?: number;
}

function documentUrlFor(documentKey: string, baseUrl: string): string {
  const token = createDocumentAccessToken(documentKey);
  return `${trimOrigin(baseUrl)}/onlyoffice/document/${encodeURIComponent(documentKey)}?token=${encodeURIComponent(token)}`;
}

function converterResponseError(payload: unknown): string | null {
  if (!payload || typeof payload !== "object" || !("error" in payload)) {
    return null;
  }
  const { error } = payload;
  if (error === undefined || error === 0) {
    return null;
  }
  return `ONLYOFFICE conversion failed (${String(error)})`;
}

function onlyOfficeResultUrl(value: string, internalUrl: string): string {
  let result: URL;
  try {
    result = new URL(value);
  } catch {
    throw new Error("ONLYOFFICE returned an invalid download URL");
  }
  if (
    (result.protocol !== "http:" && result.protocol !== "https:") ||
    result.origin !== exactOrigin(internalUrl)
  ) {
    throw new Error("ONLYOFFICE returned a disallowed download URL");
  }
  return result.toString();
}

async function boundedResponseBytes(
  response: Response,
  maximumBytes: number
): Promise<Uint8Array> {
  const contentLength = Number(response.headers.get("content-length") ?? "");
  if (Number.isFinite(contentLength) && contentLength > maximumBytes) {
    throw new Error("ONLYOFFICE download is too large");
  }
  const reader = response.body?.getReader();
  if (!reader) {
    return new Uint8Array();
  }
  const chunks: Uint8Array[] = [];
  let byteLength = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    byteLength += value.byteLength;
    if (byteLength > maximumBytes) {
      await reader.cancel();
      throw new Error("ONLYOFFICE download is too large");
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(byteLength);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

export function createOnlyOfficeClient(
  options: OnlyOfficeClientOptions = {}
): OnlyOfficeClient {
  const request = options.fetch ?? fetch;
  const internalUrl = options.internalUrl ?? env.ONLYOFFICE_INTERNAL_URL;
  const documentBaseUrl =
    options.documentBaseUrl ?? env.ONLYOFFICE_DOCUMENT_BASE_URL;
  const maximumDownloadBytes =
    options.maxDownloadBytes ?? maxOnlyOfficeDownloadBytes;

  return {
    async convertDocxToPdf(documentKey): Promise<Uint8Array> {
      const endpoint = `${trimOrigin(internalUrl)}/converter?shardkey=${encodeURIComponent(documentKey)}`;
      const conversion = {
        async: false,
        filetype: "docx",
        key: documentKey,
        outputtype: "pdf",
        title: `${documentKey}.docx`,
        url: documentUrlFor(documentKey, documentBaseUrl),
      };
      const response = await request(endpoint, {
        body: JSON.stringify(conversion),
        headers: {
          Authorization: createOnlyOfficeAuthorization(conversion),
          "Content-Type": "application/json",
        },
        method: "POST",
      });
      if (!response.ok) {
        throw new Error(
          `ONLYOFFICE converter returned HTTP ${response.status}`
        );
      }

      const payload = (await response.json()) as ConverterResponse;
      const conversionError = converterResponseError(payload);
      if (conversionError) {
        throw new Error(conversionError);
      }
      const outputUrlValue = payload.fileUrl ?? payload.url;
      if (!outputUrlValue) {
        throw new Error("ONLYOFFICE conversion did not return a PDF URL");
      }
      const outputUrl = onlyOfficeResultUrl(outputUrlValue, internalUrl);
      const pdfResponse = await request(outputUrl, {
        headers: {
          Authorization: createOnlyOfficeAuthorization({ url: outputUrl }),
        },
        redirect: "error",
      });
      if (!pdfResponse.ok) {
        throw new Error(
          `Failed to download converted PDF: HTTP ${pdfResponse.status}`
        );
      }
      return boundedResponseBytes(pdfResponse, maximumDownloadBytes);
    },

    async forceSave(documentKey, userdata): Promise<boolean> {
      const endpoint = `${trimOrigin(internalUrl)}/command?shardkey=${encodeURIComponent(documentKey)}`;
      const command = { c: "forcesave", key: documentKey, userdata };
      const response = await request(endpoint, {
        body: JSON.stringify(command),
        headers: {
          Authorization: createOnlyOfficeAuthorization(command),
          "Content-Type": "application/json",
        },
        method: "POST",
      });
      if (!response.ok) {
        throw new Error(
          `ONLYOFFICE command service returned HTTP ${response.status}`
        );
      }
      const payload: unknown = await response.json();
      if (!payload || typeof payload !== "object") {
        throw new Error("ONLYOFFICE command service returned invalid JSON");
      }
      const error = "error" in payload ? payload.error : undefined;
      if (error === 4) {
        return false;
      }
      if (error !== undefined && error !== 0) {
        const message =
          "message" in payload && typeof payload.message === "string"
            ? payload.message
            : `ONLYOFFICE command failed (${String(error)})`;
        throw new Error(message);
      }
      return true;
    },
  };
}

const defaultOnlyOfficeClient = createOnlyOfficeClient();

export function forceSave(
  documentKey: string,
  userdata: string
): Promise<boolean> {
  return defaultOnlyOfficeClient.forceSave(documentKey, userdata);
}

export function convertDocxToPdf(documentKey: string): Promise<Uint8Array> {
  return defaultOnlyOfficeClient.convertDocxToPdf(documentKey);
}
