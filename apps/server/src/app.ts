// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import {
  createHash,
  createHmac,
  randomBytes,
  timingSafeEqual,
} from "node:crypto";
import path from "node:path";
import { isDeepStrictEqual } from "node:util";

import { cors } from "@elysiajs/cors";
import { auth } from "@onlyoffice/auth";
import type { OperationType } from "@onlyoffice/db";
import {
  AuditOutcome,
  FieldType,
  FillMethod,
  FormStatus,
  HandoffStatus,
  LegacyAccountLinkStatus,
  OperationStatus,
  OperationTargetType,
  PrefillPolicy,
  Prisma,
  ResponseStatus,
  prisma,
} from "@onlyoffice/db";
import { env } from "@onlyoffice/env/server";
import { Elysia } from "elysia";
import { strToU8, unzipSync, zipSync } from "fflate";
import { SaxesParser } from "saxes";

import {
  AiAuthoringSessions,
  authoringDisclosure,
} from "./ai-authoring";
import type { GeneratedTemplate, OmniRouteConfig } from "./ai-authoring";
import { AiAuthoringError } from "./ai-authoring-error";
import {
  callbackClaim,
  createCallbackUserdata,
  createEditorCapability,
  createOnlyOfficeAuthorization,
  createOnlyOfficeClient,
  editorConfig,
  pluginGuid,
  readOnlyViewerConfig,
  verifyDocumentAccessToken,
  verifyEditorCapability,
  verifyOnlyOfficeAuthorization,
} from "./onlyoffice";
import type {
  EditorCapabilityAction,
  EditorCapabilityClaims,
  EditorCapabilityTarget,
  OnlyOfficeClient,
} from "./onlyoffice";
import {
  DOCX_CONTENT_TYPE,
  deleteObject,
  objectExists,
  objectKey,
  putObject,
  readObject,
  streamObject,
} from "./storage";

type Form = Prisma.FormGetPayload<Prisma.FormDefaultArgs>;
type Operation = Prisma.OperationGetPayload<Prisma.OperationDefaultArgs>;
type PrefillSnapshot =
  Prisma.PrefillSnapshotGetPayload<Prisma.PrefillSnapshotDefaultArgs>;
type PublishedTemplate =
  Prisma.PublishedTemplateGetPayload<Prisma.PublishedTemplateDefaultArgs>;
type Response = Prisma.ResponseGetPayload<Prisma.ResponseDefaultArgs>;
type Correction = Prisma.CorrectionGetPayload<Prisma.CorrectionDefaultArgs>;
type Submission = Prisma.SubmissionGetPayload<Prisma.SubmissionDefaultArgs>;
type SubmissionWithManifest = Prisma.SubmissionGetPayload<{
  include: {
    corrections: true;
    form: true;
    owner: true;
    response: {
      include: {
        publishedTemplate: {
          include: {
            manifest: {
              include: { fields: true };
            };
          };
        };
      };
    };
  };
}>;
type SubmissionWithRevisions = Prisma.SubmissionGetPayload<{
  include: { corrections: true; form: true; owner: true };
}>;
type TemplateDraft =
  Prisma.TemplateDraftGetPayload<Prisma.TemplateDraftDefaultArgs>;
type DraftFieldRule =
  Prisma.DraftFieldRuleGetPayload<Prisma.DraftFieldRuleDefaultArgs>;

function originOf(value: string): string | null {
  try {
    return new URL(value).origin;
  } catch {
    return null;
  }
}
function configuredPrefillReturnUrl(value: string): string {
  const url = new URL(value);
  if (
    !["http:", "https:"].includes(url.protocol) ||
    url.username ||
    url.password
  ) {
    throw new Error(
      "PREFILL_RETURN_URL must be an HTTP(S) URL without credentials"
    );
  }
  return url.toString();
}
const onlyOfficePluginSdkUrlPlaceholder = "__ONLYOFFICE_PLUGIN_SDK_URL__";
const htmlEscape = (value: string): string =>
  value.replaceAll(
    /[&<>"']/gu,
    (character) =>
      ({
        '"': "&quot;",
        "&": "&amp;",
        "'": "&#39;",
        "<": "&lt;",
        ">": "&gt;",
      })[character] ?? character
  );

const pluginDir = path.resolve(import.meta.dirname, "../../onlyoffice-plugin");
const pluginIndexResponse = async () => {
  const html = await Bun.file(path.resolve(pluginDir, "index.html")).text();
  const onlyOfficeBaseUrl = env.ONLYOFFICE_URL.replace(/\/+$/u, "");
  const sdkUrl = `${onlyOfficeBaseUrl}/sdkjs-plugins/v1/plugins.js`;
  return new Response(
    html.replace(onlyOfficePluginSdkUrlPlaceholder, htmlEscape(sdkUrl)),
    { headers: { "content-type": "text/html; charset=utf-8" } }
  );
};
const fallbackTemplatePath = path.resolve(
  import.meta.dirname,
  "../../../onlyoffice-templates/template.docx"
);
const idPattern = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/iu;
const publicIdPattern = /^[0-9a-f]{32}$/u;
const accountEmailPattern = /^[^\s@]+@[^\s@]+$/iu;
const operationTimeoutMs = 4 * 60_000;
const editorLeaseDurationMs = 90_000;
const callbackClaimLifetimeSeconds = 5 * 60;
const objectCleanupIntentGraceMs = 15 * 60_000;
const handoffCodeLifetimeMs = 120_000;
const pendingClaimLifetimeSeconds = 10 * 60;
const legacySsoTransactionLifetimeSeconds = 5 * 60;
const legacySsoSessionTransferLifetimeSeconds = 60;
const legacySsoExchangeBodyMaximumBytes = 8 * 1024;
const legacySsoCookieName = "__Host-folio-sso";
const legacySsoSessionCookieName = "__Host-folio-sso-session";
const legacySsoSwitchCookieName = "__Host-folio-sso-switch";
const legacySsoTransactionIdentifier = "legacy-sso-transaction";
const legacySsoSessionIdentifier = "legacy-sso-session";
const legacySsoSwitchIdentifier = "legacy-sso-switch";

const handoffExpirySweepBatchSize = 100;
const handoffCodeMaximumLength = 256;
const handoffExternalReferenceMaximumLength = 512;
const prefillHandoffBodyMaximumBytes = 512 * 1024;
const maxTemplateUploadBytes = 25 * 1024 * 1024;
const maxTemplateMultipartOverheadBytes = 64 * 1024;
const maxTemplateMultipartBodyBytes =
  maxTemplateUploadBytes + maxTemplateMultipartOverheadBytes;
const maxAuthoringPdfBytes = 10 * 1024 * 1024;
const maxAuthoringMultipartBodyBytes = 11 * 1024 * 1024;
const maxTemplateArchiveExpandedBytes = 64 * 1024 * 1024;
const maxTemplateArchiveEntries = 2048;
const maxCallbackDocumentBytes = maxTemplateUploadBytes;
const maxCallbackBodyBytes = 64 * 1024;
const loginFailureLimit = 5;
const loginFailureWindowMs = 15 * 60_000;
const passwordMinimumLength = 12;
const passwordMaximumLength = 128;
const correctionReasonMaximumLength = 2000;
const maxResponseDataBytes = 256 * 1024;
const maxNativePictureMultipartOverheadBytes =
  maxResponseDataBytes + maxTemplateMultipartOverheadBytes;
const maxResponseTextLength = 10_000;
const fieldRuleBodyMaximumBytes = 8 * 1024;
const dateFieldPattern = /^\d{4}-\d{2}-\d{2}$/u;
const isValidDateFieldValue = (value: string): boolean => {
  if (!dateFieldPattern.test(value)) {
    return false;
  }
  const parsed = new Date(`${value}T00:00:00.000Z`);
  return (
    !Number.isNaN(parsed.getTime()) &&
    parsed.toISOString().slice(0, 10) === value
  );
};
const fieldTagMaximumLength = 512;
const fieldPointerMaximumLength = 2048;
const schemaPageSize = 5;
const schemaQueryMaximumLength = 200;
const accountUserPageSize = 20;
const adminResultPageSize = 25;
const auditEventPageSize = 50;
const accountBodyMaximumBytes = 64 * 1024;
const documentActionBodyMaximumBytes = 8 * 1024;
const accountEmailMaximumLength = 254;
const accountNameMaximumLength = 120;
// ponytail: one global account lock caps mutation throughput; shard locks only if needed.
const accountMutationLockId = 1_604_619_418;
const callbackInternalOrigin = originOf(env.ONLYOFFICE_INTERNAL_URL);
const callbackPublicOrigin = originOf(env.ONLYOFFICE_URL);
const callbackDocumentOrigin = originOf(env.ONLYOFFICE_DOCUMENT_BASE_URL);
const callbackOrigins = new Set(
  [callbackInternalOrigin, callbackPublicOrigin, callbackDocumentOrigin].filter(
    (origin): origin is string => Boolean(origin)
  )
);
const corsOrigin = new URL(env.CORS_ORIGIN).origin;
const pluginOrigins = new Set(
  [env.API_BASE, env.ONLYOFFICE_URL, env.CORS_ORIGIN]
    .map(originOf)
    .filter((origin): origin is string => Boolean(origin))
);

type UserRole = "admin" | "user";
export interface LegacySsoConfig {
  authorizeUrl: string;
  callbackUrl: string;
  clientId: string;
  clientSecret: string;
  exchangeUrl: string;
  providerId: string;
}
interface LegacySsoTransaction {
  callbackUrl: string;
  clientId: string;
  codeVerifier: string;
  generation: string;
  initiatingSessionId?: string;
  initiatingUserId?: string;
  returnTo: string;
  state: string;
}
interface LegacySsoSwitchChallenge {
  initiatingSessionId: string;
  initiatingUserId: string;
  returnTo: string;
  targetUserId: string;
}
const legacySsoSwitchLifetimeSeconds = 5 * 60;
interface LegacyIdentity {
  email: string;
  subject: string;
}
interface Identity {
  email: string;
  expiresAt: Date;
  id: string;
  isSso: boolean;
  mustChangePassword: boolean;
  name: string;
  role: UserRole;
  sessionId: string;
}
type LegacySsoAccountUser = Pick<Identity, "id" | "role"> & {
  enabled: boolean;
};
type LegacySsoAccountResolution =
  | { kind: "failed" }
  | { kind: "linked"; user: LegacySsoAccountUser }
  | { kind: "pending" };
type Actor = Pick<
  Identity,
  "email" | "id" | "mustChangePassword" | "name" | "role"
>;
type EditorCapabilityIdentity = Actor & Pick<Identity, "expiresAt" | "isSso">;
type AccountAuditAction =
  | "create_user"
  | "delete_user"
  | "enable_user"
  | "disable_user"
  | "change_user_email"
  | "promote_user"
  | "demote_user"
  | "reset_user_password"
  | "update_user"
  | "approve_legacy_account_link"
  | "reject_legacy_account_link";
const adminUserSelect = {
  createdAt: true,
  email: true,
  enabled: true,
  id: true,
  mustChangePassword: true,
  name: true,
  role: true,
  updatedAt: true,
} as const;
type AdminUser = Prisma.UserGetPayload<{ select: typeof adminUserSelect }>;
interface EditorAuthorization {
  actor: EditorCapabilityIdentity;
  capability: EditorCapabilityClaims | null;
}
type ActionEditorAuthorization = EditorAuthorization & {
  capability: EditorCapabilityClaims;
};
interface EditorCapabilityScope {
  action: EditorCapabilityAction;
  documentKey: string;
  formId: string;
  operationId?: string;
  targetId: string;
  targetType: EditorCapabilityTarget;
}
interface ClaimedEditorLease {
  expiresAt: Date;
  id: string;
  workspaceBaseDocumentKey?: string;
  workspaceBaseRevision?: number;
  workspaceDocumentKey?: string;
  workspaceObjectKey?: string;
}
interface CorrectionWorkspaceInput {
  baseDocumentKey: string;
  baseRevision: number;
  documentKey: string;
  objectKey: string;
}
interface EditorLeaseGrant extends ClaimedEditorLease {
  proof: string;
}
interface ActiveEditorLease {
  capabilityDigest: string;
  holderSessionId: string;
  holderUserId: string;
}
type JsonRecord = Record<string, unknown>;
type ExternalSchemaType = "string" | "number" | "boolean" | "null";
interface ExternalSchemaItem {
  pointer: string;
  type: ExternalSchemaType;
}

const externalMockSchema = {
  account: {
    active: true,
    address: {
      city: "",
      country: "",
      postalCode: "",
    },
    consent: {
      privacy: true,
      terms: true,
    },
    contact: {
      email: "",
      phone: "",
    },
    contacts: [{ name: "" }],
    "display/name": "",
    id: "",
    loginCount: 0,
    score: 0,
    settings: {
      language: "",
      timezone: "",
    },
    "tilde~key": "",
  },
  ignoredObject: {
    nested: {
      value: null,
    },
  },
  person: {
    birthDate: "",
    "contact/details": {
      "line~1": "",
    },
    name: "",
  },
} as const;

function externalSchemaPointerSegment(value: string): string {
  return value.replaceAll("~", "~0").replaceAll("/", "~1");
}

function externalSchemaLeafType(value: unknown): ExternalSchemaType | null {
  if (value === null) {
    return "null";
  }
  if (typeof value === "string") {
    return "string";
  }
  if (typeof value === "number") {
    return "number";
  }
  if (typeof value === "boolean") {
    return "boolean";
  }
  return null;
}

function flattenExternalSchema(
  value: unknown,
  parentPointer = "",
  items: ExternalSchemaItem[] = []
): ExternalSchemaItem[] {
  const type = externalSchemaLeafType(value);
  if (type) {
    items.push({ pointer: parentPointer, type });
    return items;
  }
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return items;
  }
  for (const key of Object.keys(value).toSorted()) {
    const pointer = `${parentPointer}/${externalSchemaPointerSegment(key)}`;
    flattenExternalSchema(
      (value as Record<string, unknown>)[key],
      pointer,
      items
    );
  }
  return items;
}

const externalSchemaItems = flattenExternalSchema(externalMockSchema);

function schemaCursor(query: string, offset: number): string {
  const payload = Buffer.from(JSON.stringify({ offset, query })).toString(
    "base64url"
  );
  const unsigned = `schema-v1.${payload}`;
  const signature = createHmac("sha256", env.EDITOR_CAPABILITY_SECRET)
    .update(unsigned)
    .digest("base64url");
  return `${unsigned}.${signature}`;
}

function schemaCursorOffset(cursor: string, query: string): number {
  const [version, payload, signature, extra] = cursor.split(".");
  if (!version || !payload || !signature || extra || version !== "schema-v1") {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  const unsigned = `${version}.${payload}`;
  const expected = createHmac("sha256", env.EDITOR_CAPABILITY_SECRET)
    .update(unsigned)
    .digest("base64url");
  if (signature !== expected) {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  let decoded: unknown;
  try {
    decoded = JSON.parse(Buffer.from(payload, "base64url").toString("utf-8"));
  } catch {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  if (
    !decoded ||
    typeof decoded !== "object" ||
    Array.isArray(decoded) ||
    !("offset" in decoded) ||
    !("query" in decoded) ||
    typeof decoded.offset !== "number" ||
    !Number.isInteger(decoded.offset) ||
    decoded.offset < 0 ||
    typeof decoded.query !== "string" ||
    decoded.query !== query ||
    decoded.offset > externalSchemaItems.length
  ) {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  return decoded.offset;
}

function schemaQueryValue(value: unknown): string {
  if (value === undefined) {
    return "";
  }
  if (typeof value !== "string") {
    fail(400, "invalid_schema_query", "Schema query must be a string");
  }
  const query = value.trim().toLowerCase();
  if (query.length > schemaQueryMaximumLength) {
    fail(400, "invalid_schema_query", "Schema query is too long");
  }
  return query;
}

function schemaPage(
  queryValue: unknown,
  cursorValue: unknown
): { items: ExternalSchemaItem[]; nextCursor: string | null } {
  const query = schemaQueryValue(queryValue);
  const matching = externalSchemaItems.filter((item) =>
    item.pointer.toLowerCase().includes(query)
  );
  const offset =
    cursorValue === undefined
      ? 0
      : typeof cursorValue === "string"
        ? schemaCursorOffset(cursorValue, query)
        : fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  if (offset > matching.length) {
    fail(400, "invalid_schema_cursor", "Schema cursor is invalid");
  }
  const items = matching.slice(offset, offset + schemaPageSize);
  const nextOffset = offset + items.length;
  return {
    items,
    nextCursor:
      nextOffset < matching.length ? schemaCursor(query, nextOffset) : null,
  };
}
type OperationAction =
  | "save-template"
  | "publish"
  | "save-draft"
  | "submit"
  | "save-correction";
const operationTypeForAction: Record<OperationAction, OperationType> = {
  publish: "publish_form",
  "save-correction": "save_correction",
  "save-draft": "save_draft",
  "save-template": "save_template_draft",
  submit: "submit_response",
};
type OperationErrorCode =
  | "callback_claim_invalid"
  | "callback_document_unavailable"
  | "callback_key_mismatch"
  | "callback_processing_failed"
  | "document_save_failed"
  | "fill_method_changed"
  | "force_save_failed"
  | "invalid_template"
  | "onlyoffice_document_error"
  | "operation_timeout"
  | "pdf_conversion_failed";
type FormSource = "blank" | "upload";
type FormAuditAction =
  | "archive_form"
  | "configure_field_rule"
  | "create_form"
  | "create_handoff"
  | "delete_form"
  | "duplicate_form"
  | "launch_handoff"
  | "publish_form"
  | "redeem_handoff"
  | "save_template_draft"
  | "unarchive_form"
  | "update_form_metadata";
type FormAuditErrorCode =
  | "callback_claim_invalid"
  | "callback_document_unavailable"
  | "callback_key_mismatch"
  | "document_save_failed"
  | "fill_method_changed"
  | "native_fill_unsupported"
  | "callback_processing_failed"
  | "document_unavailable"
  | "editor_capability_required"
  | "editor_capability_scope_mismatch"
  | "editor_in_use"
  | "editor_lease_inactive"
  | "force_save_failed"
  | "invalid_template"
  | "form_has_responses"
  | "form_not_draft"
  | "handoff_unavailable"
  | "internal_error"
  | "invalid_editor_capability"
  | "invalid_file_type"
  | "invalid_request"
  | "not_found"
  | "onlyoffice_document_error"
  | "operation_in_progress"
  | "operation_timeout"
  | "pdf_conversion_failed"
  | "payload_too_large"
  | "published_immutable"
  | "stale_document"
  | "stale_operation"
  | "unauthorized";
interface FormAuditMetadata {
  errorCode?: FormAuditErrorCode;
  source?: FormSource;
  sourcePublicId?: string;
  status?: FormStatus;
  fillMethod?: FillMethod;
}
interface OperationMetadata extends JsonRecord {
  action: OperationAction;
  baseDocumentKey?: string;
  baseRevision?: number;
  callbackEmptyFieldTags?: string[];
  cleanupObjectKeys?: string[];
  correctionId?: string;
  data?: JsonRecord;
  finalObjectKey: string;
  formId: string;
  nativeDocumentStaged?: boolean;
  nextDocumentKey?: string;
  publicId?: string;
  publishedKey?: string;
  publishedVersion?: number;
  reason?: string;
  responseId?: string;
  result?: JsonRecord;
  serverHeldFieldTags?: string[];
  stagedObjectKey: string;
  submissionDocumentKey?: string;
  submissionId?: string;
  workspaceDocumentKey?: string;
  workspaceObjectKey?: string;
}

interface OperationCompletion {
  cleanupObjectKeys: string[];
}

type FormWithDocuments = Form & {
  templateDraft: TemplateDraft | null;
  publishedTemplate: PublishedTemplate | null;
};

type ResponseWithSnapshot = Response & {
  prefillSnapshot: PrefillSnapshot | null;
};
interface CallbackPayload {
  key?: unknown;
  status?: unknown;
  url?: unknown;
  userdata?: unknown;
}

export interface AppOptions {
  clock?: () => Date;
  deleteObject?: (key: string) => Promise<void>;
  putObject?: typeof putObject;
  legacySso?: LegacySsoConfig | null;
  omniRoute?: OmniRouteConfig | null;
  onlyOffice?: OnlyOfficeClient;
  onlyOfficeCallbackOrigins?: readonly string[];
  onlyOfficeCallbackMaxBytes?: number;
  prefillHandoffSecret?: string;
  prefillReturnUrl?: string;
  requestIp?: (request: Request) => string | null | undefined;
}

class HttpError extends Error {
  readonly httpStatus: number;
  readonly code: string;

  constructor(status: number, code: string, message: string) {
    super(message);
    this.name = "HttpError";
    this.httpStatus = status;
    this.code = code;
  }
}

function fail(status: number, code: string, message: string): never {
  throw new HttpError(status, code, message);
}

function asRecord(
  value: unknown,
  message = "Request body must be a JSON object"
): JsonRecord {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail(400, "invalid_request", message);
  }
  return value as JsonRecord;
}
async function readRequestBytes(
  request: Request,
  maximumBytes: number,
  missingMessage = "Request body is required"
): Promise<Uint8Array> {
  const contentLength = Number(request.headers.get("content-length") ?? "");
  if (Number.isFinite(contentLength) && contentLength > maximumBytes) {
    fail(413, "payload_too_large", "Request body is too large");
  }
  const reader = request.body?.getReader();
  if (!reader) {
    fail(400, "invalid_request", missingMessage);
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
      fail(413, "payload_too_large", "Request body is too large");
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

async function readJsonRecord(
  request: Request,
  maximumBytes: number
): Promise<JsonRecord> {
  const bytes = await readRequestBytes(
    request,
    maximumBytes,
    "Request body must be a JSON object"
  );
  try {
    return asRecord(JSON.parse(new TextDecoder().decode(bytes)));
  } catch {
    fail(400, "invalid_request", "Request body must be valid JSON");
  }
}

async function readAuthoringCreationInput(
  request: Request
): Promise<{ consent: unknown; pdfBytes?: Uint8Array; prompt: string }> {
  const contentType = request.headers.get("content-type");
  if (!contentType || !/^multipart\/form-data(?:\s*;|$)/iu.test(contentType)) {
    const input = await readJsonRecord(request, accountBodyMaximumBytes);
    if (
      Object.keys(input).length !== 2 ||
      !Object.hasOwn(input, "consent") ||
      !Object.hasOwn(input, "prompt") ||
      typeof input.prompt !== "string"
    ) {
      fail(400, "invalid_request", "Only consent and prompt are accepted");
    }
    return { consent: input.consent, prompt: input.prompt };
  }

  const requestBytes = await readRequestBytes(
    request,
    maxAuthoringMultipartBodyBytes,
    "Multipart form data is required"
  );
  const multipartRequest = new Request(request.url, {
    body: requestBytes,
    headers: { "content-type": contentType },
    method: "POST",
  });
  const formData = await multipartRequest.formData().catch(() => {
    fail(400, "invalid_request", "Multipart form data is invalid");
  });
  const entries = new Map<string, unknown>();
  for (const [key, value] of formData.entries()) {
    if (key !== "prompt" && key !== "consent" && key !== "pdf") {
      fail(400, "invalid_request", "Only prompt, consent, and pdf are accepted");
    }
    if (entries.has(key)) {
      fail(400, "invalid_request", `${key} must be provided once`);
    }
    entries.set(key, value);
  }
  const prompt = entries.get("prompt");
  const consent = entries.get("consent");
  const pdf = entries.get("pdf");
  if (typeof prompt !== "string" || (consent !== undefined && typeof consent !== "string")) {
    fail(400, "invalid_request", "prompt and consent must be text");
  }
  if (!(pdf instanceof File)) {
    fail(400, "invalid_request", "pdf is required");
  }
  const normalizedType = pdf.type.trim().toLowerCase();
  if (
    normalizedType !== "" &&
    normalizedType !== "application/octet-stream" &&
    normalizedType !== "application/pdf"
  ) {
    fail(415, "invalid_file_type", "Source must be a PDF file");
  }
  if (pdf.size > maxAuthoringPdfBytes) {
    fail(413, "payload_too_large", "Source PDF is too large");
  }
  const pdfBytes = new Uint8Array(await pdf.arrayBuffer());
  if (
    pdfBytes[0] !== 0x25 ||
    pdfBytes[1] !== 0x50 ||
    pdfBytes[2] !== 0x44 ||
    pdfBytes[3] !== 0x46 ||
    pdfBytes[4] !== 0x2d
  ) {
    fail(415, "invalid_file_type", "Source must be a PDF file");
  }
  return { consent: consent === "true", pdfBytes, prompt };
}

interface TemplateCreationInput {
  description: string;
  source: FormSource;
  templateBytes?: Uint8Array;
  title: string;
}

async function readTemplateCreationInput(
  request: Request
): Promise<TemplateCreationInput> {
  const contentType = request.headers.get("content-type");
  if (!contentType || !/^multipart\/form-data(?:\s*;|$)/iu.test(contentType)) {
    fail(
      415,
      "invalid_file_type",
      "Form creation requires multipart form data"
    );
  }
  const requestBytes = await readRequestBytes(
    request,
    maxTemplateMultipartBodyBytes,
    "Multipart form data is required"
  );
  const multipartRequest = new Request(request.url, {
    body: requestBytes,
    headers: { "content-type": contentType },
    method: "POST",
  });
  let formData: Awaited<ReturnType<typeof multipartRequest.formData>>;
  try {
    formData = await multipartRequest.formData();
  } catch {
    fail(400, "invalid_request", "Multipart form data is invalid");
  }

  const entries = new Map<string, unknown>();
  for (const [key, value] of formData.entries()) {
    if (
      key !== "description" &&
      key !== "source" &&
      key !== "template" &&
      key !== "title"
    ) {
      fail(
        400,
        "invalid_request",
        "Only title, description, source, and template are accepted"
      );
    }
    if (entries.has(key)) {
      fail(400, "invalid_request", `${key} must be provided once`);
    }
    entries.set(key, value);
  }

  const titleValue = entries.get("title");
  const sourceValue = entries.get("source");
  if (typeof titleValue !== "string" || titleValue.trim().length === 0) {
    fail(400, "invalid_request", "title is required");
  }
  if (typeof sourceValue !== "string" || sourceValue.trim().length === 0) {
    fail(400, "invalid_request", "source is required");
  }
  const title = titleValue.trim();
  const source = sourceValue.trim();
  if (title.length > 200) {
    fail(400, "invalid_request", "Title is too long");
  }
  if (source !== "blank" && source !== "upload") {
    fail(400, "invalid_request", "source must be blank or upload");
  }
  const descriptionValue = entries.get("description");
  if (
    descriptionValue !== undefined &&
    (typeof descriptionValue !== "string" ||
      descriptionValue.trim().length > 2000)
  ) {
    fail(400, "invalid_request", "Description is too long");
  }
  const description =
    typeof descriptionValue === "string" ? descriptionValue.trim() : "";
  const template = entries.get("template");
  if (source === "blank") {
    if (template !== undefined) {
      fail(400, "invalid_request", "template is only accepted for upload");
    }
    return { description, source, title };
  }
  if (!(template instanceof File)) {
    fail(400, "invalid_request", "template is required for upload");
  }
  if (!/\.docx$/iu.test(template.name)) {
    fail(415, "invalid_file_type", "Template must be a DOCX file");
  }
  const normalizedType = template.type.trim().toLowerCase();
  if (
    normalizedType !== "" &&
    normalizedType !== "application/octet-stream" &&
    normalizedType !== DOCX_CONTENT_TYPE
  ) {
    fail(415, "invalid_file_type", "Template must be a DOCX file");
  }
  if (template.size > maxTemplateUploadBytes) {
    fail(413, "payload_too_large", "Template upload is too large");
  }
  const templateBytes = new Uint8Array(await template.arrayBuffer());
  if (templateBytes.byteLength > maxTemplateUploadBytes) {
    fail(413, "payload_too_large", "Template upload is too large");
  }
  return { description, source, templateBytes, title };
}

function requiredString(record: JsonRecord, key: string): string {
  const value = record[key];
  if (typeof value !== "string" || value.trim().length === 0) {
    fail(400, "invalid_request", `${key} is required`);
  }
  return value.trim();
}

async function readDocumentKeyInput(request: Request): Promise<string> {
  const input = await readJsonRecord(request, documentActionBodyMaximumBytes);
  if (Object.keys(input).length !== 1 || !Object.hasOwn(input, "documentKey")) {
    fail(400, "invalid_request", "Only documentKey is accepted");
  }

  return requiredString(input, "documentKey");
}
interface CorrectionInput {
  data: JsonRecord;
  documentKey: string;
  reason: string;
}

function correctionInput(body: unknown): CorrectionInput {
  const input = asRecord(body);
  const keys = Object.keys(input);
  if (
    keys.length !== 3 ||
    !keys.includes("data") ||
    !keys.includes("documentKey") ||
    !keys.includes("reason")
  ) {
    fail(400, "invalid_request", "documentKey, data, and reason are required");
  }
  const reason = requiredString(input, "reason");
  if (reason.length > correctionReasonMaximumLength) {
    fail(400, "invalid_request", "reason is too long");
  }
  return {
    data: jsonRecord(input.data),
    documentKey: requiredString(input, "documentKey"),
    reason,
  };
}
interface FormMetadataInput {
  description?: string | null;
  fillMethod?: FillMethod;
  status?: FormStatus;
  title?: string;
}

async function readFormMetadataInput(
  request: Request
): Promise<FormMetadataInput> {
  const input = await readJsonRecord(request, documentActionBodyMaximumBytes);
  const keys = Object.keys(input);
  if (
    keys.length === 0 ||
    keys.some(
      (key) =>
        key !== "description" &&
        key !== "fillMethod" &&
        key !== "status" &&
        key !== "title"
    )
  ) {
    fail(
      400,
      "invalid_request",
      "Only title, description, fillMethod, and status are accepted"
    );
  }
  if (Object.hasOwn(input, "status") && keys.length !== 1) {
    fail(
      400,
      "invalid_request",
      "Status changes must not include other metadata"
    );
  }
  if (Object.hasOwn(input, "fillMethod") && keys.length !== 1) {
    fail(
      400,
      "invalid_request",
      "Fill Method changes must not include other metadata"
    );
  }
  const metadata: FormMetadataInput = {};
  if (Object.hasOwn(input, "title")) {
    if (typeof input.title !== "string" || input.title.trim().length === 0) {
      fail(400, "invalid_request", "title must be a non-empty string");
    }
    if (input.title.trim().length > 200) {
      fail(400, "invalid_request", "Title is too long");
    }
    metadata.title = input.title.trim();
  }
  if (Object.hasOwn(input, "description")) {
    if (input.description !== null && typeof input.description !== "string") {
      fail(400, "invalid_request", "description must be a string or null");
    }
    if (
      typeof input.description === "string" &&
      input.description.trim().length > 2000
    ) {
      fail(400, "invalid_request", "Description is too long");
    }
    metadata.description =
      typeof input.description === "string"
        ? input.description.trim() || null
        : null;
  }
  if (Object.hasOwn(input, "status")) {
    if (
      input.status !== FormStatus.archived &&
      input.status !== FormStatus.published
    ) {
      fail(400, "invalid_request", "status must be archived or published");
    }
    metadata.status = input.status;
  }
  if (Object.hasOwn(input, "fillMethod")) {
    if (
      input.fillMethod !== FillMethod.native &&
      input.fillMethod !== FillMethod.onlyoffice
    ) {
      fail(400, "invalid_request", "fillMethod must be native or onlyoffice");
    }
    metadata.fillMethod = input.fillMethod;
  }
  return metadata;
}

function jsonRecord(
  value: unknown,
  message = "data must be a JSON object"
): JsonRecord {
  return asRecord(value, message);
}

function operationMetadata(value: unknown): OperationMetadata {
  const metadata = asRecord(value, "Operation metadata is invalid");
  const {
    action,
    callbackEmptyFieldTags,
    cleanupObjectKeys,
    finalObjectKey,
    formId,
    nativeDocumentStaged,
    serverHeldFieldTags,
    stagedObjectKey,
    workspaceDocumentKey,
    workspaceObjectKey,
  } = metadata;
  if (
    (action !== "save-template" &&
      action !== "publish" &&
      action !== "save-draft" &&
      action !== "submit" &&
      action !== "save-correction") ||
    typeof formId !== "string" ||
    typeof stagedObjectKey !== "string" ||
    typeof finalObjectKey !== "string" ||
    (nativeDocumentStaged !== undefined &&
      typeof nativeDocumentStaged !== "boolean") ||
    (callbackEmptyFieldTags !== undefined &&
      (!Array.isArray(callbackEmptyFieldTags) ||
        callbackEmptyFieldTags.some((tag) => typeof tag !== "string"))) ||
    (serverHeldFieldTags !== undefined &&
      (!Array.isArray(serverHeldFieldTags) ||
        serverHeldFieldTags.some((tag) => typeof tag !== "string"))) ||
    (cleanupObjectKeys !== undefined &&
      (!Array.isArray(cleanupObjectKeys) ||
        cleanupObjectKeys.some((key) => typeof key !== "string"))) ||
    (workspaceDocumentKey !== undefined &&
      typeof workspaceDocumentKey !== "string") ||
    (workspaceObjectKey !== undefined && typeof workspaceObjectKey !== "string")
  ) {
    fail(500, "invalid_operation", "Operation metadata is invalid");
  }
  return metadata as unknown as OperationMetadata;
}

function databaseErrorCode(error: unknown): string | null {
  if (!error || typeof error !== "object" || !("code" in error)) {
    return null;
  }
  const code = error.code;
  return typeof code === "string" ? code : null;
}
function isSerializationConflict(error: unknown): boolean {
  if (databaseErrorCode(error) === "P2034") {
    return true;
  }
  if (databaseErrorCode(error) !== "P2010") {
    return false;
  }
  if (
    !error ||
    typeof error !== "object" ||
    !("meta" in error) ||
    !error.meta ||
    typeof error.meta !== "object" ||
    !("code" in error.meta)
  ) {
    return false;
  }
  return error.meta.code === "40001";
}

function contentHash(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}
function tokenDigest(value: string): string {
  return createHash("sha256").update(value).digest("hex");
}

function uniqueObjectKeys(
  keys: readonly (string | null | undefined)[]
): string[] {
  return [
    ...new Set(
      keys.filter(
        (key): key is string => typeof key === "string" && key.length > 0
      )
    ),
  ];
}

async function deleteObjects(
  keys: readonly (string | null | undefined)[]
): Promise<void> {
  const uniqueKeys = uniqueObjectKeys(keys);
  await Promise.all(
    uniqueKeys.map(async (key) => {
      try {
        await deleteObject(key);
      } catch (error) {
        console.error(`Could not remove object ${key}`, error);
      }
    })
  );
}

async function drainObjectCleanupIntents(
  objectKeys?: readonly string[],
  deletionResponseLookupDigest?: string,
  removeObject: (key: string) => Promise<void> = deleteObject
): Promise<void> {
  const cleanupAfter = new Date();
  const intents = await prisma.objectCleanupIntent.findMany({
    orderBy: { createdAt: "asc" },
    where: {
      cleanupAfter: { lte: cleanupAfter },
      ...(deletionResponseLookupDigest ? { deletionResponseLookupDigest } : {}),
      ...(objectKeys ? { objectKey: { in: [...objectKeys] } } : {}),
    },
  });
  for (const intent of intents) {
    if (!(await deleteObjectUnlessCanonical(intent.objectKey, removeObject))) {
      continue;
    }
    await prisma.objectCleanupIntent.deleteMany({
      where: { id: intent.id, objectKey: intent.objectKey },
    });
  }
}

async function deleteObjectUnlessCanonical(
  key: string,
  removeObject: (key: string) => Promise<void> = deleteObject
): Promise<boolean> {
  try {
    const references = await Promise.all([
      prisma.templateDraft.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.publishedTemplate.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.response.findFirst({
        select: { id: true },
        where: { draftObjectKey: key },
      }),
      prisma.submission.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.correction.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.editorLease.findFirst({
        select: { id: true },
        where: { workspaceObjectKey: key },
      }),
    ]);
    if (references.some((reference) => reference !== null)) {
      return false;
    }
    await removeObject(key);
    return !(await objectExists(key));
  } catch (error) {
    console.error(`Could not verify whether object ${key} is canonical`, error);
    return false;
  }
}

async function cleanupTerminalOperationObjects(
  operation: Operation
): Promise<void> {
  if (
    operation.status !== OperationStatus.completed &&
    operation.status !== OperationStatus.failed
  ) {
    return;
  }
  const metadata = operationMetadata(operation.metadata);
  if (operation.status === OperationStatus.completed) {
    await deleteObjects(
      metadata.cleanupObjectKeys ?? [metadata.stagedObjectKey]
    );
    return;
  }
  await deleteObjects([metadata.stagedObjectKey]);
  await deleteObjectUnlessCanonical(metadata.finalObjectKey);
  if (metadata.action === "save-correction" && metadata.workspaceDocumentKey) {
    const workspaceLease = await prisma.editorLease.findUnique({
      select: { workspaceObjectKey: true },
      where: { workspaceDocumentKey: metadata.workspaceDocumentKey },
    });
    if (!workspaceLease && metadata.workspaceObjectKey) {
      await deleteObjectUnlessCanonical(metadata.workspaceObjectKey);
    }
  }
}

function jsonValue(value: unknown): Prisma.InputJsonValue {
  return value as Prisma.InputJsonValue;
}

function editableFieldsForSnapshot(snapshot: PrefillSnapshot): JsonRecord {
  const lockedFields = jsonRecord(snapshot.lockedFields);
  return Object.fromEntries(
    Object.entries(lockedFields).map(([field, locked]) => [
      field,
      locked !== true,
    ])
  );
}

async function identityFor(request: Request): Promise<Identity | null> {
  if (!bearerTokenFor(request)) {
    return null;
  }
  const result = await auth.api.getSession({ headers: request.headers });
  if (!result?.user || !result.session) {
    return null;
  }
  const sessionUser = result.user as unknown as {
    email?: unknown;
    enabled?: unknown;
    id?: unknown;
    mustChangePassword?: unknown;
    name?: unknown;
    role?: unknown;
  };
  const liveSession = result.session as unknown as {
    expiresAt?: unknown;
    id?: unknown;
    isSso?: unknown;
  };
  const expiresAt =
    liveSession.expiresAt instanceof Date
      ? liveSession.expiresAt
      : new Date(String(liveSession.expiresAt));
  if (
    typeof sessionUser.id !== "string" ||
    typeof sessionUser.email !== "string" ||
    sessionUser.enabled !== true ||
    (liveSession.isSso === true && sessionUser.role !== "user") ||
    typeof liveSession.id !== "string" ||
    Number.isNaN(expiresAt.getTime()) ||
    expiresAt.getTime() <= Date.now()
  ) {
    return null;
  }
  return {
    email: sessionUser.email,
    expiresAt,
    id: sessionUser.id,
    isSso: liveSession.isSso === true,
    mustChangePassword: sessionUser.mustChangePassword === true,
    name:
      typeof sessionUser.name === "string" && sessionUser.name.length > 0
        ? sessionUser.name
        : sessionUser.email,
    role: sessionUser.role === "admin" ? "admin" : "user",
    sessionId: liveSession.id,
  };
}
function bearerTokenFor(request: Request): string | undefined {
  const authorization = request.headers.get("authorization");
  if (!authorization) {
    return undefined;
  }
  const match = /^Bearer\s+(?<token>.+)$/iu.exec(authorization);
  return match?.groups?.token;
}

function normalizeEmail(value: string): string {
  return value.trim().toLowerCase();
}
const pendingClaimCookieName = "__Host-folio-pending-claim";

interface PrefillHandoffCreateInput {
  email: string;
  externalReference: string;
  publicId: string;
  values: JsonRecord;
}

function prefillHandoffSecretMatches(
  expectedSecret: string,
  receivedSecret: string | null
): boolean {
  if (!receivedSecret) {
    return false;
  }
  const expectedDigest = createHash("sha256").update(expectedSecret).digest();
  const receivedDigest = createHash("sha256").update(receivedSecret).digest();
  return timingSafeEqual(expectedDigest, receivedDigest);
}

function pendingClaimCookie(value: string, maxAge: number): string {
  return `${pendingClaimCookieName}=${value}; Path=/; Max-Age=${maxAge}; HttpOnly; Secure; SameSite=Lax`;
}

function pendingClaimFor(request: Request): string | undefined {
  const cookieHeader = request.headers.get("cookie");
  if (!cookieHeader) {
    return undefined;
  }
  for (const part of cookieHeader.split(";")) {
    const separator = part.indexOf("=");
    if (separator === -1) {
      continue;
    }
    const name = part.slice(0, separator).trim();
    if (name !== pendingClaimCookieName) {
      continue;
    }
    const value = part.slice(separator + 1).trim();
    return value.length > 0 && value.length <= handoffCodeMaximumLength
      ? value
      : undefined;
  }
  return undefined;
}

function legacySsoConfigurationFromEnv(): LegacySsoConfig | null {
  const {
    LEGACY_SSO_AUTHORIZE_URL: authorizeUrl,
    LEGACY_SSO_CALLBACK_URL: callbackUrl,
    LEGACY_SSO_CLIENT_ID: clientId,
    LEGACY_SSO_CLIENT_SECRET: clientSecret,
    LEGACY_SSO_EXCHANGE_URL: exchangeUrl,
    LEGACY_SSO_PROVIDER_ID: providerId,
  } = env;
  const values = [
    authorizeUrl,
    callbackUrl,
    clientId,
    clientSecret,
    exchangeUrl,
    providerId,
  ];
  if (values.every((value) => value === undefined)) {
    return null;
  }
  if (
    typeof authorizeUrl !== "string" ||
    typeof callbackUrl !== "string" ||
    typeof clientId !== "string" ||
    typeof clientSecret !== "string" ||
    typeof exchangeUrl !== "string" ||
    typeof providerId !== "string"
  ) {
    throw new TypeError(
      "All six LEGACY_SSO settings must be configured together"
    );
  }
  return validatedLegacySsoConfiguration({
    authorizeUrl,
    callbackUrl,
    clientId,
    clientSecret,
    exchangeUrl,
    providerId,
  });
}

function validatedLegacySsoConfiguration(
  config: LegacySsoConfig
): LegacySsoConfig {
  const parseEndpoint = (value: string, label: string): URL => {
    let url: URL;
    try {
      url = new URL(value);
    } catch {
      throw new Error(`${label} must be an absolute URL`);
    }
    if (
      (url.protocol !== "https:" &&
        !(
          url.protocol === "http:" &&
          (url.hostname === "localhost" ||
            url.hostname === "127.0.0.1" ||
            url.hostname === "[::1]")
        )) ||
      url.username ||
      url.password ||
      url.hash
    ) {
      throw new Error(`${label} must use HTTPS without URL credentials`);
    }
    return url;
  };
  const authorize = parseEndpoint(
    config.authorizeUrl,
    "LEGACY_SSO_AUTHORIZE_URL"
  );
  const callback = parseEndpoint(config.callbackUrl, "LEGACY_SSO_CALLBACK_URL");
  const exchange = parseEndpoint(config.exchangeUrl, "LEGACY_SSO_EXCHANGE_URL");
  if (
    authorize.search ||
    exchange.search ||
    callback.search ||
    callback.pathname !== "/api/legacy-sso/callback" ||
    config.clientId.length === 0 ||
    config.clientId.length > 128 ||
    config.clientId.trim() !== config.clientId ||
    config.clientId.includes(":") ||
    config.clientSecret.length < 32 ||
    config.providerId.length === 0 ||
    config.providerId.length > 128 ||
    config.providerId.trim() !== config.providerId
  ) {
    throw new Error("Legacy SSO settings are invalid");
  }
  return {
    ...config,
    authorizeUrl: authorize.href,
    callbackUrl: callback.href,
    exchangeUrl: exchange.href,
  };
}

function safeLegacyFormReturnPath(
  value: unknown,
  callbackUrl: string
): string | null {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > 2048 ||
    !value.startsWith("/") ||
    value.startsWith("//") ||
    value.includes("\\") ||
    /[\u0000-\u001F\u007F]/u.test(value)
  ) {
    return null;
  }
  let decoded: string;
  try {
    decoded = decodeURIComponent(value);
  } catch {
    return null;
  }
  if (
    decoded.startsWith("//") ||
    decoded.includes("\\") ||
    /[\u0000-\u001F\u007F]/u.test(decoded)
  ) {
    return null;
  }
  const callback = new URL(callbackUrl);
  let returnUrl: URL;
  try {
    returnUrl = new URL(value, callback.origin);
  } catch {
    return null;
  }
  const isFormPath = /^\/forms\/[0-9a-f]{32}\/fill$/u.test(returnUrl.pathname);
  const isDashboardPath =
    returnUrl.pathname === "/dashboard" && !returnUrl.search;
  if (
    returnUrl.origin !== callback.origin ||
    returnUrl.hash ||
    (!isFormPath && !isDashboardPath)
  ) {
    return null;
  }
  const queryKeys = [...returnUrl.searchParams.keys()];
  if (queryKeys.some((key) => key !== "responseId") || queryKeys.length > 1) {
    return null;
  }
  const responseIds = returnUrl.searchParams.getAll("responseId");
  if (
    responseIds.length > 1 ||
    (responseIds.length === 1 && !idPattern.test(responseIds[0] ?? ""))
  ) {
    return null;
  }
  return `${returnUrl.pathname}${returnUrl.search}`;
}

function cookieValueFor(
  request: Request,
  name: string,
  maximumLength: number
): string | null {
  const cookieHeader = request.headers.get("cookie");
  if (!cookieHeader) {
    return null;
  }
  let match: string | null = null;
  for (const part of cookieHeader.split(";")) {
    const separator = part.indexOf("=");
    if (separator === -1 || part.slice(0, separator).trim() !== name) {
      continue;
    }
    const value = part.slice(separator + 1).trim();
    if (
      match !== null ||
      value.length === 0 ||
      value.length > maximumLength ||
      !/^[A-Za-z0-9_-]+$/u.test(value)
    ) {
      return null;
    }
    match = value;
  }
  return match;
}

function hostOnlySsoCookie(
  name: string,
  value: string,
  maxAge: number,
  sameSite: "Lax" | "None" = "Lax"
): string {
  return `${name}=${value}; Path=/; Max-Age=${maxAge}; HttpOnly; Secure; SameSite=${sameSite}`;
}

function legacySsoRedirectResponse(
  result: "failed" | "pending",
  returnTo?: string
): globalThis.Response {
  const location = new URL(`/login?legacySso=${result}`, corsOrigin);
  if (returnTo) {
    location.searchParams.set("returnTo", returnTo);
  }
  const headers = new Headers({
    "Cache-Control": "no-store",
    Location: location.href,
    "Referrer-Policy": "no-referrer",
  });
  headers.append("Set-Cookie", hostOnlySsoCookie(legacySsoCookieName, "", 0));
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSessionCookieName, "", 0, "None")
  );
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSwitchCookieName, "", 0, "None")
  );
  return new Response(null, { headers, status: 303 });
}
function legacySsoFailureResponse(returnTo?: string): globalThis.Response {
  return legacySsoRedirectResponse("failed", returnTo);
}
function legacySsoPendingResponse(returnTo?: string): globalThis.Response {
  return legacySsoRedirectResponse("pending", returnTo);
}

function legacySsoTransaction(value: string): LegacySsoTransaction | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return null;
  }
  const transaction = parsed as Record<string, unknown>;
  if (
    typeof transaction.callbackUrl !== "string" ||
    typeof transaction.clientId !== "string" ||
    typeof transaction.codeVerifier !== "string" ||
    typeof transaction.generation !== "string" ||
    !/^[1-9]\d*$/u.test(transaction.generation) ||
    typeof transaction.returnTo !== "string" ||
    typeof transaction.state !== "string" ||
    !/^[A-Za-z0-9._~-]{43,128}$/u.test(transaction.codeVerifier) ||
    !/^[A-Za-z0-9_-]{43}$/u.test(transaction.state) ||
    (transaction.initiatingUserId !== undefined &&
      typeof transaction.initiatingUserId !== "string") ||
    (transaction.initiatingSessionId !== undefined &&
      typeof transaction.initiatingSessionId !== "string") ||
    (typeof transaction.initiatingUserId === "string") !==
      (typeof transaction.initiatingSessionId === "string")
  ) {
    return null;
  }
  return {
    callbackUrl: transaction.callbackUrl,
    clientId: transaction.clientId,
    codeVerifier: transaction.codeVerifier,
    generation: transaction.generation,
    initiatingSessionId: transaction.initiatingSessionId as string | undefined,
    initiatingUserId: transaction.initiatingUserId as string | undefined,
    returnTo: transaction.returnTo,
    state: transaction.state,
  };
}

async function readLegacyIdentityResponse(
  response: globalThis.Response
): Promise<LegacyIdentity | null> {
  if (
    !/^application\/json(?:\s*;|$)/iu.test(
      response.headers.get("content-type") ?? ""
    )
  ) {
    return null;
  }
  const contentLength = Number(response.headers.get("content-length") ?? "");
  if (
    Number.isFinite(contentLength) &&
    contentLength > legacySsoExchangeBodyMaximumBytes
  ) {
    return null;
  }
  const reader = response.body?.getReader();
  if (!reader) {
    return null;
  }
  const chunks: Uint8Array[] = [];
  let byteLength = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    byteLength += value.byteLength;
    if (byteLength > legacySsoExchangeBodyMaximumBytes) {
      await reader.cancel();
      return null;
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(byteLength);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  let payload: unknown;
  try {
    payload = JSON.parse(new TextDecoder().decode(bytes));
  } catch {
    return null;
  }
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    return null;
  }
  const claims = payload as Record<string, unknown>;
  if (
    typeof claims.sub !== "string" ||
    claims.sub.length === 0 ||
    claims.sub.length > 255 ||
    claims.sub.trim() !== claims.sub ||
    /[\u0000-\u001F\u007F]/u.test(claims.sub) ||
    typeof claims.email !== "string" ||
    claims.email.length > accountEmailMaximumLength ||
    !accountEmailPattern.test(claims.email) ||
    claims.email_verified !== true
  ) {
    return null;
  }
  return { email: normalizeEmail(claims.email), subject: claims.sub };
}

async function exchangeLegacyCode(
  config: LegacySsoConfig,
  code: string,
  codeVerifier: string
): Promise<LegacyIdentity | null> {
  try {
    const response = await fetch(config.exchangeUrl, {
      body: new URLSearchParams({
        client_id: config.clientId,
        code,
        code_verifier: codeVerifier,
        grant_type: "authorization_code",
        redirect_uri: config.callbackUrl,
      }),
      headers: {
        Accept: "application/json",
        Authorization: `Basic ${Buffer.from(
          `${config.clientId}:${config.clientSecret}`
        ).toString("base64")}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      method: "POST",
      redirect: "error",
      signal: AbortSignal.timeout(5000),
    });
    if (!response.ok) {
      return null;
    }
    return await readLegacyIdentityResponse(response);
  } catch {
    return null;
  }
}
async function nextLegacySsoGeneration(
  tx: Prisma.TransactionClient,
  providerId: string
): Promise<bigint> {
  await tx.$executeRaw`
    SELECT pg_advisory_xact_lock(
      hashtext('legacy-sso-generation'),
      hashtext(${providerId})
    )
  `;
  const [row] = await tx.$queryRaw<{ generation: bigint }[]>(
    Prisma.sql`SELECT nextval('"legacy_sso_generation_seq"') AS generation`
  );
  if (!row) {
    throw new Error("Database did not return SSO generation");
  }
  return row.generation;
}

async function linkApprovedLegacyAccount(
  config: LegacySsoConfig,
  identity: LegacyIdentity,
  requestId: string,
  userId: string,
  transactionGeneration: bigint
): Promise<LegacySsoAccountUser | null> {
  return prisma.$transaction(async (tx) => {
    const [lockedUser] = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${userId} FOR UPDATE`
    );
    if (!lockedUser) {
      return null;
    }
    const [lockedRequest] = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`SELECT "id" FROM "legacy_account_link_requests" WHERE "id" = ${requestId}::uuid FOR UPDATE`
    );
    if (!lockedRequest) {
      return null;
    }
    const linkRequest = await tx.legacyAccountLinkRequest.findUnique({
      where: { id: requestId },
    });
    if (
      !linkRequest ||
      (linkRequest.status !== LegacyAccountLinkStatus.approved &&
        linkRequest.status !== LegacyAccountLinkStatus.linked) ||
      !linkRequest.reviewedGeneration ||
      transactionGeneration <= linkRequest.reviewedGeneration ||
      linkRequest.providerId !== config.providerId ||
      linkRequest.subject !== identity.subject ||
      linkRequest.email !== identity.email
    ) {
      return null;
    }
    const user = await tx.user.findUnique({
      select: { email: true, enabled: true, id: true, role: true },
      where: { id: userId },
    });
    if (
      !user ||
      !user.enabled ||
      user.role !== "user" ||
      user.email !== linkRequest.email
    ) {
      return null;
    }
    const account = await tx.account.findUnique({
      select: { userId: true },
      where: {
        providerId_accountId: {
          accountId: identity.subject,
          providerId: config.providerId,
        },
      },
    });
    if (account && account.userId !== user.id) {
      return null;
    }
    if (!account) {
      await tx.account.create({
        data: {
          accountId: identity.subject,
          id: crypto.randomUUID(),
          issuer: config.providerId,
          providerId: config.providerId,
          userId: user.id,
        },
      });
    }
    if (linkRequest.status === LegacyAccountLinkStatus.approved) {
      await tx.legacyAccountLinkRequest.update({
        data: { status: LegacyAccountLinkStatus.linked },
        where: { id: requestId },
      });
    }
    return user;
  });
}

async function resolveLegacySsoLinkRequest(
  config: LegacySsoConfig,
  identity: LegacyIdentity,
  transactionGeneration: bigint
): Promise<LegacySsoAccountResolution | null> {
  const linkRequest = await prisma.legacyAccountLinkRequest.findUnique({
    where: {
      providerId_subject: {
        providerId: config.providerId,
        subject: identity.subject,
      },
    },
  });
  if (!linkRequest) {
    return null;
  }
  if (linkRequest.email !== identity.email) {
    return { kind: "failed" };
  }
  if (linkRequest.status === LegacyAccountLinkStatus.pending) {
    return { kind: "pending" };
  }
  if (
    linkRequest.status !== LegacyAccountLinkStatus.approved &&
    linkRequest.status !== LegacyAccountLinkStatus.linked
  ) {
    return { kind: "failed" };
  }
  const user = await linkApprovedLegacyAccount(
    config,
    identity,
    linkRequest.id,
    linkRequest.userId,
    transactionGeneration
  );
  return user ? { kind: "linked", user } : { kind: "failed" };
}

async function resolveLegacySsoAccount(
  config: LegacySsoConfig,
  identity: LegacyIdentity,
  transactionGeneration: bigint
): Promise<LegacySsoAccountResolution | null> {
  const existingRequest = await resolveLegacySsoLinkRequest(
    config,
    identity,
    transactionGeneration
  );
  if (existingRequest) {
    return existingRequest;
  }
  const linkedAccountLookup = {
    include: {
      user: {
        select: { enabled: true, id: true, role: true },
      },
    },
    where: {
      providerId_accountId: {
        accountId: identity.subject,
        providerId: config.providerId,
      },
    },
  } as const;
  const linkedAccount = await prisma.account.findUnique(linkedAccountLookup);
  if (linkedAccount) {
    const latestRequest = await resolveLegacySsoLinkRequest(
      config,
      identity,
      transactionGeneration
    );
    return latestRequest ?? { kind: "linked", user: linkedAccount.user };
  }
  try {
    const user = await prisma.user.create({
      data: {
        accounts: {
          create: {
            accountId: identity.subject,
            id: crypto.randomUUID(),
            issuer: config.providerId,
            providerId: config.providerId,
          },
        },
        email: identity.email,
        emailVerified: true,
        enabled: true,
        id: crypto.randomUUID(),
        mustChangePassword: false,
        name: identity.email,
        role: "user",
      },
      select: { enabled: true, id: true, role: true },
    });
    return { kind: "linked", user };
  } catch (error) {
    if (databaseErrorCode(error) !== "P2002") {
      throw error;
    }
    const racedAccount = await prisma.account.findUnique(linkedAccountLookup);
    if (racedAccount) {
      const latestRequest = await resolveLegacySsoLinkRequest(
        config,
        identity,
        transactionGeneration
      );
      return latestRequest ?? { kind: "linked", user: racedAccount.user };
    }
    const existingUser = await prisma.user.findUnique({
      select: { id: true },
      where: { email: identity.email },
    });
    if (!existingUser) {
      return null;
    }
    try {
      await prisma.legacyAccountLinkRequest.create({
        data: {
          email: identity.email,
          providerId: config.providerId,
          subject: identity.subject,
          userId: existingUser.id,
        },
      });
      return { kind: "pending" };
    } catch (requestError) {
      if (
        databaseErrorCode(requestError) !== "P2002" &&
        databaseErrorCode(requestError) !== "P2003"
      ) {
        throw requestError;
      }
      return resolveLegacySsoLinkRequest(
        config,
        identity,
        transactionGeneration
      );
    }
  }
}

async function createSsoSessionTransfer(
  userId: string,
  now: Date
): Promise<string> {
  const authContext = await auth.$context;
  const session = await authContext.internalAdapter.createSession(
    userId,
    false,
    { isSso: true },
    true
  );
  const transferId = randomBytes(32).toString("base64url");
  try {
    await prisma.verification.create({
      data: {
        expiresAt: new Date(
          now.getTime() + legacySsoSessionTransferLifetimeSeconds * 1000
        ),
        id: transferId,
        identifier: legacySsoSessionIdentifier,
        value: session.token,
      },
    });
  } catch (error) {
    await authContext.internalAdapter.deleteSession(session.token);
    throw error;
  }
  return transferId;
}
function legacySsoSwitchChallenge(
  value: string
): LegacySsoSwitchChallenge | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return null;
  }
  const challenge = parsed as Record<string, unknown>;
  if (
    typeof challenge.initiatingSessionId !== "string" ||
    typeof challenge.initiatingUserId !== "string" ||
    typeof challenge.returnTo !== "string" ||
    typeof challenge.targetUserId !== "string"
  ) {
    return null;
  }
  return {
    initiatingSessionId: challenge.initiatingSessionId,
    initiatingUserId: challenge.initiatingUserId,
    returnTo: challenge.returnTo,
    targetUserId: challenge.targetUserId,
  };
}

function legacySsoSwitchCookie(value: string, maxAge: number): string {
  return hostOnlySsoCookie(legacySsoSwitchCookieName, value, maxAge, "None");
}

function legacySsoSwitchFailure(): globalThis.Response {
  return Response.json(
    { error: "unauthorized", message: "The account switch is unavailable" },
    { headers: { "Cache-Control": "no-store" }, status: 401 }
  );
}

async function startLegacySso(
  request: Request,
  config: LegacySsoConfig,
  clock: () => Date,
  identity: Identity | null
): Promise<globalThis.Response> {
  if (identity && identity.role !== "user") {
    fail(403, "legacy_sso_user_required", "Only Users can switch accounts");
  }
  if (new URL(request.url).search) {
    fail(
      400,
      "invalid_return_path",
      "Return path must be sent in the JSON body"
    );
  }
  const input = await readJsonRecord(request, 8 * 1024);
  if (
    Object.keys(input).some((key) => key !== "returnTo") ||
    (input.returnTo !== undefined && typeof input.returnTo !== "string")
  ) {
    fail(400, "invalid_return_path", "Only a Form return path is accepted");
  }
  const returnTo =
    input.returnTo === undefined
      ? "/dashboard"
      : safeLegacyFormReturnPath(input.returnTo, config.callbackUrl);
  if (!returnTo) {
    fail(400, "invalid_return_path", "The Form return path is unsafe");
  }
  const state = randomBytes(32).toString("base64url");
  const codeVerifier = randomBytes(32).toString("base64url");
  const codeChallenge = createHash("sha256")
    .update(codeVerifier)
    .digest("base64url");
  const transactionId = randomBytes(32).toString("base64url");
  const now = clock();
  const expiresAt = new Date(
    now.getTime() + legacySsoTransactionLifetimeSeconds * 1000
  );
  await prisma.verification.deleteMany({
    where: {
      expiresAt: { lte: now },
      identifier: {
        in: [
          legacySsoTransactionIdentifier,
          legacySsoSessionIdentifier,
          legacySsoSwitchIdentifier,
        ],
      },
    },
  });
  await prisma.$transaction(async (tx) => {
    const generation = await nextLegacySsoGeneration(tx, config.providerId);
    await tx.verification.create({
      data: {
        expiresAt,
        id: transactionId,
        identifier: legacySsoTransactionIdentifier,
        value: JSON.stringify({
          callbackUrl: config.callbackUrl,
          clientId: config.clientId,
          codeVerifier,
          generation: generation.toString(),
          initiatingSessionId: identity?.sessionId,
          initiatingUserId: identity?.id,
          returnTo,
          state,
        } satisfies LegacySsoTransaction),
      },
    });
  });
  const authorizeUrl = new URL(config.authorizeUrl);
  authorizeUrl.searchParams.set("client_id", config.clientId);
  authorizeUrl.searchParams.set("redirect_uri", config.callbackUrl);
  authorizeUrl.searchParams.set("state", state);
  authorizeUrl.searchParams.set("code_challenge", codeChallenge);
  authorizeUrl.searchParams.set("code_challenge_method", "S256");
  const headers = new Headers({
    "Cache-Control": "no-store",
    "Referrer-Policy": "no-referrer",
  });
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(
      legacySsoCookieName,
      transactionId,
      legacySsoTransactionLifetimeSeconds,
      "None"
    )
  );
  headers.append("Set-Cookie", legacySsoSwitchCookie("", 0));
  return Response.json({ authorizationUrl: authorizeUrl.href }, { headers });
}

async function completeLegacySso(
  request: Request,
  config: LegacySsoConfig,
  clock: () => Date
): Promise<globalThis.Response> {
  const requestUrl = new URL(request.url);
  const forwardedProtocol = request.headers.get("x-forwarded-proto")?.trim();
  const publicOrigin = forwardedProtocol
    ? `${forwardedProtocol}://${requestUrl.host}`
    : requestUrl.origin;
  const callbackUrl = new URL(config.callbackUrl);
  const queryKeys = [...requestUrl.searchParams.keys()];
  const codeValues = requestUrl.searchParams.getAll("code");
  const stateValues = requestUrl.searchParams.getAll("state");
  if (
    `${publicOrigin}${requestUrl.pathname}` !== callbackUrl.href ||
    queryKeys.length !== 2 ||
    !queryKeys.includes("code") ||
    !queryKeys.includes("state") ||
    codeValues.length !== 1 ||
    stateValues.length !== 1
  ) {
    return legacySsoFailureResponse();
  }
  const code = codeValues[0];
  const state = stateValues[0];
  const transactionId = cookieValueFor(request, legacySsoCookieName, 43);
  if (
    !code ||
    code.length > handoffCodeMaximumLength ||
    !state ||
    !/^[A-Za-z0-9_-]{43}$/u.test(state) ||
    !transactionId
  ) {
    return legacySsoFailureResponse();
  }
  const now = clock();
  const storedTransaction = await prisma.verification.findUnique({
    where: { id: transactionId },
  });
  const transaction = storedTransaction
    ? legacySsoTransaction(storedTransaction.value)
    : null;
  if (
    !storedTransaction ||
    storedTransaction.identifier !== legacySsoTransactionIdentifier ||
    storedTransaction.expiresAt.getTime() <= now.getTime() ||
    !transaction ||
    transaction.state !== state ||
    transaction.clientId !== config.clientId ||
    transaction.callbackUrl !== config.callbackUrl ||
    !safeLegacyFormReturnPath(transaction.returnTo, config.callbackUrl)
  ) {
    return legacySsoFailureResponse();
  }
  const sessionReturnTo =
    transaction.initiatingUserId && transaction.initiatingSessionId
      ? transaction.returnTo
      : undefined;
  const consumed = await prisma.verification.deleteMany({
    where: {
      expiresAt: { gt: now },
      id: transactionId,
      identifier: legacySsoTransactionIdentifier,
    },
  });
  if (consumed.count !== 1) {
    return legacySsoFailureResponse();
  }
  const identity = await exchangeLegacyCode(
    config,
    code,
    transaction.codeVerifier
  );
  if (!identity) {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  const accountResolution = await resolveLegacySsoAccount(
    config,
    identity,
    BigInt(transaction.generation)
  );
  if (!accountResolution) {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  if (accountResolution.kind === "pending") {
    return legacySsoPendingResponse(sessionReturnTo);
  }
  if (accountResolution.kind === "failed") {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  const user = accountResolution.user;
  if (!user.enabled || user.role !== "user") {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  if (transaction.initiatingUserId && transaction.initiatingSessionId) {
    const initiatingSession = await prisma.session.findUnique({
      select: { expiresAt: true, userId: true },
      where: { id: transaction.initiatingSessionId },
    });
    if (
      !initiatingSession ||
      initiatingSession.userId !== transaction.initiatingUserId
    ) {
      return legacySsoFailureResponse(sessionReturnTo);
    }
    if (initiatingSession.expiresAt.getTime() > now.getTime()) {
      const challengeId = randomBytes(32).toString("base64url");
      await prisma.verification.create({
        data: {
          expiresAt: new Date(
            now.getTime() + legacySsoSwitchLifetimeSeconds * 1000
          ),
          id: challengeId,
          identifier: legacySsoSwitchIdentifier,
          value: JSON.stringify({
            initiatingSessionId: transaction.initiatingSessionId,
            initiatingUserId: transaction.initiatingUserId,
            returnTo: transaction.returnTo,
            targetUserId: user.id,
          } satisfies LegacySsoSwitchChallenge),
        },
      });
      const headers = new Headers({
        "Cache-Control": "no-store",
        Location: new URL("/legacy-sso/confirm", corsOrigin).href,
        "Referrer-Policy": "no-referrer",
      });
      headers.append(
        "Set-Cookie",
        hostOnlySsoCookie(legacySsoCookieName, "", 0)
      );
      headers.append(
        "Set-Cookie",
        legacySsoSwitchCookie(challengeId, legacySsoSwitchLifetimeSeconds)
      );
      return new Response(null, { headers, status: 303 });
    }
    if (user.id !== transaction.initiatingUserId) {
      return legacySsoFailureResponse(sessionReturnTo);
    }
  }
  const transferId = await createSsoSessionTransfer(user.id, now);
  const headers = new Headers({
    "Cache-Control": "no-store",
    Location: new URL(transaction.returnTo, corsOrigin).href,
    "Referrer-Policy": "no-referrer",
  });
  headers.append("Set-Cookie", hostOnlySsoCookie(legacySsoCookieName, "", 0));
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(
      legacySsoSessionCookieName,
      transferId,
      legacySsoSessionTransferLifetimeSeconds,
      "None"
    )
  );
  return new Response(null, { headers, status: 303 });
}

function legacySsoSessionUnavailableResponse(): globalThis.Response {
  const headers = new Headers({ "Cache-Control": "no-store" });
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSessionCookieName, "", 0, "None")
  );
  return Response.json(
    {
      error: "unauthorized",
      message: "The legacy SSO session handoff is unavailable",
    },
    { headers, status: 401 }
  );
}

async function claimLegacySsoSession(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null
): Promise<globalThis.Response> {
  const requestOrigin = request.headers.get("origin");
  const allowedOrigins = [
    corsOrigin,
    originOf(env.BETTER_AUTH_URL),
    config ? originOf(config.callbackUrl) : null,
  ];
  if (!config || !requestOrigin || !allowedOrigins.includes(requestOrigin)) {
    return Response.json(
      {
        error: "forbidden",
        message: "The legacy SSO session handoff origin is invalid",
      },
      { headers: { "Cache-Control": "no-store" }, status: 403 }
    );
  }
  const transferId = cookieValueFor(request, legacySsoSessionCookieName, 43);
  if (!transferId) {
    return legacySsoSessionUnavailableResponse();
  }
  const now = clock();
  const transfer = await prisma.verification.findUnique({
    where: { id: transferId },
  });
  if (
    !transfer ||
    transfer.identifier !== legacySsoSessionIdentifier ||
    transfer.expiresAt.getTime() <= now.getTime()
  ) {
    return legacySsoSessionUnavailableResponse();
  }
  const consumed = await prisma.verification.deleteMany({
    where: {
      expiresAt: { gt: now },
      id: transferId,
      identifier: legacySsoSessionIdentifier,
    },
  });
  if (consumed.count !== 1) {
    return legacySsoSessionUnavailableResponse();
  }
  const headers = new Headers({ "Cache-Control": "no-store" });
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSessionCookieName, "", 0, "None")
  );
  return Response.json({ token: transfer.value }, { headers });
}
interface LegacySsoSwitchContext {
  challenge: LegacySsoSwitchChallenge;
  challengeId: string;
  confirmationFingerprint: string | null;
  identity: Identity;
  target: { email: string; id: string; name: string } | null;
}

function legacySsoConfirmationFingerprint(
  identity: Identity,
  challenge: LegacySsoSwitchChallenge,
  target: { email: string; id: string; name: string }
): string {
  return createHash("sha256")
    .update(
      JSON.stringify([
        identity.id,
        identity.sessionId,
        identity.email,
        identity.name,
        target.id,
        target.email,
        target.name,
        challenge.returnTo,
      ])
    )
    .digest("base64url");
}

async function legacySsoSwitchContext(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null
): Promise<LegacySsoSwitchContext | null> {
  const requestOrigin = request.headers.get("origin");
  const requestOriginAllowed =
    requestOrigin === null
      ? request.method === "GET"
      : config !== null &&
        [
          corsOrigin,
          originOf(env.BETTER_AUTH_URL),
          originOf(config.callbackUrl),
        ].includes(requestOrigin);
  const identity = await identityFor(request);
  const challengeId = cookieValueFor(request, legacySsoSwitchCookieName, 43);
  if (
    !config ||
    !requestOriginAllowed ||
    !identity ||
    identity.role !== "user" ||
    !challengeId
  ) {
    return null;
  }
  const now = clock();
  const stored = await prisma.verification.findUnique({
    where: { id: challengeId },
  });
  const challenge =
    stored?.identifier === legacySsoSwitchIdentifier &&
    stored.expiresAt.getTime() > now.getTime()
      ? legacySsoSwitchChallenge(stored.value)
      : null;
  if (
    !challenge ||
    challenge.initiatingUserId !== identity.id ||
    challenge.initiatingSessionId !== identity.sessionId ||
    !safeLegacyFormReturnPath(challenge.returnTo, config.callbackUrl)
  ) {
    return null;
  }
  const target = await prisma.user.findUnique({
    select: { email: true, enabled: true, id: true, name: true, role: true },
    where: { id: challenge.targetUserId },
  });
  const switchTarget =
    target?.enabled && target.role === "user"
      ? { email: target.email, id: target.id, name: target.name }
      : null;
  return {
    challenge,
    challengeId,
    confirmationFingerprint: switchTarget
      ? legacySsoConfirmationFingerprint(identity, challenge, switchTarget)
      : null,
    identity,
    target: switchTarget,
  };
}

async function readLegacySsoSwitch(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null
): Promise<globalThis.Response> {
  const context = await legacySsoSwitchContext(request, clock, config);
  if (!context?.target || !context.confirmationFingerprint) {
    return legacySsoSwitchFailure();
  }
  const currentIdentity = context.identity;
  return Response.json(
    {
      confirmationFingerprint: context.confirmationFingerprint,
      current: { email: currentIdentity.email, name: currentIdentity.name },
      legacy: { email: context.target.email, name: context.target.name },
      returnTo: context.challenge.returnTo,
      sameUser: context.target.id === currentIdentity.id,
    },
    { headers: { "Cache-Control": "no-store" } }
  );
}

async function consumeLegacySsoSwitch(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null,
  confirm: boolean
): Promise<globalThis.Response> {
  const context = await legacySsoSwitchContext(request, clock, config);
  if (!context?.target || !context.confirmationFingerprint) {
    return legacySsoSwitchFailure();
  }
  const input = await readJsonRecord(request, 1024);
  if (
    Object.keys(input).length !== 1 ||
    typeof input.confirmationFingerprint !== "string" ||
    input.confirmationFingerprint !== context.confirmationFingerprint
  ) {
    return legacySsoSwitchFailure();
  }
  const now = clock();
  const consumed = await prisma.verification.deleteMany({
    where: {
      expiresAt: { gt: now },
      id: context.challengeId,
      identifier: legacySsoSwitchIdentifier,
    },
  });
  if (consumed.count !== 1) {
    return legacySsoSwitchFailure();
  }
  const headers = new Headers({ "Cache-Control": "no-store" });
  headers.append("Set-Cookie", legacySsoSwitchCookie("", 0));
  if (!confirm) {
    return Response.json({ cancelled: true }, { headers });
  }
  const transferId = await createSsoSessionTransfer(context.target.id, now);
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(
      legacySsoSessionCookieName,
      transferId,
      legacySsoSessionTransferLifetimeSeconds,
      "None"
    )
  );
  return Response.json({ returnTo: context.challenge.returnTo }, { headers });
}

function handoffUnavailable(): never {
  fail(409, "handoff_unavailable", "The prefill handoff is unavailable");
}

function prefillRequired(): never {
  fail(409, "prefill_required", "A prefill handoff is required");
}

function handoffEmail(value: unknown): string {
  if (typeof value !== "string") {
    fail(400, "invalid_request", "email must be a valid email address");
  }
  const email = normalizeEmail(value);
  if (
    email.length > accountEmailMaximumLength ||
    !accountEmailPattern.test(email)
  ) {
    fail(400, "invalid_request", "email must be a valid email address");
  }
  return email;
}

function handoffExternalReference(value: unknown): string {
  if (
    typeof value !== "string" ||
    value.trim().length === 0 ||
    value.trim().length > handoffExternalReferenceMaximumLength
  ) {
    fail(
      400,
      "invalid_request",
      "externalReference must be a non-empty bounded string"
    );
  }
  return value.trim();
}

function handoffCode(value: unknown): string {
  if (
    typeof value !== "string" ||
    value.trim().length === 0 ||
    value.trim().length > handoffCodeMaximumLength
  ) {
    fail(400, "invalid_request", "code is required");
  }
  return value.trim();
}

function handoffCreateInput(input: JsonRecord): PrefillHandoffCreateInput {
  const keys = Object.keys(input);
  const expectedKeys = new Set([
    "email",
    "externalReference",
    "publicId",
    "values",
  ]);
  if (
    keys.length !== expectedKeys.size ||
    keys.some((key) => !expectedKeys.has(key))
  ) {
    fail(
      400,
      "invalid_request",
      "publicId, email, externalReference, and values are required"
    );
  }
  if (typeof input.publicId !== "string") {
    fail(404, "not_found", "Form was not found");
  }
  const values = asRecord(input.values, "values must be a JSON object");
  return {
    email: handoffEmail(input.email),
    externalReference: handoffExternalReference(input.externalReference),
    publicId: input.publicId,
    values,
  };
}
function handoffStatusInput(input: JsonRecord): {
  externalReference: string;
} {
  if (
    Object.keys(input).length !== 1 ||
    Object.keys(input)[0] !== "externalReference"
  ) {
    fail(400, "invalid_request", "externalReference is required");
  }
  return {
    externalReference: handoffExternalReference(input.externalReference),
  };
}

function requireTopLevelNavigation(request: Request): void {
  const fetchMode = request.headers.get("sec-fetch-mode");
  const fetchDestination = request.headers.get("sec-fetch-dest");
  if (
    (fetchMode !== null && fetchMode !== "navigate") ||
    (fetchDestination !== null && fetchDestination !== "document")
  ) {
    fail(400, "invalid_request", "The handoff must be a top-level navigation");
  }
}

async function readPrefillHandoffCode(request: Request): Promise<string> {
  const contentType =
    request.headers
      .get("content-type")
      ?.split(";", 1)[0]
      ?.trim()
      .toLowerCase() ?? "";
  if (
    contentType !== "application/x-www-form-urlencoded" &&
    contentType !== "multipart/form-data"
  ) {
    fail(415, "invalid_request", "The handoff code body format is unsupported");
  }
  const bytes = await readRequestBytes(
    request,
    prefillHandoffBodyMaximumBytes,
    "The handoff code is required"
  );
  if (contentType === "application/x-www-form-urlencoded") {
    const entries = new Map<string, string>();
    for (const [key, value] of new URLSearchParams(
      new TextDecoder().decode(bytes)
    )) {
      if (key !== "code" || entries.has(key)) {
        fail(400, "invalid_request", "Only code is accepted");
      }
      entries.set(key, value);
    }
    return handoffCode(entries.get("code"));
  }
  if (contentType === "multipart/form-data") {
    const multipartRequest = new Request(request.url, {
      body: bytes,
      headers: {
        "content-type": request.headers.get("content-type") as string,
      },
      method: "POST",
    });
    let formData: Awaited<ReturnType<typeof multipartRequest.formData>>;
    try {
      formData = await multipartRequest.formData();
    } catch {
      fail(400, "invalid_request", "Multipart form data is invalid");
    }
    let code: string | undefined;
    for (const [key, value] of formData.entries()) {
      if (key !== "code" || typeof value !== "string" || code !== undefined) {
        fail(400, "invalid_request", "Only code is accepted");
      }
      code = value;
    }
    return handoffCode(code);
  }
  fail(415, "invalid_request", "The handoff code body format is unsupported");
}

function externalPointerSegments(pointer: string): string[] | null {
  if (!pointer.startsWith("/")) {
    return null;
  }
  if (pointer === "/") {
    return [""];
  }
  return pointer
    .slice(1)
    .split("/")
    .map((segment) => {
      let decoded = "";
      for (let index = 0; index < segment.length; index += 1) {
        const character = segment[index];
        if (character !== "~") {
          decoded += character;
          continue;
        }
        const escape = segment[index + 1];
        if (escape === "0") {
          decoded += "~";
        } else if (escape === "1") {
          decoded += "/";
        } else {
          return "";
        }
        index += 1;
      }
      return decoded;
    });
}

function externalValueAtPointer(values: JsonRecord, pointer: string): unknown {
  if (Object.hasOwn(values, pointer)) {
    return values[pointer];
  }
  const segments = externalPointerSegments(pointer);
  if (!segments) {
    return undefined;
  }
  let current: unknown = values;
  for (const segment of segments) {
    if (
      !current ||
      typeof current !== "object" ||
      Array.isArray(current) ||
      !Object.hasOwn(current, segment)
    ) {
      return undefined;
    }
    current = (current as JsonRecord)[segment];
  }
  return current;
}

function externalValueMatchesSchema(pointer: string, value: unknown): boolean {
  const schemaItem = externalSchemaItems.find(
    (candidate) => candidate.pointer === pointer
  );
  return (
    schemaItem !== undefined &&
    externalSchemaLeafType(value) === schemaItem.type
  );
}
function filteredPrefillValues(
  values: JsonRecord,
  fields: readonly { pointer: string; tag: string }[]
): JsonRecord {
  const filtered: JsonRecord = {};
  for (const field of fields) {
    const value = externalValueAtPointer(values, field.pointer);
    if (value === undefined) {
      continue;
    }
    if (!externalValueMatchesSchema(field.pointer, value)) {
      handoffUnavailable();
    }
    filtered[field.tag] = value;
  }
  return filtered;
}

function loginDigest(kind: "email" | "ip", value: string): string {
  return createHmac("sha256", env.BETTER_AUTH_SECRET)
    .update(`${kind}:${value}`)
    .digest("hex");
}

function loginFailureKey(email: string, sourceIp: string) {
  return {
    emailDigest: loginDigest("email", email),
    ipDigest: loginDigest("ip", sourceIp.trim().toLowerCase() || "unknown"),
  };
}

async function reserveLoginAttempt(
  emailDigest: string,
  ipDigest: string
): Promise<number> {
  const now = new Date();
  await prisma.loginFailure.deleteMany({
    where: { expiresAt: { lte: now } },
  });
  const expiresAt = new Date(now.getTime() + loginFailureWindowMs);
  const [failure] = await prisma.$queryRaw<{ attempts: number }[]>`
    INSERT INTO "login_failures" AS "login_failure" (
      "id",
      "email_digest",
      "ip_digest",
      "attempts",
      "window_started_at",
      "expires_at",
      "updated_at"
    )
    VALUES (
      ${crypto.randomUUID()}::uuid,
      ${emailDigest},
      ${ipDigest},
      1,
      ${now},
      ${expiresAt},
      ${now}
    )
    ON CONFLICT ("email_digest", "ip_digest") DO UPDATE SET
      "attempts" = CASE
        WHEN "login_failure"."expires_at" <= ${now} THEN 1
        ELSE "login_failure"."attempts" + 1
      END,
      "window_started_at" = CASE
        WHEN "login_failure"."expires_at" <= ${now} THEN ${now}
        ELSE "login_failure"."window_started_at"
      END,
      "expires_at" = CASE
        WHEN "login_failure"."expires_at" <= ${now} THEN ${expiresAt}
        ELSE "login_failure"."expires_at"
      END,
      "updated_at" = ${now}
    RETURNING "attempts"
  `;
  return failure?.attempts ?? loginFailureLimit + 1;
}

async function clearLoginFailures(
  emailDigest: string,
  ipDigest: string
): Promise<void> {
  await prisma.loginFailure.deleteMany({ where: { emailDigest, ipDigest } });
}

function loginError(
  status: 401 | 429,
  error: "invalid_credentials" | "login_throttled"
): globalThis.Response {
  return Response.json(
    { error, message: "Email or password is invalid" },
    {
      headers: status === 429 ? { "Retry-After": "900" } : undefined,
      status,
    }
  );
}

async function withAiAuthoringSessionsEnding<T>(
  aiAuthoring: AiAuthoringSessions,
  ownerSessionIds: readonly string[],
  revokeOwnerSessions: () => Promise<T>
): Promise<T> {
  const uniqueSessionIds = [...new Set(ownerSessionIds)];
  if (uniqueSessionIds.length === 0) {
    return await revokeOwnerSessions();
  }
  let revocation: Promise<T> | undefined;
  const revokeOnce = (): Promise<T> =>
    (revocation ??= Promise.resolve().then(revokeOwnerSessions));
  const endings = await Promise.allSettled(
    uniqueSessionIds.map((sessionId) =>
      aiAuthoring.endForSession(sessionId, revokeOnce)
    )
  );
  const errors = [
    ...new Set(
      endings.flatMap((ending) =>
        ending.status === "rejected" ? [ending.reason] : []
      )
    ),
  ];
  if (errors.length === 1) {
    throw errors[0];
  }
  if (errors.length > 1) {
    throw new AggregateError(errors, "AI authoring session cleanup failed");
  }
  return await (revocation ?? revokeOnce());
}

async function endAiAuthoringSessions(
  aiAuthoring: AiAuthoringSessions,
  ownerSessionIds: readonly string[]
): Promise<void> {
  await withAiAuthoringSessionsEnding(
    aiAuthoring,
    ownerSessionIds,
    async () => {}
  );
}
async function handleEmailSignIn(
  request: Request,
  body: unknown,
  sourceIp: string,
  aiAuthoring: AiAuthoringSessions
): Promise<globalThis.Response> {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    return loginError(401, "invalid_credentials");
  }
  const input = body as JsonRecord;
  if (typeof input.email !== "string" || typeof input.password !== "string") {
    return loginError(401, "invalid_credentials");
  }
  const email = normalizeEmail(input.email);
  const { emailDigest, ipDigest } = loginFailureKey(email, sourceIp);
  if ((await reserveLoginAttempt(emailDigest, ipDigest)) > loginFailureLimit) {
    return loginError(429, "login_throttled");
  }
  if (
    input.password.length < passwordMinimumLength ||
    input.password.length > passwordMaximumLength
  ) {
    return loginError(401, "invalid_credentials");
  }

  const headers = new Headers(request.headers);
  headers.delete("content-length");
  const authResponse = await auth.handler(
    new Request(request.url, {
      body: JSON.stringify({ email, password: input.password }),
      headers,
      method: "POST",
    })
  );
  if (!authResponse.ok) {
    return loginError(401, "invalid_credentials");
  }

  const user = await prisma.user.findUnique({ where: { email } });
  if (!user?.enabled) {
    if (user) {
      const sessions = await prisma.session.findMany({
        select: { id: true },
        where: { userId: user.id },
      });
      await withAiAuthoringSessionsEnding(
        aiAuthoring,
        sessions.map(({ id }) => id),
        () => prisma.session.deleteMany({ where: { userId: user.id } })
      );
    }
    return loginError(401, "invalid_credentials");
  }

  await clearLoginFailures(emailDigest, ipDigest);
  return authResponse;
}

async function requireIdentity(request: Request): Promise<Identity> {
  const identity = await identityFor(request);
  if (!identity) {
    fail(401, "unauthorized", "Authentication is required");
  }
  if (identity.mustChangePassword && !identity.isSso) {
    fail(403, "password_change_required", "Password replacement is required");
  }
  return identity;
}
async function editorAuthorization(
  request: Request
): Promise<EditorAuthorization> {
  const token = request.headers.get("x-editor-capability")?.trim();
  if (!token) {
    return { actor: await requireIdentity(request), capability: null };
  }
  const capability = verifyEditorCapability(token);
  if (!capability) {
    fail(
      401,
      "invalid_editor_capability",
      "Editor capability is invalid or expired"
    );
  }
  const user = await prisma.user.findUnique({
    select: {
      email: true,
      enabled: true,
      id: true,
      mustChangePassword: true,
      name: true,
      role: true,
    },
    where: { id: capability.actorId },
  });
  const ssoUser = capability.isSso === true && user?.role === "user";
  if (
    !user?.enabled ||
    (capability.isSso === true && user.role !== "user") ||
    (user.mustChangePassword && !ssoUser) ||
    user.role !== capability.role
  ) {
    fail(
      401,
      "invalid_editor_capability",
      "Editor capability is invalid or expired"
    );
  }
  return {
    actor: {
      email: user.email,
      expiresAt: new Date(
        (capability.sessionExpiresAt ?? capability.expiresAt) * 1000
      ),
      id: user.id,
      isSso: capability.isSso === true,
      mustChangePassword: false,
      name: user.name,
      role: user.role,
    },
    capability,
  };
}
function actionEditorCapability(
  identity: EditorCapabilityIdentity,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">,
  action: Exclude<EditorCapabilityAction, "poll-operation">,
  lease: EditorLeaseGrant
): string {
  return createEditorCapability({
    ...scope,
    action,
    actorId: identity.id,
    expiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
    isSso: identity.isSso === true,
    leaseId: lease.id,
    leaseProof: lease.proof,
    role: identity.role,
    sessionExpiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
  });
}

function operationEditorCapability(
  identity: EditorCapabilityIdentity,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">,
  operationId: string
): string {
  return createEditorCapability({
    ...scope,
    action: "poll-operation",
    actorId: identity.id,
    expiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
    isSso: identity.isSso === true,
    operationId,
    role: identity.role,
    sessionExpiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
  });
}
async function requireActionEditorAuthorization(
  request: Request
): Promise<ActionEditorAuthorization> {
  const authorization = await editorAuthorization(request);
  const { capability } = authorization;
  if (!capability) {
    fail(
      401,
      "editor_capability_required",
      "An editor capability is required for this action"
    );
  }
  return { actor: authorization.actor, capability };
}

function requireEditorScope(
  authorization: EditorAuthorization,
  scope: EditorCapabilityScope
): void {
  const { capability } = authorization;
  if (!capability) {
    return;
  }
  if (
    capability.action !== scope.action ||
    capability.documentKey !== scope.documentKey ||
    capability.formId !== scope.formId ||
    capability.operationId !== scope.operationId ||
    capability.targetId !== scope.targetId ||
    capability.targetType !== scope.targetType
  ) {
    fail(
      403,
      "editor_capability_scope_mismatch",
      "Editor capability does not permit this action"
    );
  }
}

function editorLeaseTargetType(
  targetType: EditorCapabilityTarget
): OperationTargetType {
  if (targetType === "template-draft") {
    return OperationTargetType.template_draft;
  }
  return targetType === "correction"
    ? OperationTargetType.correction
    : OperationTargetType.response;
}

function editorLeaseProof(
  holder: Pick<Identity, "id" | "sessionId">,
  targetType: EditorCapabilityTarget,
  targetId: string
): string {
  return createHmac("sha256", env.EDITOR_CAPABILITY_SECRET)
    .update(
      [
        "editor-lease-v1",
        targetType,
        targetId,
        holder.sessionId,
        holder.id,
      ].join("\0")
    )
    .digest("base64url");
}

function nextEditorLeaseExpiry(identity: Identity, now: Date): Date {
  const expiresAt = new Date(
    Math.min(
      now.getTime() + editorLeaseDurationMs,
      identity.expiresAt.getTime()
    )
  );
  if (expiresAt.getTime() <= now.getTime()) {
    fail(401, "unauthorized", "Authentication is required");
  }
  return expiresAt;
}

function editorLeaseBridge(lease: ClaimedEditorLease) {
  return {
    expiresAt: lease.expiresAt.toISOString(),
    id: lease.id,
    releaseUrl: `/api/editor-leases/${lease.id}`,
    renewUrl: `/api/editor-leases/${lease.id}/renew`,
  };
}

async function claimEditorLease(
  identity: Identity,
  targetType: EditorCapabilityTarget,
  targetId: string,
  formId: string,
  workspace?: CorrectionWorkspaceInput
): Promise<EditorLeaseGrant> {
  const now = new Date();
  const expiresAt = nextEditorLeaseExpiry(identity, now);
  const proof = editorLeaseProof(identity, targetType, targetId);
  const capabilityDigest = tokenDigest(proof);
  const databaseTargetType = editorLeaseTargetType(targetType);
  const activeOperation = await prisma.operation.findFirst({
    where: {
      status: { in: [OperationStatus.pending, OperationStatus.processing] },
      targetId,
      targetType: databaseTargetType,
    },
  });
  if (activeOperation) {
    await expireOperationIfNeeded(activeOperation);
  }
  const lease = await prisma.$transaction(async (tx) => {
    const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`
        SELECT "id"
        FROM "forms"
        WHERE "id" = ${formId}::uuid
        FOR UPDATE
      `
    );
    if (!lockedForm) {
      fail(404, "not_found", "Form was not found");
    }
    const [current] = await tx.$queryRaw<
      (ActiveEditorLease & ClaimedEditorLease)[]
    >(
      Prisma.sql`
        SELECT
          "id",
          "capability_digest" AS "capabilityDigest",
          "expires_at" AS "expiresAt",
          "holder_session_id" AS "holderSessionId",
          "holder_user_id" AS "holderUserId",
          "workspace_base_document_key" AS "workspaceBaseDocumentKey",
          "workspace_base_revision" AS "workspaceBaseRevision",
          "workspace_document_key" AS "workspaceDocumentKey",
          "workspace_object_key" AS "workspaceObjectKey"
        FROM "editor_leases"
        WHERE
          "target_type" = CAST(${databaseTargetType} AS "OperationTargetType")
          AND "target_id" = ${targetId}::uuid
        FOR UPDATE
      `
    );
    const sameHolder =
      current?.holderSessionId === identity.sessionId &&
      current?.holderUserId === identity.id;
    const leaseId =
      sameHolder && current && current.expiresAt > now
        ? current.id
        : crypto.randomUUID();
    if (!sameHolder) {
      const activeOperation = await tx.operation.findFirst({
        select: { id: true },
        where: {
          status: {
            in: [OperationStatus.pending, OperationStatus.processing],
          },
          targetId,
          targetType: databaseTargetType,
        },
      });
      if (activeOperation) {
        return null;
      }
    }
    const [claimed] = await tx.$queryRaw<ClaimedEditorLease[]>(
      Prisma.sql`
        INSERT INTO "editor_leases" (
          "id",
          "target_type",
          "target_id",
          "holder_session_id",
          "holder_user_id",
          "capability_digest",
          "expires_at",
          "workspace_base_document_key",
          "workspace_base_revision",
          "workspace_document_key",
          "workspace_object_key",
          "renewed_at"
        )
        VALUES (
          ${leaseId}::uuid,
          CAST(${databaseTargetType} AS "OperationTargetType"),
          ${targetId}::uuid,
          ${identity.sessionId},
          ${identity.id},
          ${capabilityDigest},
          ${expiresAt},
          ${workspace?.baseDocumentKey ?? null},
          ${workspace?.baseRevision ?? null},
          ${workspace?.documentKey ?? null},
          ${workspace?.objectKey ?? null},
          ${now}
        )
        ON CONFLICT ("target_type", "target_id") DO UPDATE SET
          "id" = EXCLUDED."id",
          "holder_session_id" = EXCLUDED."holder_session_id",
          "holder_user_id" = EXCLUDED."holder_user_id",
          "capability_digest" = EXCLUDED."capability_digest",
          "expires_at" = EXCLUDED."expires_at",
          "workspace_base_document_key" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_base_document_key"
            ELSE "editor_leases"."workspace_base_document_key"
          END,
          "workspace_base_revision" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_base_revision"
            ELSE "editor_leases"."workspace_base_revision"
          END,
          "workspace_document_key" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_document_key"
            ELSE "editor_leases"."workspace_document_key"
          END,
          "workspace_object_key" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_object_key"
            ELSE "editor_leases"."workspace_object_key"
          END,
          "created_at" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              THEN EXCLUDED."created_at"
            ELSE "editor_leases"."created_at"
          END,
          "renewed_at" = EXCLUDED."renewed_at"
        WHERE
          "editor_leases"."expires_at" <= ${now}
          OR (
            "editor_leases"."holder_session_id" = ${identity.sessionId}
            AND "editor_leases"."holder_user_id" = ${identity.id}
          )
        RETURNING
          "id",
          "expires_at" AS "expiresAt",
          "workspace_base_document_key" AS "workspaceBaseDocumentKey",
          "workspace_base_revision" AS "workspaceBaseRevision",
          "workspace_document_key" AS "workspaceDocumentKey",
          "workspace_object_key" AS "workspaceObjectKey"
      `
    );
    return claimed ?? null;
  });
  if (!lease) {
    fail(409, "editor_in_use", "This document is open in another session");
  }
  return { ...lease, proof };
}

async function renewEditorLease(
  identity: Identity,
  leaseId: string
): Promise<ClaimedEditorLease> {
  const now = new Date();
  const expiresAt = nextEditorLeaseExpiry(identity, now);
  const renewed = await prisma.editorLease.updateMany({
    data: { expiresAt, renewedAt: now },
    where: {
      expiresAt: { gt: now },
      holderSessionId: identity.sessionId,
      holderUserId: identity.id,
      id: leaseId,
    },
  });
  if (renewed.count !== 1) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  return { expiresAt, id: leaseId };
}

async function releaseEditorLease(
  identity: Identity,
  leaseId: string
): Promise<void> {
  const workspaceObjectKey = await prisma.$transaction(async (tx) => {
    const [lease] = await tx.$queryRaw<
      {
        targetId: string;
        targetType: OperationTargetType;
        workspaceObjectKey: string | null;
      }[]
    >(
      Prisma.sql`
        SELECT
          "target_id" AS "targetId",
          "target_type" AS "targetType",
          "workspace_object_key" AS "workspaceObjectKey"
        FROM "editor_leases"
        WHERE
          "id" = ${leaseId}::uuid
          AND "holder_session_id" = ${identity.sessionId}
          AND "holder_user_id" = ${identity.id}
        FOR UPDATE
      `
    );
    if (!lease) {
      fail(
        409,
        "editor_lease_inactive",
        "The editor lease is no longer active"
      );
    }
    const activeOperation = await tx.operation.findFirst({
      select: { id: true },
      where: {
        status: {
          in: [OperationStatus.pending, OperationStatus.processing],
        },
        targetId: lease.targetId,
        targetType: lease.targetType,
      },
    });
    if (activeOperation) {
      return null;
    }
    const released = await tx.editorLease.deleteMany({
      where: { id: leaseId },
    });
    if (released.count !== 1) {
      fail(
        409,
        "editor_lease_inactive",
        "The editor lease is no longer active"
      );
    }
    return lease.workspaceObjectKey;
  });
  if (workspaceObjectKey) {
    await drainObjectCleanupIntents([workspaceObjectKey]);
  }
}

function editorLeaseClaims(authorization: ActionEditorAuthorization): {
  id: string;
  proof: string;
} {
  const { leaseId, leaseProof } = authorization.capability;
  if (!leaseId || !leaseProof) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  return { id: leaseId, proof: leaseProof };
}

function requireEditorLeaseProof(
  scope: Omit<EditorCapabilityScope, "action" | "operationId">,
  lease: ActiveEditorLease,
  proof: string
): void {
  const expectedProof = editorLeaseProof(
    { id: lease.holderUserId, sessionId: lease.holderSessionId },
    scope.targetType,
    scope.targetId
  );
  if (
    proof !== expectedProof ||
    tokenDigest(proof) !== lease.capabilityDigest
  ) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
}

async function requireActiveEditorLease(
  authorization: ActionEditorAuthorization,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">
): Promise<void> {
  const claim = editorLeaseClaims(authorization);
  const now = new Date();
  const lease = await prisma.editorLease.findFirst({
    select: {
      capabilityDigest: true,
      holderSessionId: true,
      holderUserId: true,
    },
    where: {
      expiresAt: { gt: now },
      holderSession: { expiresAt: { gt: now } },
      holderUserId: authorization.actor.id,
      id: claim.id,
      targetId: scope.targetId,
      targetType: editorLeaseTargetType(scope.targetType),
    },
  });
  if (!lease) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  requireEditorLeaseProof(scope, lease, claim.proof);
}

async function lockActiveEditorLease(
  tx: Prisma.TransactionClient,
  authorization: ActionEditorAuthorization,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">
): Promise<void> {
  const claim = editorLeaseClaims(authorization);
  const databaseTargetType = editorLeaseTargetType(scope.targetType);
  const [lease] = await tx.$queryRaw<ActiveEditorLease[]>(
    Prisma.sql`
      SELECT
        "editor_leases"."capability_digest" AS "capabilityDigest",
        "editor_leases"."holder_session_id" AS "holderSessionId",
        "editor_leases"."holder_user_id" AS "holderUserId"
      FROM "editor_leases"
      INNER JOIN "session"
        ON "session"."id" = "editor_leases"."holder_session_id"
      WHERE
        "editor_leases"."id" = ${claim.id}::uuid
        AND "editor_leases"."target_type" =
          CAST(${databaseTargetType} AS "OperationTargetType")
        AND "editor_leases"."target_id" = ${scope.targetId}::uuid
        AND "editor_leases"."holder_user_id" = ${authorization.actor.id}
        AND "editor_leases"."expires_at" > NOW()
        AND "session"."expires_at" > NOW()
      FOR UPDATE OF "editor_leases"
    `
  );
  if (!lease) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  requireEditorLeaseProof(scope, lease, claim.proof);
}

function requireAdmin(identity: Pick<Actor, "role">): void {
  if (identity.role !== "admin") {
    fail(403, "forbidden", "Administrator access is required");
  }
}

function validateId(value: string, label: string): string {
  if (!idPattern.test(value)) {
    fail(404, "not_found", `${label} was not found`);
  }
  return value;
}
function queryString(record: JsonRecord, key: string): string | undefined {
  const value = record[key];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    fail(400, "invalid_request", `${key} must be a string`);
  }
  return value;
}
interface AdminResultCursor {
  id: string;
  updatedAt: Date;
}

function adminResultCursor(
  value: string | undefined
): AdminResultCursor | null {
  if (value === undefined) {
    return null;
  }
  try {
    const decoded = JSON.parse(
      Buffer.from(value, "base64url").toString("utf-8")
    ) as { id?: unknown; updatedAt?: unknown };
    if (
      typeof decoded.id !== "string" ||
      !idPattern.test(decoded.id) ||
      typeof decoded.updatedAt !== "string"
    ) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    const updatedAt = new Date(decoded.updatedAt);
    if (!Number.isFinite(updatedAt.getTime())) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    return { id: decoded.id, updatedAt };
  } catch {
    fail(400, "invalid_request", "cursor is invalid");
  }
}

function adminResultCursorValue(result: {
  id: string;
  updatedAt: Date;
}): string {
  return Buffer.from(
    JSON.stringify({ id: result.id, updatedAt: result.updatedAt.toISOString() })
  ).toString("base64url");
}

function adminResultDate(value: string | undefined, key: string): Date | null {
  if (value === undefined) {
    return null;
  }
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) {
    fail(400, "invalid_request", `${key} is invalid`);
  }
  return date;
}
interface AuditEventCursor {
  createdAt: Date;
  id: string;
}

function auditEventCursor(value: string | undefined): AuditEventCursor | null {
  if (value === undefined) {
    return null;
  }
  try {
    const decoded = JSON.parse(
      Buffer.from(value, "base64url").toString("utf-8")
    ) as { createdAt?: unknown; id?: unknown };
    if (
      typeof decoded.createdAt !== "string" ||
      typeof decoded.id !== "string" ||
      !idPattern.test(decoded.id)
    ) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    const createdAt = new Date(decoded.createdAt);
    if (!Number.isFinite(createdAt.getTime())) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    return { createdAt, id: decoded.id };
  } catch {
    fail(400, "invalid_request", "cursor is invalid");
  }
}

function auditEventCursorValue(event: AuditEventCursor): string {
  return Buffer.from(
    JSON.stringify({
      createdAt: event.createdAt.toISOString(),
      id: event.id,
    })
  ).toString("base64url");
}

function auditFilterValue(
  value: string | undefined,
  key: string,
  maximumLength: number
): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  const trimmed = value.trim();
  if (!trimmed || trimmed.length > maximumLength) {
    fail(400, "invalid_request", `${key} is invalid`);
  }
  return trimmed;
}

function auditOutcomeValue(
  value: string | undefined
): AuditOutcome | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (value !== AuditOutcome.failure && value !== AuditOutcome.success) {
    fail(400, "invalid_request", "outcome is invalid");
  }
  return value;
}

const auditMetadataKeys = new Set([
  "change",
  "errorCode",
  "format",
  "revision",
  "source",
  "sourcePublicId",
  "state",
  "status",
]);

function safeAuditMetadata(value: unknown): JsonRecord {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }
  const source = value as JsonRecord;
  const result: JsonRecord = {};
  for (const key of auditMetadataKeys) {
    const item = source[key];
    if (
      item === null ||
      typeof item === "boolean" ||
      typeof item === "number" ||
      typeof item === "string"
    ) {
      result[key] = item;
    }
  }
  return result;
}

function normalizedAccountEmail(value: string): string {
  const email = normalizeEmail(value);
  if (
    email.length > accountEmailMaximumLength ||
    !accountEmailPattern.test(email)
  ) {
    fail(400, "invalid_request", "email must be a valid email address");
  }
  return email;
}
function accountAuditTargetId(value: string): string | null {
  return idPattern.test(value) ? value : null;
}

function accountUserSummary(user: AdminUser): AdminUser {
  return {
    createdAt: user.createdAt,
    email: user.email,
    enabled: user.enabled,
    id: user.id,
    mustChangePassword: user.mustChangePassword,
    name: user.name,
    role: user.role,
    updatedAt: user.updatedAt,
  };
}

function generateTemporaryPassword(): string {
  return randomBytes(24).toString("base64url");
}

async function lockAccountMutationActor(
  tx: Prisma.TransactionClient,
  identity: Identity
): Promise<void> {
  const [actor] = await tx.$queryRaw<{ id: string }[]>(
    Prisma.sql`
      SELECT "account_actor"."id"
      FROM "user" AS "account_actor"
      INNER JOIN "session" AS "account_session"
        ON "account_session"."user_id" = "account_actor"."id"
      WHERE
        "account_actor"."id" = ${identity.id}
        AND "account_actor"."role" = CAST('admin' AS "UserRole")
        AND "account_actor"."enabled" = TRUE
        AND "account_actor"."must_change_password" = FALSE
        AND "account_session"."id" = ${identity.sessionId}
        AND "account_session"."expires_at" > NOW()
      FOR UPDATE OF "account_actor", "account_session"
    `
  );
  if (!actor) {
    fail(403, "forbidden", "Administrator authorization changed");
  }
}

function accountTransaction<T>(
  identity: Identity,
  operation: (tx: Prisma.TransactionClient) => Promise<T>
): Promise<T> {
  return prisma.$transaction(async (tx) => {
    await tx.$executeRaw`SELECT pg_advisory_xact_lock(${accountMutationLockId})`;
    await lockAccountMutationActor(tx, identity);
    return operation(tx);
  });
}

async function lockAccountUser(
  tx: Prisma.TransactionClient,
  userId: string
): Promise<AdminUser> {
  const [locked] = await tx.$queryRaw<{ id: string }[]>(
    Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${userId} FOR UPDATE`
  );
  if (!locked) {
    fail(404, "not_found", "User was not found");
  }
  const user = await tx.user.findUnique({
    select: adminUserSelect,
    where: { id: userId },
  });
  if (!user) {
    fail(404, "not_found", "User was not found");
  }
  return user;
}

interface AccountAuditMetadata {
  change?: string;
  errorCode?: string;
}

async function createAccountAudit(
  tx: Prisma.TransactionClient,
  {
    action,
    actorId,
    outcome,
    safeMetadata,
    targetId,
  }: {
    action: AccountAuditAction;
    actorId: string | null;
    outcome: AuditOutcome;
    safeMetadata: AccountAuditMetadata;
    targetId: string | null;
  }
): Promise<void> {
  await tx.auditEvent.create({
    data: {
      action,
      actorId,
      outcome,
      safeMetadata: jsonValue(safeMetadata),
      targetId,
      targetType: "user",
    },
  });
}

function accountErrorCode(error: unknown): string {
  if (databaseErrorCode(error) === "P2002") {
    return "email_in_use";
  }
  return error instanceof HttpError ? error.code : "internal_error";
}

async function createAccountFailureAudit({
  action,
  actorId,
  error,
  targetId,
}: {
  action: AccountAuditAction;
  actorId: string | null;
  error: unknown;
  targetId: string | null;
}): Promise<void> {
  await prisma.auditEvent.create({
    data: {
      action,
      actorId,
      outcome: AuditOutcome.failure,
      safeMetadata: jsonValue({ errorCode: accountErrorCode(error) }),
      targetId,
      targetType: "user",
    },
  });
}

async function withAdminMutation<T>(
  request: Request,
  action: AccountAuditAction,
  targetId: string | null | (() => Promise<string | null>),
  operation: (
    identity: Identity,
    setAction: (action: AccountAuditAction) => void
  ) => Promise<T>
): Promise<T> {
  const identity = await identityFor(request);
  if (!identity) {
    fail(401, "unauthorized", "Authentication is required");
  }
  requireAdmin(identity);
  let auditAction = action;
  let auditTargetId = typeof targetId === "function" ? null : targetId;
  try {
    if (identity.mustChangePassword) {
      fail(403, "password_change_required", "Password replacement is required");
    }
    if (typeof targetId === "function") {
      auditTargetId = await targetId();
    }
    return await operation(identity, (nextAction) => {
      auditAction = nextAction;
    });
  } catch (error) {
    const normalizedError =
      databaseErrorCode(error) === "P2002"
        ? new HttpError(409, "email_in_use", "Email is already in use")
        : error;
    try {
      await createAccountFailureAudit({
        action: auditAction,
        actorId: identity.id,
        error: normalizedError,
        targetId: auditTargetId,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw normalizedError;
  }
}

async function reviewLegacyAccountLink(
  identity: Identity,
  requestId: string,
  decision: "approved" | "rejected"
): Promise<{ ok: true }> {
  const action =
    decision === "approved"
      ? "approve_legacy_account_link"
      : "reject_legacy_account_link";
  await accountTransaction(identity, async (tx) => {
    const requestSnapshot = await tx.legacyAccountLinkRequest.findUnique({
      select: { userId: true },
      where: { id: requestId },
    });
    if (!requestSnapshot) {
      fail(404, "not_found", "Account link request was not found");
    }
    const user = await lockAccountUser(tx, requestSnapshot.userId);
    const [lockedRequest] = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`SELECT "id" FROM "legacy_account_link_requests" WHERE "id" = ${requestId}::uuid FOR UPDATE`
    );
    if (!lockedRequest) {
      fail(404, "not_found", "Account link request was not found");
    }
    const linkRequest = await tx.legacyAccountLinkRequest.findUnique({
      where: { id: requestId },
    });
    if (!linkRequest) {
      fail(404, "not_found", "Account link request was not found");
    }
    const generation =
      decision === "approved"
        ? await nextLegacySsoGeneration(tx, linkRequest.providerId)
        : null;
    if (linkRequest.status !== LegacyAccountLinkStatus.pending) {
      fail(
        409,
        "account_link_not_pending",
        "Account link request is not pending"
      );
    }
    if (decision === "approved") {
      if (
        !user.enabled ||
        user.role !== "user" ||
        user.email !== linkRequest.email
      ) {
        fail(
          409,
          "account_link_not_eligible",
          "Candidate account is not eligible for linking"
        );
      }
      const account = await tx.account.findUnique({
        select: { userId: true },
        where: {
          providerId_accountId: {
            accountId: linkRequest.subject,
            providerId: linkRequest.providerId,
          },
        },
      });
      if (account) {
        fail(
          409,
          "identity_already_linked",
          "Legacy identity is already linked"
        );
      }
    }
    const [databaseTime] = await tx.$queryRaw<{ currentTime: Date }[]>(
      Prisma.sql`SELECT CURRENT_TIMESTAMP AS "currentTime"`
    );
    if (!databaseTime) {
      throw new Error("Database did not return current time");
    }
    const updated = await tx.legacyAccountLinkRequest.updateMany({
      data: {
        reviewedAt: databaseTime.currentTime,
        reviewedById: identity.id,
        reviewedGeneration: generation,
        status:
          decision === "approved"
            ? LegacyAccountLinkStatus.approved
            : LegacyAccountLinkStatus.rejected,
        updatedAt: databaseTime.currentTime,
      },
      where: { id: requestId, status: LegacyAccountLinkStatus.pending },
    });
    if (updated.count !== 1) {
      fail(
        409,
        "account_link_not_pending",
        "Account link request is not pending"
      );
    }
    await createAccountAudit(tx, {
      action,
      actorId: identity.id,
      outcome: AuditOutcome.success,
      safeMetadata: { change: decision },
      targetId: user.id,
    });
  });
  return { ok: true };
}

const formAuditErrorCodes: Record<string, true> = {
  blank_template_unavailable: true,
  callback_claim_invalid: true,
  callback_document_unavailable: true,
  callback_key_mismatch: true,
  callback_processing_failed: true,
  document_save_failed: true,
  document_unavailable: true,
  editor_capability_required: true,
  editor_capability_scope_mismatch: true,
  editor_in_use: true,
  editor_lease_inactive: true,
  fill_method_changed: true,
  force_save_failed: true,
  form_has_responses: true,
  form_not_draft: true,
  handoff_unavailable: true,
  internal_error: true,
  invalid_editor_capability: true,
  invalid_file_type: true,
  invalid_request: true,
  invalid_template: true,
  native_fill_unsupported: true,
  not_found: true,
  onlyoffice_document_error: true,
  operation_in_progress: true,
  operation_timeout: true,
  payload_too_large: true,
  published_immutable: true,
  stale_document: true,
  stale_operation: true,
  unauthorized: true,
};

function formAuditErrorCodeFromCode(code: string): FormAuditErrorCode {
  return formAuditErrorCodes[code]
    ? (code as FormAuditErrorCode)
    : "internal_error";
}

function formAuditErrorCode(error: unknown): FormAuditErrorCode {
  const code = error instanceof HttpError ? error.code : "internal_error";
  return formAuditErrorCodeFromCode(code);
}

async function createFormAudit(
  tx: Prisma.TransactionClient,
  {
    action,
    actorId,
    outcome,
    safeMetadata,
    targetId,
  }: {
    action: FormAuditAction;
    actorId: string | null;
    outcome: AuditOutcome;
    safeMetadata: FormAuditMetadata;
    targetId: string | null;
  }
): Promise<void> {
  await tx.auditEvent.create({
    data: {
      action,
      actorId,
      outcome,
      safeMetadata: jsonValue(safeMetadata),
      targetId,
      targetType: "form",
    },
  });
}

async function createFormFailureAudit({
  action,
  actorId,
  error,
  source,
  targetId,
}: {
  action: FormAuditAction;
  actorId: string | null;
  error: unknown;
  source?: FormSource;
  targetId: string | null;
}): Promise<void> {
  await prisma.auditEvent.create({
    data: {
      action,
      actorId,
      outcome: AuditOutcome.failure,
      safeMetadata: jsonValue({
        ...(source ? { source } : {}),
        errorCode: formAuditErrorCode(error),
      }),
      targetId,
      targetType: "form",
    },
  });
}
type ResponseAuditAction =
  | "create_correction"
  | "export_correction"
  | "export_response"
  | "view_correction"
  | "view_response";
type ResponseAuditTargetType = "correction" | "response" | "submission";
interface ResponseAuditMetadata {
  errorCode?: OperationErrorCode;
  format?: "docx" | "json" | "pdf";
  revision?: number;
  state: "draft" | "submitted";
}

async function createResponseAudit({
  action,
  actorId,
  outcome,
  safeMetadata,
  targetId,
  targetType,
}: {
  action: ResponseAuditAction;
  actorId: string;
  outcome: AuditOutcome;
  safeMetadata: ResponseAuditMetadata;
  targetId: string;
  targetType: ResponseAuditTargetType;
}): Promise<void> {
  await prisma.auditEvent.create({
    data: {
      action,
      actorId,
      outcome,
      safeMetadata: jsonValue(safeMetadata),
      targetId,
      targetType,
    },
  });
}
type ResponseDeletionAuditAction =
  | "delete_response"
  | "delete_response_pending";
interface ResponseDeletionAuditMetadata {
  errorCode?: string;
}

async function createResponseDeletionAudit({
  action = "delete_response",
  actorId,
  error,
  outcome,
  targetId,
  tx = prisma,
}: {
  action?: ResponseDeletionAuditAction;
  actorId: string;
  error?: unknown;
  outcome: AuditOutcome;
  targetId: string;
  tx?: Prisma.TransactionClient;
}): Promise<void> {
  await tx.auditEvent.create({
    data: {
      action,
      actorId,
      outcome,
      safeMetadata: jsonValue(
        error
          ? {
              errorCode:
                error instanceof HttpError ? error.code : "internal_error",
            }
          : ({} satisfies ResponseDeletionAuditMetadata)
      ),
      targetId,
      targetType: "response",
    },
  });
}

interface ResponseDeletionResult {
  alreadyDeleted: boolean;
  deleted: boolean;
  ownerSessionIds: string[];
}

function responseDeletionPrefixes(
  responseId: string,
  submissionId: string | null,
  operationPrefixes: readonly string[]
): string[] {
  return [
    `responses/${responseId}/`,
    ...(submissionId ? [`submissions/${submissionId}/`] : []),
    ...operationPrefixes,
  ];
}

function responseDeletionKeyBelongs(
  key: string,
  prefixes: readonly string[]
): boolean {
  return prefixes.some((prefix) => key.startsWith(prefix));
}

async function deleteResponseData({
  actor,
  aiAuthoring,
  allowActiveLease,
  missingOk,
  removeObject,
  responseId,
  revokeOwnerSessions,
}: {
  actor: Identity;
  aiAuthoring: AiAuthoringSessions;
  allowActiveLease: boolean;
  missingOk: boolean;
  removeObject: (key: string) => Promise<void>;
  responseId: string;
  revokeOwnerSessions: boolean;
}): Promise<ResponseDeletionResult> {
  const responseLookupDigest = tokenDigest(responseId);
  const existingTombstone = await prisma.deletionTombstone.findUnique({
    where: { responseLookupDigest },
  });
  if (existingTombstone) {
    await drainObjectCleanupIntents(
      undefined,
      responseLookupDigest,
      removeObject
    );
    const remaining = await prisma.objectCleanupIntent.count({
      where: { deletionResponseLookupDigest: responseLookupDigest },
    });
    if (remaining > 0) {
      fail(503, "deletion_cleanup_failed", "Response objects remain");
    }
    return { alreadyDeleted: true, deleted: true, ownerSessionIds: [] };
  }
  const result = await prisma.$transaction(
    async (tx) => {
      const [lockedResponse] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "responses"
          WHERE "id" = ${responseId}::uuid
          FOR UPDATE
        `
      );
      if (!lockedResponse) {
        if (missingOk) {
          return {
            alreadyDeleted: false,
            deleted: false,
            ownerSessionIds: [],
          };
        }
        fail(404, "not_found", "Response was not found");
      }
      const response = await tx.response.findUnique({
        include: {
          corrections: { orderBy: { revision: "asc" } },
          submission: true,
        },
        where: { id: responseId },
      });
      if (!response) {
        if (missingOk) {
          return {
            alreadyDeleted: false,
            deleted: false,
            ownerSessionIds: [],
          };
        }
        fail(404, "not_found", "Response was not found");
      }
      if (actor.role !== "admin" && response.userId !== actor.id) {
        fail(403, "forbidden", "You may only delete your own Draft");
      }
      if (actor.role !== "admin" && response.status !== ResponseStatus.draft) {
        fail(409, "draft_unavailable", "Only a Draft can be discarded");
      }
      const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "forms"
          WHERE "id" = ${response.formId}::uuid
          FOR UPDATE
        `
      );
      if (!lockedForm) {
        fail(409, "stale_response", "The Response form is unavailable");
      }
      const operations = await tx.operation.findMany({
        select: {
          id: true,
          metadata: true,
          stagingObjectKey: true,
          status: true,
        },
        where: {
          OR: [
            { responseId: response.id },
            ...(response.submission
              ? [{ submissionId: response.submission.id }]
              : []),
            ...(response.corrections.length > 0
              ? [
                  {
                    correctionId: {
                      in: response.corrections.map(
                        (correction) => correction.id
                      ),
                    },
                  },
                ]
              : []),
          ],
        },
      });
      if (
        operations.some(
          (operation) =>
            operation.status === OperationStatus.pending ||
            operation.status === OperationStatus.processing
        )
      ) {
        fail(409, "operation_in_progress", "The Response operation is active");
      }
      const correctionIds = response.corrections.map(
        (correction) => correction.id
      );
      const leases = await tx.editorLease.findMany({
        select: { id: true, workspaceObjectKey: true },
        where: {
          OR: [
            {
              targetId: response.id,
              targetType: OperationTargetType.response,
            },
            {
              targetId: response.id,
              targetType: OperationTargetType.correction,
            },
          ],
        },
      });
      if (leases.length > 0 && !allowActiveLease) {
        fail(409, "editor_in_use", "The Response is open in an editor");
      }

      const candidateKeys = [
        response.draftObjectKey,
        response.submission?.objectKey,
        ...response.corrections.map((correction) => correction.objectKey),
        ...leases.map((lease) => lease.workspaceObjectKey),
        ...operations.flatMap((operation) => {
          const metadata = operationMetadata(operation.metadata);
          return [
            operation.stagingObjectKey,
            metadata.finalObjectKey,
            metadata.stagedObjectKey,
            ...(metadata.cleanupObjectKeys ?? []),
            metadata.workspaceObjectKey,
          ];
        }),
      ];
      const linkedObjectPrefixes = [
        ...new Set(
          uniqueObjectKeys(candidateKeys).flatMap((key) => {
            const [scope, pathId] = key.split("/", 3);
            return (scope === "operations" || scope === "submissions") && pathId
              ? [`${scope}/${pathId}/`]
              : [];
          })
        ),
      ];
      const prefixes = responseDeletionPrefixes(
        response.id,
        response.submission?.id ?? null,
        linkedObjectPrefixes
      );
      for (const key of uniqueObjectKeys(candidateKeys)) {
        if (!responseDeletionKeyBelongs(key, prefixes)) {
          fail(
            500,
            "invalid_object_key",
            "Response object ownership is invalid"
          );
        }
      }
      const cleanupIntents = await tx.objectCleanupIntent.findMany({
        select: { objectKey: true },
        where: {
          OR: prefixes.map((prefix) => ({
            objectKey: { startsWith: prefix },
          })),
        },
      });
      const objectKeys = uniqueObjectKeys([
        ...candidateKeys,
        ...cleanupIntents.map((intent) => intent.objectKey),
      ]);
      for (const objectKeyValue of objectKeys) {
        await tx.objectCleanupIntent.upsert({
          create: {
            deletionOwnerUserId: response.userId,
            deletionResponseLookupDigest: responseLookupDigest,
            objectKey: objectKeyValue,
          },
          update: {
            cleanupAfter: new Date(),
            deletionOwnerUserId: response.userId,
            deletionResponseLookupDigest: responseLookupDigest,
          },
          where: { objectKey: objectKeyValue },
        });
      }

      const handoffs = await tx.handoff.findMany({
        select: { id: true },
        where: {
          OR: [
            { responseId: response.id },
            ...(response.externalReferenceDigest
              ? [{ externalReferenceDigest: response.externalReferenceDigest }]
              : []),
          ],
        },
      });
      if (handoffs.length > 0) {
        const handoffIds = handoffs.map((handoff) => handoff.id);
        await tx.pendingClaim.deleteMany({
          where: { handoffId: { in: handoffIds } },
        });
        await tx.handoff.updateMany({
          data: {
            codeDigest: null,
            configurationHash: null,
            consumedAt: null,
            deletionResponseLookupDigest: responseLookupDigest,
            filteredValues: Prisma.JsonNull,
            formId: null,
            normalizedEmail: null,
            reservedAt: null,
            responseId: null,
            status: HandoffStatus.deleted,
          },
          where: { id: { in: handoffIds } },
        });
      }
      await tx.editorLease.deleteMany({
        where: {
          OR: [
            {
              targetId: response.id,
              targetType: OperationTargetType.response,
            },
            {
              targetId: response.id,
              targetType: OperationTargetType.correction,
            },
          ],
        },
      });
      await tx.operation.deleteMany({
        where: {
          OR: [
            { responseId: response.id },
            ...(response.submission
              ? [{ submissionId: response.submission.id }]
              : []),
            ...(correctionIds.length > 0
              ? [{ correctionId: { in: correctionIds } }]
              : []),
          ],
        },
      });
      const ownerSessionIds = revokeOwnerSessions
        ? (
            await tx.session.findMany({
              select: { id: true },
              where: { userId: response.userId },
            })
          ).map(({ id }) => id)
        : [];
      if (revokeOwnerSessions) {
        await tx.session.deleteMany({ where: { userId: response.userId } });
      }
      await tx.correction.deleteMany({ where: { responseId: response.id } });
      await tx.submission.deleteMany({ where: { responseId: response.id } });
      await tx.prefillSnapshot.deleteMany({
        where: { responseId: response.id },
      });
      await tx.response.delete({ where: { id: response.id } });
      await tx.deletionTombstone.create({
        data: {
          actorId: actor.id,
          externalReferenceDigest: response.externalReferenceDigest,
          id: crypto.randomUUID(),
          outcome: AuditOutcome.success,
          responseLookupDigest,
        },
      });
      await createResponseDeletionAudit({
        action: "delete_response_pending",
        actorId: actor.id,
        outcome: AuditOutcome.success,
        targetId: response.id,
        tx,
      });
      return { alreadyDeleted: false, deleted: true, ownerSessionIds };
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  if (!result.deleted) {
    return result;
  }
  const cleanupResults = await Promise.allSettled([
    endAiAuthoringSessions(aiAuthoring, result.ownerSessionIds),
    (async () => {
      await drainObjectCleanupIntents(
        undefined,
        responseLookupDigest,
        removeObject
      );
      const remaining = await prisma.objectCleanupIntent.count({
        where: { deletionResponseLookupDigest: responseLookupDigest },
      });
      if (remaining > 0) {
        fail(503, "deletion_cleanup_failed", "Response objects remain");
      }
    })(),
  ]);
  const cleanupErrors = cleanupResults.flatMap((cleanup) =>
    cleanup.status === "rejected" ? [cleanup.reason] : []
  );
  if (cleanupErrors.length === 1) {
    throw cleanupErrors[0];
  }
  if (cleanupErrors.length > 1) {
    throw new AggregateError(cleanupErrors, "Response deletion cleanup failed");
  }
  await createResponseDeletionAudit({
    actorId: actor.id,
    outcome: AuditOutcome.success,
    targetId: responseId,
  });
  return result;
}

interface FormCounts {
  activeDraftCount: number;
  submissionCount: number;
}

function formDto(
  form: FormWithDocuments,
  { activeDraftCount, submissionCount }: FormCounts
): JsonRecord {
  return {
    activeDraftCount,
    createdAt: form.createdAt,
    description: form.description ?? "",
    fillMethod: form.fillMethod,
    hasTemplateDraft: Boolean(form.templateDraft),
    publicId: form.publicId,
    status: form.status,
    submissionCount,
    title: form.title,
    updatedAt: form.updatedAt,
    version: form.version,
  };
}
function supportsNativeFill(fields: readonly { type: FieldType }[]): boolean {
  return fields.length > 0;
}

interface ManifestDisplayField {
  label: string;
  placeholder: string | null;
  position: number;
  tag: string;
}

async function manifestFieldsWithDocumentMetadata<
  T extends ManifestDisplayField,
>(
  objectKey: string,
  displayMetadataVersion: number,
  fields: T[],
  documentBytes?: Uint8Array
): Promise<T[]> {
  if (displayMetadataVersion !== 0) {
    return fields;
  }
  const documentFields = parseTemplateFields(
    documentBytes ?? (await readObject(objectKey))
  );
  if (documentFields.length !== fields.length) {
    fail(
      500,
      "internal_error",
      "The published Field Manifest does not match its document"
    );
  }
  const fieldsByTag = new Map(fields.map((field) => [field.tag, field]));
  return documentFields.map((documentField, position) => {
    const field = fieldsByTag.get(documentField.tag);
    if (!field) {
      fail(
        500,
        "internal_error",
        "The published Field Manifest does not match its document"
      );
    }
    return {
      ...field,
      label: documentField.label,
      placeholder: documentField.placeholder,
      position,
    };
  });
}

async function supportsNativeTemplate(
  objectKey: string | null | undefined,
  fields: readonly { tag: string; type: FieldType }[],
  documentBytes?: Uint8Array
): Promise<boolean> {
  if (!objectKey || !supportsNativeFill(fields)) {
    return false;
  }
  const { archive, xmlPaths } = safeTemplateArchive(
    documentBytes ?? (await readObject(objectKey))
  );
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  if (
    templatePartsHaveUnsupportedDateSettings(archive, controlParts) ||
    templatePartsHaveNestedControls(archive, controlParts)
  ) {
    return false;
  }
  const fieldTags = new Set(fields.map(({ tag }) => tag));
  for (const archivePath of controlParts) {
    if (
      !nativeFlattenAlternateFieldsXml(
        templateArchiveText(archive, archivePath),
        fieldTags,
        true
      ).supported
    ) {
      return false;
    }
  }
  return true;
}
async function nativeFields(
  publishedTemplateId: string,
  responseDocumentBytes: Uint8Array
): Promise<{
  fields: {
    label: string;
    options: { displayText: string; value: string }[];
    pictureMaxBytes: number | null;
    pictureMaxHeight: number | null;
    pictureMaxWidth: number | null;
    placeholder: string | null;
    position: number;
    required: boolean;
    tag: string;
    type: FieldType;
  }[];
  pictures: Record<string, boolean>;
}> {
  const [manifest, publishedTemplate] = await Promise.all([
    prisma.fieldManifest.findUnique({
      include: { fields: { orderBy: { position: "asc" } } },
      where: { publishedTemplateId },
    }),
    prisma.publishedTemplate.findUnique({
      select: { objectKey: true },
      where: { id: publishedTemplateId },
    }),
  ]);
  if (!manifest || !publishedTemplate?.objectKey) {
    fail(
      409,
      "native_fill_unsupported",
      "Native filling requires supported fields"
    );
  }
  const documentBytes = await readObject(publishedTemplate.objectKey);
  if (
    !(await supportsNativeTemplate(
      publishedTemplate.objectKey,
      manifest.fields,
      documentBytes
    ))
  ) {
    fail(
      409,
      "native_fill_unsupported",
      "Native filling requires supported fields"
    );
  }
  const fields = await manifestFieldsWithDocumentMetadata(
    publishedTemplate.objectKey,
    manifest.displayMetadataVersion,
    manifest.fields,
    documentBytes
  );
  return {
    fields: fields.map((field) => ({
      ...receiptField(field),
      pictureMaxBytes: field.pictureMaxBytes,
      pictureMaxHeight: field.pictureMaxHeight,
      pictureMaxWidth: field.pictureMaxWidth,
      required: field.required,
    })),
    pictures: responsePicturePresence(
      responseDocumentBytes,
      manifest.fields.filter(({ type }) => type === FieldType.picture)
    ),
  };
}

async function nativePictureManifestFields(
  publishedTemplateId: string
): Promise<ResponsePictureManifestField[]> {
  const manifest = await prisma.fieldManifest.findUnique({
    select: {
      fields: {
        select: {
          pictureMaxBytes: true,
          pictureMaxHeight: true,
          pictureMaxWidth: true,
          required: true,
          tag: true,
          type: true,
        },
      },
    },
    where: { publishedTemplateId },
  });
  if (!manifest) {
    fail(500, "internal_error", "The published Field Manifest is unavailable");
  }
  return manifest.fields.filter(({ type }) => type === FieldType.picture);
}

function responseSummary(
  response: Response,
  extra: {
    formPublicId?: string;
    formTitle?: string;
    latestCorrectionNumber?: number | null;
    submissionId?: string | null;
    submittedAt?: Date | null;
  } = {}
): JsonRecord {
  return {
    createdAt: response.createdAt,
    formPublicId: extra.formPublicId,
    formTitle: extra.formTitle,
    hasDraft: Boolean(response.draftObjectKey && response.draftData),
    id: response.id,
    latestCorrectionNumber: extra.latestCorrectionNumber,
    publishedVersion: response.publishedVersion,
    status: response.status,
    submissionId: extra.submissionId,
    submittedAt: extra.submittedAt,
    updatedAt: response.updatedAt,
  };
}

function submissionSummary(
  submission: Submission,
  extra: {
    formPublicId?: string;
    formTitle?: string;
    userEmail?: string;
  } = {}
): JsonRecord {
  return {
    createdAt: submission.createdAt,
    formPublicId: extra.formPublicId,
    formTitle: extra.formTitle,
    id: submission.id,
    responseId: submission.responseId,
    status: "submitted",
    submittedAt: submission.createdAt,
    userEmail: extra.userEmail,
    userId: submission.userId,
  };
}

function adminResultSummary(response: {
  corrections: { revision: number }[];
  createdAt: Date;
  form: { publicId: string; title: string };
  id: string;
  owner: { email: string };
  status: ResponseStatus;
  submission: { createdAt: Date; id: string } | null;
  updatedAt: Date;
}): JsonRecord {
  const submitted = response.status === ResponseStatus.submitted;
  return {
    createdAt: response.createdAt,
    formPublicId: response.form.publicId,
    formTitle: response.form.title,
    id: response.id,
    latestCorrectionNumber: response.corrections[0]?.revision ?? null,
    state: submitted ? "submitted" : "draft",
    submissionId: response.submission?.id ?? null,
    submittedAt: response.submission?.createdAt ?? null,
    updatedAt: response.updatedAt,
    userEmail: response.owner.email,
  };
}
async function findTemplateSource(): Promise<string | null> {
  const configuredSource = Bun.file(env.TEMPLATE_PATH);
  if (await configuredSource.exists()) {
    return env.TEMPLATE_PATH;
  }
  const fallbackSource = Bun.file(fallbackTemplatePath);
  return (await fallbackSource.exists()) ? fallbackTemplatePath : null;
}
async function readinessStatus(): Promise<boolean> {
  try {
    const rustfsReady = await fetch(
      new URL("/health/ready", env.RUSTFS_ENDPOINT),
      { signal: AbortSignal.timeout(2000) }
    );
    const [templateSource] = await Promise.all([
      findTemplateSource(),
      prisma.$queryRaw`SELECT 1`,
    ]);
    return rustfsReady.ok && templateSource !== null;
  } catch {
    return false;
  }
}
export function resolveCallbackDocumentUrl(
  value: unknown,
  allowedOrigins: ReadonlySet<string> = callbackOrigins,
  publicBaseUrl: string | null = env.ONLYOFFICE_URL,
  internalBaseUrl: string | null = env.ONLYOFFICE_INTERNAL_URL
): string | null {
  if (typeof value !== "string") {
    return null;
  }
  try {
    let url = new URL(value);
    if (
      !["http:", "https:"].includes(url.protocol) ||
      url.username ||
      url.password ||
      url.hash ||
      !allowedOrigins.has(url.origin)
    ) {
      return null;
    }
    const publicBase = publicBaseUrl ? new URL(publicBaseUrl) : null;
    const internalBase = internalBaseUrl ? new URL(internalBaseUrl) : null;
    if (
      publicBase &&
      internalBase &&
      publicBase.toString() !== internalBase.toString() &&
      url.origin === publicBase.origin
    ) {
      const publicPrefix = publicBase.pathname.replace(/\/+$/u, "");
      const matchesPublicPrefix =
        publicPrefix === "" ||
        url.pathname === publicPrefix ||
        url.pathname.startsWith(`${publicPrefix}/`);
      if (matchesPublicPrefix) {
        const suffix =
          publicPrefix === ""
            ? url.pathname
            : url.pathname.slice(publicPrefix.length);
        const internalPath = internalBase.pathname.replace(/\/+$/u, "");
        internalBase.pathname = `${internalPath}${suffix}` || "/";
      } else {
        internalBase.pathname = url.pathname;
      }
      internalBase.search = url.search;
      url = internalBase;
    }
    return url.toString();
  } catch {
    return null;
  }
}

async function readCallbackDocument(
  url: string,
  maximumBytes = maxCallbackDocumentBytes
): Promise<Uint8Array> {
  const response = await fetch(url, {
    headers: { Authorization: createOnlyOfficeAuthorization({ url }) },
    redirect: "error",
  });
  if (!response.ok) {
    throw new Error(
      `Failed to download ONLYOFFICE document: HTTP ${response.status}`
    );
  }
  const contentLength = Number(response.headers.get("content-length") ?? "");
  if (Number.isFinite(contentLength) && contentLength > maximumBytes) {
    throw new Error("ONLYOFFICE document callback payload is too large");
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
      throw new Error("ONLYOFFICE document callback payload is too large");
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

const templatePackageRelationshipNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const templatePackageContentTypesNamespace =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const templateOfficeDocumentRelationships = new Set([
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
]);
const templateWordMainNamespace =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const templateWordStrictNamespace =
  "http://purl.oclc.org/ooxml/wordprocessingml/main";
const templateCheckboxNamespace =
  "http://schemas.microsoft.com/office/word/2010/wordml";
const templateWord2012Namespace =
  "http://schemas.microsoft.com/office/word/2012/wordml";
const templateMarkupCompatibilityNamespace =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const templateXmlnsNamespace = "http://www.w3.org/2000/xmlns/";
const templateWordNamespaces = new Set([
  templateWordStrictNamespace,
  templateWordMainNamespace,
]);
const templateDrawingMlNamespaces = new Set([
  "http://purl.oclc.org/ooxml/drawingml/main",
  "http://schemas.openxmlformats.org/drawingml/2006/main",
]);
const templateOfficeRelationshipNamespaces = new Set([
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
]);
const templateGlossaryDocumentRelationships = new Set([
  "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
]);
const templateVmlNamespace = "urn:schemas-microsoft-com:vml";
const responsePictureJpegSofMarkers = new Set([
  0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf,
]);
const responsePictureJpegStandaloneMarkers = new Set([
  0x01, 0xd0, 0xd1, 0xd2, 0xd3, 0xd4, 0xd5, 0xd6, 0xd7, 0xd8,
]);
const responsePicturePngSignature = [
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
] as const;
const templateControlNamespaces = new Set([
  ...templateWordNamespaces,
  templateCheckboxNamespace,
  templateWord2012Namespace,
]);
const templateControlKey = (namespace: string, local: string): string =>
  `${namespace}#${local}`;
const templateDocumentContentType =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

interface TemplateXmlAttribute {
  local: string;
  uri: string;
  value: string;
}

interface TemplateXmlElement {
  attributes: TemplateXmlAttribute[];
  local: string;
  uri: string;
}

interface TemplateXmlVisitor {
  close?: (element: TemplateXmlElement) => void;
  open?: (element: TemplateXmlElement) => void;
  text?: (value: string) => void;
}

function templateXmlElement(tag: {
  attributes: Record<string, { local: string; uri: string; value: string }>;
  local: string;
  uri: string;
}): TemplateXmlElement {
  return {
    attributes: Object.values(tag.attributes).map((attribute) => ({
      local: attribute.local,
      uri: attribute.uri,
      value: attribute.value,
    })),
    local: tag.local,
    uri: tag.uri,
  };
}

function parseTemplateXml(xml: string, visitor: TemplateXmlVisitor = {}): void {
  const parser = new SaxesParser({ position: false, xmlns: true });
  parser.on("doctype", () => {
    throw new Error("DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    visitor.open?.(templateXmlElement(tag));
  });
  parser.on("closetag", (tag) => {
    visitor.close?.(templateXmlElement(tag));
  });
  parser.on("text", (value) => {
    visitor.text?.(value);
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
}

function templateAttribute(
  element: TemplateXmlElement,
  local: string,
  uri = ""
): string | undefined {
  return element.attributes.find(
    (attribute) => attribute.local === local && attribute.uri === uri
  )?.value;
}

function readTemplateArchive(
  bytes: Uint8Array,
  include: (archivePath: string) => boolean
): Record<string, Uint8Array> {
  let expandedBytes = 0;
  let entryCount = 0;
  const names = new Set<string>();
  try {
    return unzipSync(bytes, {
      filter: (file) => {
        entryCount += 1;
        const archivePath = file.name;
        const segments = archivePath.split("/");
        if (
          !archivePath ||
          archivePath.startsWith("/") ||
          archivePath.includes("\\") ||
          archivePath.includes("\0") ||
          segments.some(
            (segment, index) =>
              segment === "." ||
              segment === ".." ||
              (segment.length === 0 && index !== segments.length - 1)
          ) ||
          names.has(archivePath)
        ) {
          throw new Error("Unsafe DOCX archive path");
        }
        names.add(archivePath);
        if (
          !Number.isSafeInteger(file.originalSize) ||
          file.originalSize < 0 ||
          expandedBytes > maxTemplateArchiveExpandedBytes - file.originalSize
        ) {
          throw new Error("DOCX archive expands beyond the safety limit");
        }
        expandedBytes += file.originalSize;
        if (entryCount > maxTemplateArchiveEntries) {
          throw new Error("DOCX archive contains too many entries");
        }
        return include(archivePath);
      },
    });
  } catch {
    fail(422, "invalid_template", "The template is not a safe DOCX archive");
  }
}

function templateArchiveText(
  archive: Record<string, Uint8Array>,
  archivePath: string
): string {
  const bytes = archive[archivePath];
  if (!bytes || bytes.byteLength === 0) {
    fail(422, "invalid_template", "The DOCX package is incomplete");
  }
  try {
    let decodedBytes = bytes;
    let encoding: "utf-16" | "utf-8" = "utf-8";
    if (
      (bytes[0] === 0xff && bytes[1] === 0xfe) ||
      (bytes[0] === 0x3c && bytes[1] === 0x00)
    ) {
      encoding = "utf-16";
    } else if (
      (bytes[0] === 0xfe && bytes[1] === 0xff) ||
      (bytes[0] === 0x00 && bytes[1] === 0x3c)
    ) {
      if (bytes.byteLength % 2 !== 0) {
        throw new Error("UTF-16 XML must contain complete code units");
      }
      decodedBytes = new Uint8Array(bytes);
      for (let index = 0; index < decodedBytes.byteLength; index += 2) {
        const firstByte = decodedBytes[index] ?? 0;
        decodedBytes[index] = decodedBytes[index + 1] ?? 0;
        decodedBytes[index + 1] = firstByte;
      }
      encoding = "utf-16";
    }
    return new TextDecoder(encoding, { fatal: true }).decode(decodedBytes);
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
}
const externalRelationshipTargetPattern = /^[a-z][a-z\d+.-]*:/iu;

function isExternalRelationshipTarget(target: string | undefined): boolean {
  const normalized = target?.trim() ?? "";
  return (
    normalized.startsWith("//") ||
    normalized.startsWith("\\\\") ||
    externalRelationshipTargetPattern.test(normalized)
  );
}

function validateOfficeRelationships(
  archive: Record<string, Uint8Array>
): void {
  for (const archivePath of Object.keys(archive)) {
    if (!archivePath.toLowerCase().endsWith(".rels")) {
      continue;
    }
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      open: (element) => {
        if (
          element.local !== "Relationship" ||
          element.uri !== templatePackageRelationshipNamespace
        ) {
          return;
        }
        const targetMode = templateAttribute(element, "TargetMode");
        const target = templateAttribute(element, "Target");
        if (
          targetMode?.trim().toLowerCase() === "external" ||
          isExternalRelationshipTarget(target)
        ) {
          fail(
            422,
            "invalid_template",
            "External DOCX relationships are not allowed"
          );
        }
      },
    });
  }
}

function validateOfficeRelationshipsBytes(bytes: Uint8Array): void {
  const archive = readTemplateArchive(bytes, (archivePath) =>
    archivePath.toLowerCase().endsWith(".rels")
  );
  validateOfficeRelationships(archive);
}

function isTemplateXmlContentType(contentType: string | undefined): boolean {
  const normalized = contentType?.split(";", 1)[0]?.trim().toLowerCase();
  return (
    normalized === "application/xml" ||
    normalized === "text/xml" ||
    normalized === "application/vnd.openxmlformats-officedocument.vmldrawing" ||
    normalized?.endsWith("+xml") === true
  );
}

function templateRelationshipPartPath(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  const directory = separator === -1 ? "" : partPath.slice(0, separator + 1);
  const filename = partPath.slice(separator + 1);
  return `${directory}_rels/${filename}.rels`;
}

function resolveTemplateRelationshipTarget(
  sourcePath: string,
  target: string | undefined
): string | null {
  const normalizedTarget = target?.trim().split("#", 1)[0] ?? "";
  if (!normalizedTarget || isExternalRelationshipTarget(normalizedTarget)) {
    return null;
  }
  const separator = sourcePath.lastIndexOf("/");
  const directory = separator === -1 ? "" : sourcePath.slice(0, separator + 1);
  const segments =
    `${normalizedTarget.startsWith("/") ? "" : directory}${normalizedTarget.replace(/^\/+/u, "")}`.split(
      "/"
    );
  const normalizedSegments: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      if (normalizedSegments.length === 0) {
        return null;
      }
      normalizedSegments.pop();
      continue;
    }
    normalizedSegments.push(segment);
  }
  return normalizedSegments.join("/");
}

function reachableTemplateParts(
  archive: Record<string, Uint8Array>,
  startPath: string
): Set<string> {
  const reachable = new Set<string>([startPath]);
  const queue = [startPath];
  while (queue.length > 0) {
    const sourcePath = queue.shift();
    if (!sourcePath) {
      continue;
    }
    const relationshipPath = templateRelationshipPartPath(sourcePath);
    const relationshipBytes = archive[relationshipPath];
    if (!relationshipBytes) {
      continue;
    }
    parseTemplateXml(templateArchiveText(archive, relationshipPath), {
      open: (element) => {
        if (
          element.local !== "Relationship" ||
          element.uri !== templatePackageRelationshipNamespace ||
          templateAttribute(element, "TargetMode") !== undefined
        ) {
          return;
        }
        const targetPath = resolveTemplateRelationshipTarget(
          sourcePath,
          templateAttribute(element, "Target")
        );
        if (
          targetPath &&
          Object.hasOwn(archive, targetPath) &&
          !reachable.has(targetPath)
        ) {
          reachable.add(targetPath);
          queue.push(targetPath);
        }
      },
    });
  }
  return reachable;
}

function reachableTemplateControlParts(
  archive: Record<string, Uint8Array>,
  xmlPaths: Set<string>
): Set<string> {
  const reachable = reachableTemplateParts(archive, "word/document.xml");
  return new Set(
    [...xmlPaths].filter(
      (archivePath) =>
        reachable.has(archivePath) &&
        templateControlPartPattern.test(archivePath)
    )
  );
}

function validateTemplatePackageArchive(
  archive: Record<string, Uint8Array>
): Set<string> {
  const contentTypes = templateArchiveText(archive, "[Content_Types].xml");
  const relationships = templateArchiveText(archive, "_rels/.rels");
  const document = templateArchiveText(archive, "word/document.xml");
  const defaultContentTypes = new Map<string, string>();
  const overrideContentTypes = new Map<string, string>();
  let contentTypesRoot = false;
  let contentTypesRootSeen = false;
  parseTemplateXml(contentTypes, {
    open: (element) => {
      if (!contentTypesRootSeen) {
        contentTypesRootSeen = true;
        contentTypesRoot =
          element.local === "Types" &&
          element.uri === templatePackageContentTypesNamespace;
      }
      if (element.uri !== templatePackageContentTypesNamespace) {
        return;
      }
      const contentType = templateAttribute(element, "ContentType");
      if (element.local === "Default") {
        const extension = templateAttribute(
          element,
          "Extension"
        )?.toLowerCase();
        if (!extension || !contentType || defaultContentTypes.has(extension)) {
          fail(422, "invalid_template", "The DOCX content types are invalid");
        }
        defaultContentTypes.set(extension, contentType);
      } else if (element.local === "Override") {
        const partName = templateAttribute(element, "PartName");
        if (
          !partName?.startsWith("/") ||
          !contentType ||
          overrideContentTypes.has(partName)
        ) {
          fail(422, "invalid_template", "The DOCX content types are invalid");
        }
        overrideContentTypes.set(partName, contentType);
      }
    },
  });
  const documentOverride =
    overrideContentTypes.get("/word/document.xml") ===
    templateDocumentContentType;

  let relationshipsRoot = false;
  let relationshipsRootSeen = false;
  let officeDocumentRelationship = false;
  parseTemplateXml(relationships, {
    open: (element) => {
      if (!relationshipsRootSeen) {
        relationshipsRootSeen = true;
        relationshipsRoot =
          element.local === "Relationships" &&
          element.uri === templatePackageRelationshipNamespace;
      }
      const target = templateAttribute(element, "Target");
      if (
        element.local === "Relationship" &&
        element.uri === templatePackageRelationshipNamespace &&
        templateOfficeDocumentRelationships.has(
          templateAttribute(element, "Type") ?? ""
        ) &&
        templateAttribute(element, "TargetMode") === undefined &&
        (target === "word/document.xml" || target === "/word/document.xml")
      ) {
        officeDocumentRelationship = true;
      }
    },
  });

  let documentBody = false;
  let documentRoot = false;
  let documentRootSeen = false;
  parseTemplateXml(document, {
    open: (element) => {
      if (!documentRootSeen) {
        documentRootSeen = true;
        documentRoot =
          element.local === "document" &&
          templateWordNamespaces.has(element.uri);
      }
      if (element.local === "body" && templateWordNamespaces.has(element.uri)) {
        documentBody = true;
      }
    },
  });

  if (
    !contentTypesRoot ||
    !documentOverride ||
    !relationshipsRoot ||
    !officeDocumentRelationship ||
    !documentRoot ||
    !documentBody
  ) {
    fail(
      422,
      "invalid_template",
      "The DOCX package is missing required document parts"
    );
  }

  const xmlPaths = new Set([
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
  ]);
  for (const archivePath of Object.keys(archive)) {
    const lowercasePath = archivePath.toLowerCase();
    const extension = lowercasePath.slice(lowercasePath.lastIndexOf(".") + 1);
    const contentType =
      overrideContentTypes.get(`/${archivePath}`) ??
      defaultContentTypes.get(extension);
    if (
      lowercasePath.endsWith(".xml") ||
      lowercasePath.endsWith(".rels") ||
      lowercasePath.endsWith(".vml") ||
      isTemplateXmlContentType(contentType)
    ) {
      xmlPaths.add(archivePath);
    }
  }
  for (const archivePath of xmlPaths) {
    if (
      archivePath !== "[Content_Types].xml" &&
      archivePath !== "_rels/.rels" &&
      archivePath !== "word/document.xml"
    ) {
      parseTemplateXml(templateArchiveText(archive, archivePath));
    }
  }
  return xmlPaths;
}

function safeTemplateArchive(bytes: Uint8Array): {
  archive: Record<string, Uint8Array>;
  xmlPaths: Set<string>;
} {
  const archive = readTemplateArchive(bytes, () => true);
  validateOfficeRelationships(archive);
  return {
    archive,
    xmlPaths: validateTemplatePackageArchive(archive),
  };
}

function validateTemplatePackage(bytes: Uint8Array): void {
  safeTemplateArchive(bytes);
}

interface ParsedTemplateField {
  label: string;
  options: { displayText: string; value: string }[] | null;
  pictureMaxBytes: number | null;
  pictureMaxHeight: number | null;
  pictureMaxWidth: number | null;
  placeholder: string | null;
  tag: string;
  type: FieldType;
}

interface TemplateControlFrame {
  alias: string | null;
  documentOrder: number;
  duplicateMarker: boolean;
  duplicateTag: boolean;
  inContentDepth: number;
  inPropertiesDepth: number;
  markers: Set<FieldType>;
  options: { displayText: string; value: string }[];
  placeholderName: string | null;
  placeholderText: string;
  propertyStack: string[];
  showingPlaceholder: boolean;
  tag: string | null;
  unsupported: string | null;
}

const templateControlTypeMarkers = new Map<string, FieldType>([
  [templateControlKey(templateWordMainNamespace, "comboBox"), FieldType.combo],
  [templateControlKey(templateWordMainNamespace, "date"), FieldType.date],
  [
    templateControlKey(templateWordMainNamespace, "dropDownList"),
    FieldType.dropdown,
  ],
  [templateControlKey(templateWordMainNamespace, "picture"), FieldType.picture],
  [templateControlKey(templateWordMainNamespace, "text"), FieldType.text],
  [
    templateControlKey(templateWordStrictNamespace, "comboBox"),
    FieldType.combo,
  ],
  [templateControlKey(templateWordStrictNamespace, "date"), FieldType.date],
  [
    templateControlKey(templateWordStrictNamespace, "dropDownList"),
    FieldType.dropdown,
  ],
  [
    templateControlKey(templateWordStrictNamespace, "picture"),
    FieldType.picture,
  ],
  [templateControlKey(templateWordStrictNamespace, "text"), FieldType.text],
  [
    templateControlKey(templateCheckboxNamespace, "checkbox"),
    FieldType.checkbox,
  ],
]);
const templateUnsupportedControlMarkers = new Set([
  templateControlKey(templateWordMainNamespace, "citation"),
  templateControlKey(templateWordMainNamespace, "docPartGallery"),
  templateControlKey(templateWordMainNamespace, "docPartList"),
  templateControlKey(templateWordMainNamespace, "docPartObj"),
  templateControlKey(templateWordMainNamespace, "equation"),
  templateControlKey(templateWordMainNamespace, "group"),
  templateControlKey(templateWordMainNamespace, "richText"),
  templateControlKey(templateWord2012Namespace, "repeatingSection"),
  templateControlKey(templateWord2012Namespace, "repeatingSectionItem"),
]);
const templateControlTypeLocals = new Set([
  "checkbox",
  "comboBox",
  "date",
  "dropDownList",
  "picture",
  "text",
]);
const templateUnsupportedControlLocals = new Set([
  "citation",
  "docPartGallery",
  "docPartList",
  "docPartObj",
  "equation",
  "group",
  "richText",
  "repeatingSection",
  "repeatingSectionItem",
]);
const templateControlPartPattern =
  /^word\/(?:document|endnotes|footnotes|footer\d+|header\d+)\.xml$/u;
const templateControlMetadataProperties = new Set([
  "alias",
  "appearance",
  "calendar",
  "checked",
  "checkedState",
  "color",
  "dataBinding",
  "dateFormat",
  "formPr",
  "id",
  "lock",
  "placeholder",
  "lid",
  "rPr",
  "showingPlcHdr",
  "tag",
  "temporary",
  "uncheckedState",
]);

function templateControlAttribute(
  element: TemplateXmlElement,
  local: string
): string | undefined {
  return element.attributes.find(
    (attribute) => attribute.local === local && attribute.uri === element.uri
  )?.value;
}

function parsedTemplateField(
  frame: TemplateControlFrame,
  glossaryPlaceholders: ReadonlyMap<string, string>
): ParsedTemplateField {
  if (!frame.tag?.trim()) {
    fail(422, "invalid_template", "Every content control must have a tag");
  }
  if (frame.unsupported) {
    fail(
      422,
      "invalid_template",
      `Unsupported content control type: ${frame.unsupported}`
    );
  }
  if (frame.duplicateTag) {
    fail(
      422,
      "invalid_template",
      "A content control cannot repeat its tag property"
    );
  }
  if (frame.duplicateMarker) {
    fail(
      422,
      "invalid_template",
      "A content control cannot repeat a field type marker"
    );
  }
  if (frame.markers.size > 1) {
    fail(
      422,
      "invalid_template",
      "A content control cannot declare multiple field types"
    );
  }
  const type = frame.markers.values().next().value ?? FieldType.text;
  if (
    (type === FieldType.dropdown || type === FieldType.combo) &&
    frame.options.length === 0
  ) {
    fail(
      422,
      "invalid_template",
      "Dropdown and combo fields must define options"
    );
  }
  if (
    type !== FieldType.dropdown &&
    type !== FieldType.combo &&
    frame.options.length > 0
  ) {
    fail(
      422,
      "invalid_template",
      "Only dropdown and combo fields may define options"
    );
  }
  const uniqueOptionLabels = new Set(
    frame.options.map((option) => option.displayText)
  );
  const uniqueOptionValues = new Set(
    frame.options.map((option) => option.value)
  );
  if (
    uniqueOptionLabels.size !== frame.options.length ||
    uniqueOptionValues.size !== frame.options.length
  ) {
    fail(422, "invalid_template", "Dropdown and combo options must be unique");
  }
  const tag = frame.tag.trim();
  const label = frame.alias?.trim();
  const placeholder =
    glossaryPlaceholders.get(frame.placeholderName ?? "")?.trim() ||
    (frame.showingPlaceholder ? frame.placeholderText.trim() : "");
  return {
    label: label || tag,
    options: frame.options.length > 0 ? frame.options : null,
    pictureMaxBytes: type === FieldType.picture ? 10 * 1024 * 1024 : null,
    pictureMaxHeight: type === FieldType.picture ? 4096 : null,
    pictureMaxWidth: type === FieldType.picture ? 4096 : null,
    placeholder: placeholder || null,
    tag,
    type,
  };
}

function templateGlossaryPlaceholders(
  archive: Record<string, Uint8Array>
): Map<string, string> {
  const relationshipsPath = templateRelationshipPartPath("word/document.xml");
  if (!archive[relationshipsPath]) {
    return new Map();
  }
  let glossaryPath: string | null = null;
  parseTemplateXml(templateArchiveText(archive, relationshipsPath), {
    open: (element) => {
      if (
        element.local === "Relationship" &&
        element.uri === templatePackageRelationshipNamespace &&
        templateGlossaryDocumentRelationships.has(
          templateAttribute(element, "Type") ?? ""
        ) &&
        templateAttribute(element, "TargetMode") === undefined
      ) {
        glossaryPath = resolveTemplateRelationshipTarget(
          "word/document.xml",
          templateAttribute(element, "Target")
        );
      }
    },
  });
  if (!glossaryPath || !archive[glossaryPath]) {
    return new Map();
  }
  const placeholders = new Map<string, string>();
  const docParts: {
    bodyDepth: number;
    name: string | null;
    text: string;
  }[] = [];
  const elementStack: string[] = [];
  parseTemplateXml(templateArchiveText(archive, glossaryPath), {
    close: (element) => {
      const docPart = docParts.at(-1);
      if (
        docPart &&
        element.local === "docPartBody" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docPart.bodyDepth -= 1;
      }
      if (
        element.local === "docPart" &&
        templateWordNamespaces.has(element.uri)
      ) {
        const completed = docParts.pop();
        if (
          completed?.name &&
          completed.text.trim() &&
          !placeholders.has(completed.name)
        ) {
          placeholders.set(completed.name, completed.text.trim());
        }
      }
      elementStack.pop();
    },
    open: (element) => {
      if (
        element.local === "docPart" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docParts.push({ bodyDepth: 0, name: null, text: "" });
      }
      const docPart = docParts.at(-1);
      if (!docPart) {
        elementStack.push(element.local);
        return;
      }
      if (
        elementStack.at(-1) === "docPartPr" &&
        element.local === "docPartName" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docPart.name = templateControlAttribute(element, "val") ?? null;
      }
      if (
        element.local === "docPartBody" &&
        templateWordNamespaces.has(element.uri)
      ) {
        docPart.bodyDepth += 1;
      }
      elementStack.push(element.local);
    },
    text: (value) => {
      const docPart = docParts.at(-1);
      if (docPart && docPart.bodyDepth > 0) {
        docPart.text += value;
      }
    },
  });
  return placeholders;
}

function parseTemplateFields(
  bytes: Uint8Array,
  content?: { paragraphs: string[]; fieldText: Map<string, string> }
): ParsedTemplateField[] {
  const { archive, xmlPaths } = safeTemplateArchive(bytes);
  const controlPaths = reachableTemplateControlParts(archive, xmlPaths);
  const fields: { documentOrder: number; field: ParsedTemplateField }[] = [];
  const glossaryPlaceholders = templateGlossaryPlaceholders(archive);
  let documentOrderOffset = 0;
  for (const archivePath of controlPaths) {
    const documentOrderBase = documentOrderOffset;
    let documentOrder = 0;
    let alternateFallbackDepth = 0;
    const controls: TemplateControlFrame[] = [];
    const paragraphs: { text: string }[] = [];
    let textDepth = 0;
    let propertyDepth = 0;
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      close: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth -= 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        if (content && templateWordNamespaces.has(element.uri)) {
          if (element.local === "t") {
            textDepth -= 1;
          }
          if (element.local.endsWith("Pr")) {
            propertyDepth -= 1;
          }
          if (element.local === "p") {
            const text = paragraphs.pop()?.text.trim();
            if (text) {
              content.paragraphs.push(text);
            }
          }
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inContentDepth -= 1;
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 0;
          frame.propertyStack.length = 0;
          return;
        }
        if (frame.inPropertiesDepth > 0) {
          frame.propertyStack.pop();
          frame.inPropertiesDepth -= 1;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.pop();
          if (content && frame.tag) {
            content.fieldText.set(frame.tag.trim(), frame.placeholderText.trim());
          }
          fields.push({
            documentOrder: documentOrderBase + frame.documentOrder,
            field: parsedTemplateField(frame, glossaryPlaceholders),
          });
        }
      },
      open: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth += 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        if (content && templateWordNamespaces.has(element.uri)) {
          if (element.local === "t") {
            textDepth += 1;
          }
          if (element.local.endsWith("Pr")) {
            propertyDepth += 1;
          }
          if (element.local === "p") {
            paragraphs.push({ text: "" });
          }
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          const controlOrder = documentOrderBase + documentOrder;
          documentOrder += 1;
          controls.push({
            alias: null,
            documentOrder: controlOrder,
            duplicateMarker: false,
            duplicateTag: false,
            inContentDepth: 0,
            inPropertiesDepth: 0,
            markers: new Set(),
            options: [],
            placeholderName: null,
            placeholderText: "",
            propertyStack: [],
            showingPlaceholder: false,
            tag: null,
            unsupported: null,
          });
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 1;
          return;
        }
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inContentDepth += 1;
          return;
        }
        if (frame.inPropertiesDepth === 0) {
          return;
        }
        const isControlElement = templateControlNamespaces.has(element.uri);
        const isDirectProperty = frame.inPropertiesDepth === 1;
        const parentProperty = frame.propertyStack.at(-1);
        if (parentProperty === "listItem") {
          fail(422, "invalid_template", "Field options are malformed");
        }
        const controlKey = templateControlKey(element.uri, element.local);
        const marker = templateControlTypeMarkers.get(controlKey);
        const isUnsupportedControl =
          templateUnsupportedControlMarkers.has(controlKey);
        const isTag =
          templateWordNamespaces.has(element.uri) && element.local === "tag";
        const isListItem =
          templateWordNamespaces.has(element.uri) &&
          element.local === "listItem";
        const isAllowedMetadata =
          (templateWordNamespaces.has(element.uri) &&
            templateControlMetadataProperties.has(element.local)) ||
          (element.uri === templateWord2012Namespace &&
            element.local === "appearance") ||
          (element.uri === templateCheckboxNamespace &&
            ["checked", "checkedState", "uncheckedState"].includes(
              element.local
            ));
        const isWrongNamespaceSemantic =
          (element.local === "tag" && !isTag) ||
          (templateControlTypeLocals.has(element.local) &&
            marker === undefined) ||
          (templateUnsupportedControlLocals.has(element.local) &&
            !isUnsupportedControl);
        if (
          (isTag ||
            marker !== undefined ||
            isUnsupportedControl ||
            isWrongNamespaceSemantic) &&
          !isDirectProperty
        ) {
          fail(
            422,
            "invalid_template",
            `Misplaced content control property: ${element.local}`
          );
        }
        if (isTag) {
          if (frame.tag !== null) {
            frame.duplicateTag = true;
          }
          frame.tag = templateControlAttribute(element, "val") ?? null;
        }
        if (
          isDirectProperty &&
          templateWordNamespaces.has(element.uri) &&
          element.local === "alias"
        ) {
          frame.alias = templateControlAttribute(element, "val") ?? null;
        }
        if (
          frame.inPropertiesDepth === 2 &&
          parentProperty === "placeholder" &&
          templateWordNamespaces.has(element.uri) &&
          element.local === "docPart"
        ) {
          frame.placeholderName =
            templateControlAttribute(element, "val") ?? null;
        }
        if (
          isDirectProperty &&
          templateWordNamespaces.has(element.uri) &&
          element.local === "showingPlcHdr"
        ) {
          const value = templateControlAttribute(element, "val")
            ?.trim()
            .toLowerCase();
          frame.showingPlaceholder =
            value === undefined || !["0", "false", "off"].includes(value);
        }
        if (marker !== undefined) {
          if (frame.markers.has(marker)) {
            frame.duplicateMarker = true;
          }
          frame.markers.add(marker);
        }
        if (isUnsupportedControl) {
          frame.unsupported = element.local;
        }
        if (element.local === "listItem") {
          if (
            !isListItem ||
            frame.inPropertiesDepth !== 2 ||
            (parentProperty !== "comboBox" && parentProperty !== "dropDownList")
          ) {
            fail(422, "invalid_template", "Field options are malformed");
          }
          if (
            !frame.markers.has(FieldType.dropdown) &&
            !frame.markers.has(FieldType.combo)
          ) {
            fail(
              422,
              "invalid_template",
              "Only dropdown and combo fields may define options"
            );
          }
          const displayText = templateControlAttribute(element, "displayText");
          const value = templateControlAttribute(element, "value");
          if (!displayText?.trim() || value === undefined) {
            fail(422, "invalid_template", "Field options are malformed");
          }
          frame.options.push({
            displayText: displayText.trim(),
            value,
          });
        }
        if (
          isDirectProperty &&
          (!isControlElement ||
            (marker === undefined &&
              !isUnsupportedControl &&
              !isAllowedMetadata))
        ) {
          fail(
            422,
            "invalid_template",
            `Unknown content control type: ${element.local}`
          );
        }
        frame.propertyStack.push(element.local);
        frame.inPropertiesDepth += 1;
      },
      text: (value) => {
        const frame = controls.at(-1);
        if (content) {
          if (alternateFallbackDepth > 0 || textDepth === 0 || propertyDepth > 0) {
            return;
          }
          if (frame && frame.inContentDepth > 0 && frame.inPropertiesDepth === 0) {
            frame.placeholderText += value;
          } else if (controls.length === 0) {
            const paragraph = paragraphs.at(-1);
            if (paragraph) {
              paragraph.text += value;
            }
          }
        } else if (frame?.showingPlaceholder && frame.inContentDepth > 0) {
          frame.placeholderText += value;
        }
      },
    });
    documentOrderOffset += documentOrder;
  }

  if (fields.length === 0) {
    fail(
      422,
      "invalid_template",
      "The template must contain at least one tagged content control"
    );
  }
  const orderedFields = fields
    .toSorted((left, right) => left.documentOrder - right.documentOrder)
    .map(({ field }) => field);
  const duplicates = orderedFields.filter(
    (field, index) =>
      orderedFields.findIndex((candidate) => candidate.tag === field.tag) !==
      index
  );
  if (duplicates.length > 0) {
    fail(
      422,
      "invalid_template",
      `Content control tags must be unique: ${[...new Set(duplicates.map((field) => field.tag))].join(", ")}`
    );
  }
  return orderedFields;
}
function responseTextControlValues(bytes: Uint8Array): Map<string, string> {
  const { archive, xmlPaths } = safeTemplateArchive(bytes);
  const values = new Map<string, string>();
  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    const controls: {
      contentDepth: number;
      inPropertiesDepth: number;
      paragraphCount: number;
      showingPlaceholder: boolean;
      tag: string | null;
      text: string;
    }[] = [];
    let alternateFallbackDepth = 0;
    let textDepth = 0;
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      close: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth -= 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        if (element.local === "t" && templateWordNamespaces.has(element.uri)) {
          textDepth -= 1;
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.contentDepth -= 1;
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 0;
          return;
        }
        if (frame.inPropertiesDepth > 0) {
          frame.inPropertiesDepth -= 1;
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.pop();
          if (frame.tag && !frame.showingPlaceholder) {
            values.set(frame.tag, frame.text);
          }
        }
      },
      open: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth += 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        if (element.local === "t" && templateWordNamespaces.has(element.uri)) {
          textDepth += 1;
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.push({
            contentDepth: 0,
            inPropertiesDepth: 0,
            paragraphCount: 0,
            showingPlaceholder: false,
            tag: null,
            text: "",
          });
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (element.local === "p" && templateWordNamespaces.has(element.uri)) {
          for (const control of controls) {
            if (
              control.contentDepth === 0 ||
              control.inPropertiesDepth > 0 ||
              control.showingPlaceholder
            ) {
              continue;
            }
            if (control.paragraphCount > 0) {
              control.text += "\n";
            }
            control.paragraphCount += 1;
          }
          return;
        }
        if (
          (element.local === "tab" ||
            element.local === "br" ||
            element.local === "cr") &&
          templateWordNamespaces.has(element.uri)
        ) {
          const separator = element.local === "tab" ? "\t" : "\n";
          for (const control of controls) {
            if (
              control.contentDepth > 0 &&
              control.inPropertiesDepth === 0 &&
              !control.showingPlaceholder
            ) {
              control.text += separator;
            }
          }
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 1;
          return;
        }
        if (
          element.local === "sdtContent" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.contentDepth += 1;
          return;
        }
        if (frame.inPropertiesDepth === 0) {
          return;
        }
        if (
          frame.inPropertiesDepth === 1 &&
          element.local === "tag" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.tag = templateControlAttribute(element, "val") ?? null;
        }
        if (
          frame.inPropertiesDepth === 1 &&
          element.local === "showingPlcHdr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          const value = templateControlAttribute(element, "val")
            ?.trim()
            .toLowerCase();
          frame.showingPlaceholder =
            value === undefined || !["0", "false", "off"].includes(value);
        }
        frame.inPropertiesDepth += 1;
      },
      text: (value) => {
        if (textDepth === 0) {
          return;
        }
        for (const frame of controls) {
          if (frame.contentDepth > 0 && !frame.showingPlaceholder) {
            frame.text += value;
          }
        }
      },
    });
  }
  return values;
}

function validateTemplateControls(
  bytes: Uint8Array,
  expected?: Omit<GeneratedTemplate, "document">
): string[] {
  const content = expected
    ? { paragraphs: [] as string[], fieldText: new Map<string, string>() }
    : undefined;
  const fields = parseTemplateFields(bytes, content);
  if (expected && content) {
    const remaining = new Map<string, number>();
    for (const text of [expected.title, expected.description, ...expected.paragraphs]) {
      if (text) {
        remaining.set(text, (remaining.get(text) ?? 0) + 1);
      }
    }
    for (const text of content.paragraphs) {
      const count = remaining.get(text) ?? 0;
      if (count > 0) {
        remaining.set(text, count - 1);
      } else if (!expected.fields.some((field) => text === field.label || text === `${field.label}:`)) {
        throw new Error("Generated DOCX static text does not match its metadata");
      }
    }
    if ([...remaining.values()].some((count) => count > 0)) {
      throw new Error("Generated DOCX static text does not match its metadata");
    }
    if (fields.some((field) => {
      const declared = expected.fields.find((candidate) => candidate.tag === field.tag);
      return !declared ||
        field.type !== FieldType.text ||
        field.label !== declared.label ||
        content.fieldText.get(field.tag) !== declared.placeholder;
    })) {
      throw new Error("Generated DOCX controls do not match its metadata");
    }
  }
  return fields.map((field) => field.tag);
}
interface ResponsePictureManifestField {
  pictureMaxBytes: number | null;
  pictureMaxHeight: number | null;
  pictureMaxWidth: number | null;
  required: boolean;
  tag: string;
  type: FieldType;
}

interface ResponsePictureControlFrame {
  inPropertiesDepth: number;
  picture: boolean;
  relationshipIds: string[];
  showingPlaceholder: boolean;
  tag: string | null;
}

interface ResponsePictureControl {
  relationshipIds: string[];
  tag: string;
}

interface ResponsePictureDimensions {
  format: "jpeg" | "png";
  height: number;
  width: number;
}

interface NativePictureUpload {
  bytes: Uint8Array;
  dimensions: ResponsePictureDimensions;
}
interface NativeResponseUploadedFile {
  arrayBuffer: () => Promise<ArrayBuffer>;
  size: number;
  type: string;
}
type NativeResponseMultipartEntry = string | NativeResponseUploadedFile;

interface NativeResponseRequest {
  input: JsonRecord;
  pictures: Map<string, NativePictureUpload>;
}

async function nativeResponseRequest(
  request: Request,
  pictureFields: readonly ResponsePictureManifestField[]
): Promise<NativeResponseRequest> {
  const contentType = request.headers.get("content-type");
  if (!contentType || !/^multipart\/form-data(?:\s*;|$)/iu.test(contentType)) {
    return {
      input: await readJsonRecord(
        request,
        maxNativePictureMultipartOverheadBytes
      ),
      pictures: new Map(),
    };
  }
  const imageByteLimit = pictureFields.reduce(
    (total, field) => total + (field.pictureMaxBytes ?? 0),
    0
  );
  const maximumBodyBytes = Math.min(
    maxTemplateArchiveExpandedBytes + maxNativePictureMultipartOverheadBytes,
    imageByteLimit + maxNativePictureMultipartOverheadBytes
  );
  const requestBytes = await readRequestBytes(
    request,
    maximumBodyBytes,
    "Multipart form data is required"
  );
  const multipartRequest = new Request(request.url, {
    body: requestBytes,
    headers: { "content-type": contentType },
    method: "POST",
  });
  let entries: IterableIterator<[string, NativeResponseMultipartEntry]>;
  try {
    const multipartFields = await multipartRequest.formData();
    entries = multipartFields.entries();
  } catch {
    fail(400, "invalid_request", "Multipart form data is invalid");
  }
  const fieldsByTag = new Map(pictureFields.map((field) => [field.tag, field]));
  const pictures = new Map<string, NativePictureUpload>();
  let input: JsonRecord | null = null;
  for (const [key, value] of entries) {
    if (key === "payload") {
      if (input || typeof value !== "string") {
        fail(400, "invalid_request", "payload must be provided once");
      }
      try {
        input = asRecord(JSON.parse(value));
      } catch {
        fail(400, "invalid_request", "payload must be a valid JSON object");
      }
      continue;
    }
    if (!key.startsWith("picture:") || typeof value === "string") {
      fail(
        400,
        "invalid_request",
        "Only payload and picture files are accepted"
      );
    }
    const tag = key.slice("picture:".length);
    const field = fieldsByTag.get(tag);
    if (!field || pictures.has(tag)) {
      fail(400, "invalid_request", "Picture field is unknown or duplicated");
    }
    if (value.size > (field.pictureMaxBytes ?? 0)) {
      fail(
        413,
        "payload_too_large",
        `Picture field ${tag} exceeds its byte limit`
      );
    }
    const fileType = value.type.trim().toLowerCase();
    if (fileType && fileType !== "image/jpeg" && fileType !== "image/png") {
      fail(
        415,
        "invalid_file_type",
        `Picture field ${tag} must be JPEG or PNG`
      );
    }
    const bytes = new Uint8Array(await value.arrayBuffer());
    const dimensions = responsePictureImageDimensions(bytes);
    if (
      !dimensions ||
      (fileType && fileType !== `image/${dimensions.format}`)
    ) {
      invalidResponsePicture(tag, "image must be a valid JPEG or PNG");
    }
    validateResponsePictureMediaBytes(tag, bytes, field, dimensions);
    pictures.set(tag, { bytes, dimensions });
  }
  if (!input) {
    fail(400, "invalid_request", "payload is required");
  }
  return { input, pictures };
}

function invalidResponsePicture(tag: string, message: string): never {
  fail(422, "invalid_template", `Invalid picture field ${tag}: ${message}`);
}

function responsePictureRelationshipId(
  element: TemplateXmlElement
): string | undefined {
  const local =
    element.local === "blip" && templateDrawingMlNamespaces.has(element.uri)
      ? "embed"
      : element.local === "imagedata" && element.uri === templateVmlNamespace
        ? "id"
        : null;
  if (!local) {
    return undefined;
  }
  return element.attributes
    .find(
      (attribute) =>
        attribute.local === local &&
        templateOfficeRelationshipNamespaces.has(attribute.uri)
    )
    ?.value.trim();
}

function responsePictureUint16(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] ?? 0) * 0x1_00 + (bytes[offset + 1] ?? 0);
}

function responsePictureUint32(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset] ?? 0) * 0x1_00_00_00 +
    (bytes[offset + 1] ?? 0) * 0x1_00_00 +
    (bytes[offset + 2] ?? 0) * 0x1_00 +
    (bytes[offset + 3] ?? 0)
  );
}

function responsePictureJpegDimensions(
  bytes: Uint8Array
): ResponsePictureDimensions | null {
  if (bytes.byteLength < 2 || bytes[0] !== 0xff || bytes[1] !== 0xd8) {
    return null;
  }
  let offset = 2;
  while (offset < bytes.byteLength) {
    if (bytes[offset] !== 0xff) {
      return null;
    }
    while (bytes[offset] === 0xff) {
      offset += 1;
    }
    const marker = bytes[offset];
    if (marker === undefined || marker === 0) {
      return null;
    }
    offset += 1;
    if (marker === 0xd9) {
      return null;
    }
    if (responsePictureJpegStandaloneMarkers.has(marker)) {
      continue;
    }
    if (offset + 2 > bytes.byteLength) {
      return null;
    }
    const segmentLength = responsePictureUint16(bytes, offset);
    if (segmentLength < 2 || offset + segmentLength > bytes.byteLength) {
      return null;
    }
    if (responsePictureJpegSofMarkers.has(marker)) {
      if (segmentLength < 7) {
        return null;
      }
      return {
        format: "jpeg",
        height: responsePictureUint16(bytes, offset + 3),
        width: responsePictureUint16(bytes, offset + 5),
      };
    }
    if (marker === 0xda) {
      return null;
    }
    offset += segmentLength;
  }
  return null;
}

function responsePictureImageDimensions(
  bytes: Uint8Array
): ResponsePictureDimensions | null {
  const hasPngSignature =
    bytes.byteLength >= responsePicturePngSignature.length &&
    responsePicturePngSignature.every((byte, index) => bytes[index] === byte);
  if (hasPngSignature) {
    if (
      bytes.byteLength < 24 ||
      responsePictureUint32(bytes, 8) !== 13 ||
      bytes[12] !== 0x49 ||
      bytes[13] !== 0x48 ||
      bytes[14] !== 0x44 ||
      bytes[15] !== 0x52
    ) {
      return null;
    }
    return {
      format: "png",
      height: responsePictureUint32(bytes, 20),
      width: responsePictureUint32(bytes, 16),
    };
  }
  return responsePictureJpegDimensions(bytes);
}

function validateResponsePictureMediaBytes(
  tag: string,
  bytes: Uint8Array,
  field: ResponsePictureManifestField,
  dimensions = responsePictureImageDimensions(bytes)
): void {
  if (
    field.pictureMaxBytes !== null &&
    bytes.byteLength > field.pictureMaxBytes
  ) {
    invalidResponsePicture(tag, "image bytes exceed the published limit");
  }
  if (!dimensions) {
    invalidResponsePicture(tag, "image must be a valid JPEG or PNG");
  }
  if (dimensions.width <= 0 || dimensions.height <= 0) {
    invalidResponsePicture(tag, "image dimensions must be positive");
  }
  if (
    field.pictureMaxWidth !== null &&
    dimensions.width > field.pictureMaxWidth
  ) {
    invalidResponsePicture(tag, "image width exceeds the published limit");
  }
  if (
    field.pictureMaxHeight !== null &&
    dimensions.height > field.pictureMaxHeight
  ) {
    invalidResponsePicture(tag, "image height exceeds the published limit");
  }
}

function responsePictureRelationships(
  archive: Record<string, Uint8Array>,
  sourcePath: string
): Map<string, string | null> {
  const relationships = new Map<string, string | null>();
  const relationshipPath = templateRelationshipPartPath(sourcePath);
  if (!archive[relationshipPath]) {
    return relationships;
  }
  parseTemplateXml(templateArchiveText(archive, relationshipPath), {
    open: (element) => {
      if (
        element.local !== "Relationship" ||
        element.uri !== templatePackageRelationshipNamespace
      ) {
        return;
      }
      const id = templateAttribute(element, "Id")?.trim();
      if (!id) {
        return;
      }
      const target = resolveTemplateRelationshipTarget(
        sourcePath,
        templateAttribute(element, "Target")
      );
      relationships.set(id, relationships.has(id) ? null : target);
    },
  });
  return relationships;
}

function validateResponsePictureControls(
  bytes: Uint8Array,
  manifestFields: readonly ResponsePictureManifestField[],
  enforceRequired: boolean
): Set<string> {
  const pictureFields = manifestFields.filter(
    (field) => field.type === FieldType.picture
  );
  const presentTags = new Set<string>();
  if (pictureFields.length === 0) {
    return presentTags;
  }
  const { archive, xmlPaths } = safeTemplateArchive(bytes);
  const fieldsByTag = new Map(pictureFields.map((field) => [field.tag, field]));
  const seenTags = new Set<string>();
  const relationshipsBySource = new Map<string, Map<string, string | null>>();
  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    const controls: ResponsePictureControlFrame[] = [];
    let alternateFallbackDepth = 0;
    const pictureControls: ResponsePictureControl[] = [];
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      close: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth -= 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 0;
          return;
        }
        if (frame.inPropertiesDepth > 0) {
          frame.inPropertiesDepth -= 1;
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.pop();
          const tag = frame.tag?.trim();
          if (frame.picture && tag) {
            pictureControls.push({
              relationshipIds: frame.showingPlaceholder
                ? []
                : frame.relationshipIds,
              tag,
            });
          }
        }
      },
      open: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth += 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.push({
            inPropertiesDepth: 0,
            picture: false,
            relationshipIds: [],
            showingPlaceholder: false,
            tag: null,
          });
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 1;
          return;
        }
        if (frame.inPropertiesDepth > 0) {
          if (
            frame.inPropertiesDepth === 1 &&
            element.local === "tag" &&
            templateWordNamespaces.has(element.uri)
          ) {
            frame.tag = templateControlAttribute(element, "val") ?? null;
          }
          if (
            frame.inPropertiesDepth === 1 &&
            element.local === "showingPlcHdr" &&
            templateWordNamespaces.has(element.uri)
          ) {
            const value = templateControlAttribute(element, "val")
              ?.trim()
              .toLowerCase();
            frame.showingPlaceholder =
              value === undefined || !["0", "false", "off"].includes(value);
          }
          if (
            frame.inPropertiesDepth === 1 &&
            element.local === "picture" &&
            templateWordNamespaces.has(element.uri)
          ) {
            frame.picture = true;
          }
          frame.inPropertiesDepth += 1;
          return;
        }
        const relationshipId = responsePictureRelationshipId(element);
        if (relationshipId === undefined) {
          return;
        }
        for (const activeFrame of controls) {
          if (activeFrame.picture && activeFrame.inPropertiesDepth === 0) {
            activeFrame.relationshipIds.push(relationshipId);
          }
        }
      },
    });
    for (const control of pictureControls) {
      const field = fieldsByTag.get(control.tag);
      if (!field) {
        continue;
      }
      if (seenTags.has(control.tag)) {
        invalidResponsePicture(control.tag, "content control is duplicated");
      }
      seenTags.add(control.tag);
      if (control.relationshipIds.length === 0) {
        if (field.required && enforceRequired) {
          invalidResponsePicture(control.tag, "a required image is missing");
        }
        continue;
      }
      if (control.relationshipIds.length > 1) {
        invalidResponsePicture(
          control.tag,
          "content control references multiple images"
        );
      }
      const relationshipId = control.relationshipIds[0];
      const relationships =
        relationshipsBySource.get(archivePath) ??
        responsePictureRelationships(archive, archivePath);
      relationshipsBySource.set(archivePath, relationships);
      const mediaPath = relationshipId
        ? relationships.get(relationshipId)
        : undefined;
      if (!mediaPath) {
        invalidResponsePicture(
          control.tag,
          "image relationship or target media is missing"
        );
      }
      const media = archive[mediaPath];
      if (!media) {
        invalidResponsePicture(control.tag, "target media is missing");
      }
      validateResponsePictureMediaBytes(control.tag, media, field);
      presentTags.add(control.tag);
    }
  }
  for (const field of pictureFields) {
    if (!seenTags.has(field.tag)) {
      invalidResponsePicture(field.tag, "content control is missing");
    }
  }
  return presentTags;
}
function responsePicturePresence(
  bytes: Uint8Array,
  pictureFields: readonly ResponsePictureManifestField[]
): Record<string, boolean> {
  const presentTags = validateResponsePictureControls(
    bytes,
    pictureFields,
    false
  );
  return Object.fromEntries(
    pictureFields.map(({ tag }) => [tag, presentTags.has(tag)])
  );
}

function appendTemplateXmlChild(
  xml: string,
  rootLocalName: string,
  rootNamespace: string,
  child: string
): string {
  const parser = new SaxesParser({ position: true, xmlns: true });
  const rootNames: string[] = [];
  const rootCloseStarts: number[] = [];
  let depth = 0;
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    if (depth === 0) {
      if (
        rootNames.length > 0 ||
        tag.local !== rootLocalName ||
        tag.uri !== rootNamespace
      ) {
        fail(422, "invalid_template", "The DOCX package XML root is invalid");
      }
      rootNames.push(tag.name);
    }
    if (!tag.isSelfClosing) {
      depth += 1;
    }
  });
  parser.on("closetag", (tag) => {
    if (
      depth === 1 &&
      tag.local === rootLocalName &&
      tag.uri === rootNamespace
    ) {
      rootCloseStarts.push(xml.lastIndexOf("</", parser.position - 1));
    }
    if (!tag.isSelfClosing) {
      depth -= 1;
    }
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  const rootName = rootNames[0];
  const rootCloseStart = rootCloseStarts[0];
  if (!rootName || rootCloseStart === undefined) {
    fail(422, "invalid_template", "The DOCX package XML root is invalid");
  }
  const prefix = rootName.includes(":")
    ? `${rootName.slice(0, rootName.indexOf(":"))}:`
    : "";
  return `${xml.slice(0, rootCloseStart)}${child.replaceAll(
    "{prefix}",
    prefix
  )}${xml.slice(rootCloseStart)}`;
}

function setNativePictureContentType(
  archive: Record<string, Uint8Array>,
  mediaPath: string,
  contentType: string
): void {
  const xml = templateArchiveText(archive, "[Content_Types].xml");
  const parser = new SaxesParser({ position: true, xmlns: true });
  const overrides: {
    contentType: string | undefined;
    end: number;
    start: number;
  }[] = [];
  parser.on("opentag", (tag) => {
    if (
      tag.local !== "Override" ||
      tag.uri !== templatePackageContentTypesNamespace ||
      templateAttribute(templateXmlElement(tag), "PartName") !== `/${mediaPath}`
    ) {
      return;
    }
    if (overrides.length > 0) {
      fail(422, "invalid_template", "The DOCX content types are invalid");
    }
    overrides.push({
      contentType: templateAttribute(templateXmlElement(tag), "ContentType"),
      end: parser.position,
      start: xml.lastIndexOf("<", parser.position - 1),
    });
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  const override = overrides[0];
  if (override) {
    if (override.contentType === contentType) {
      return;
    }
    const start = override.start;
    const end = override.end;
    const openingTag = xml.slice(start, end);
    const updatedTag = openingTag.replace(
      /(?<prefix>\sContentType\s*=\s*)(?<quote>["'])[^"']*\k<quote>/iu,
      `$<prefix>"${contentType}"`
    );
    if (updatedTag === openingTag) {
      fail(422, "invalid_template", "The DOCX content types are invalid");
    }
    archive["[Content_Types].xml"] = strToU8(
      `${xml.slice(0, start)}${updatedTag}${xml.slice(end)}`.replace(
        /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
        'encoding="UTF-8"'
      )
    );
    return;
  }
  const updatedXml = appendTemplateXmlChild(
    xml,
    "Types",
    templatePackageContentTypesNamespace,
    `<{prefix}Override PartName="/${mediaPath}" ContentType="${contentType}"/>`
  );
  archive["[Content_Types].xml"] = strToU8(
    updatedXml.replace(
      /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
      'encoding="UTF-8"'
    )
  );
}

function pictureRelationshipReferenceCount(
  xml: string,
  relationshipId: string
): number {
  let count = 0;
  parseTemplateXml(xml, {
    open: (element) => {
      if (responsePictureRelationshipId(element) === relationshipId) {
        count += 1;
      }
    },
  });
  return count;
}

function addNativePicturePackageParts(
  archive: Record<string, Uint8Array>,
  sourcePath: string,
  upload: NativePictureUpload,
  existingRelationshipIds: readonly string[]
): string {
  const existingRelationships = responsePictureRelationships(
    archive,
    sourcePath
  );
  const existingRelationshipId =
    existingRelationshipIds.length === 1 ? existingRelationshipIds[0] : null;
  const existingMediaPath = existingRelationshipId
    ? existingRelationships.get(existingRelationshipId)
    : null;
  if (
    existingRelationshipId &&
    existingMediaPath?.startsWith("word/media/picture-") &&
    archive[existingMediaPath] &&
    [...existingRelationships.values()].filter(
      (target) => target === existingMediaPath
    ).length === 1 &&
    pictureRelationshipReferenceCount(
      templateArchiveText(archive, sourcePath),
      existingRelationshipId
    ) === 1
  ) {
    archive[existingMediaPath] = upload.bytes;
    setNativePictureContentType(
      archive,
      existingMediaPath,
      upload.dimensions.format === "jpeg" ? "image/jpeg" : "image/png"
    );
    return existingRelationshipId;
  }

  const extension = upload.dimensions.format === "jpeg" ? "jpeg" : "png";
  const mediaName = `picture-${crypto.randomUUID()}.${extension}`;
  const mediaPath = `word/media/${mediaName}`;
  const relationshipPath = templateRelationshipPartPath(sourcePath);
  const relationshipId = (() => {
    let suffix = 1;
    while (existingRelationships.has(`rId${suffix}`)) {
      suffix += 1;
    }
    return `rId${suffix}`;
  })();
  const target = path.posix.relative(path.posix.dirname(sourcePath), mediaPath);
  const relationship = `<{prefix}Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`;
  const relationshipsXml = archive[relationshipPath]
    ? appendTemplateXmlChild(
        templateArchiveText(archive, relationshipPath),
        "Relationships",
        templatePackageRelationshipNamespace,
        relationship
      )
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${templatePackageRelationshipNamespace}">${relationship.replaceAll("{prefix}", "")}</Relationships>`;
  archive[relationshipPath] = strToU8(
    relationshipsXml.replace(
      /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
      'encoding="UTF-8"'
    )
  );
  setNativePictureContentType(
    archive,
    mediaPath,
    upload.dimensions.format === "jpeg" ? "image/jpeg" : "image/png"
  );
  archive[mediaPath] = upload.bytes;
  return relationshipId;
}

function nativePictureDrawing(
  tag: string,
  block: boolean,
  wordPrefix: string,
  contentNamespace: string,
  relationshipId: string,
  drawingId: number,
  dimensions: ResponsePictureDimensions
): string {
  const strict = contentNamespace === templateWordStrictNamespace;
  const drawingNamespace = strict
    ? "http://purl.oclc.org/ooxml/drawingml/main"
    : "http://schemas.openxmlformats.org/drawingml/2006/main";
  const wordDrawingNamespace = strict
    ? "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
    : "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const pictureNamespace = strict
    ? "http://purl.oclc.org/ooxml/drawingml/picture"
    : "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const relationshipNamespace = strict
    ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
    : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const cx = dimensions.width * 9525;
  const cy = dimensions.height * 9525;
  const drawing =
    `<${wordPrefix}drawing xmlns:a="${drawingNamespace}" xmlns:pic="${pictureNamespace}" xmlns:wp="${wordDrawingNamespace}" xmlns:r="${relationshipNamespace}">` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${drawingId}" name="Picture ${htmlEscape(tag)}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="${pictureNamespace}"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${drawingId}" name="Picture ${htmlEscape(tag)}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></${wordPrefix}drawing>`;
  const run = `<${wordPrefix}r>${drawing}</${wordPrefix}r>`;
  return block ? `<${wordPrefix}p>${run}</${wordPrefix}p>` : run;
}

function overlayNativeResponsePictures(
  documentBytes: Uint8Array,
  manifestFields: readonly ResponsePictureManifestField[],
  pictures: ReadonlyMap<string, NativePictureUpload>
): Uint8Array {
  if (pictures.size === 0) {
    return documentBytes;
  }
  const pictureFields = manifestFields.filter(
    ({ type }) => type === FieldType.picture
  );
  const fieldsByTag = new Map(pictureFields.map((field) => [field.tag, field]));
  const selectedTags = new Set(pictures.keys());
  for (const tag of selectedTags) {
    if (!fieldsByTag.has(tag)) {
      invalidResponsePicture(tag, "field is not a published Picture");
    }
  }
  const { archive, xmlPaths } = safeTemplateArchive(documentBytes);
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  const fieldTags = new Set(manifestFields.map(({ tag }) => tag));
  const drawingIds = new Set<number>();
  for (const archivePath of xmlPaths) {
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      open: (element) => {
        if (element.local !== "docPr") {
          return;
        }
        const id = Number(templateAttribute(element, "id"));
        if (Number.isSafeInteger(id) && id > 0) {
          drawingIds.add(id);
        }
      },
    });
  }
  const placedTags = new Set<string>();
  for (const archivePath of controlParts) {
    const originalXml = templateArchiveText(archive, archivePath);
    const flattened = nativeFlattenAlternateFieldsXml(originalXml, fieldTags);
    if (!flattened.supported) {
      fail(422, "invalid_template", "Native picture control is unsupported");
    }
    const xml = flattened.xml;
    const controls: {
      block: boolean;
      contentEnd: number | null;
      contentName: string | null;
      contentNamespace: string;
      contentPrefix: string;
      contentSelfClosing: boolean;
      contentStart: number | null;
      contentTokenStart: number | null;
      inPropertiesDepth: number;
      picture: boolean;
      relationshipIds: string[];
      showingPlaceholderRanges: { end: number; start: number }[];
      showingPlaceholderStarts: number[];
      tag: string | null;
    }[] = [];
    const elements: { local: string; uri: string }[] = [];
    const patches: { end: number; replacement: string; start: number }[] = [];
    const parser = new SaxesParser({ position: true, xmlns: true });
    parser.on("doctype", () => {
      fail(422, "invalid_template", "DOCX XML document types are not allowed");
    });
    parser.on("opentag", (tag) => {
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        controls.push({
          block: !elements.some(
            (element) =>
              element.local === "p" && templateWordNamespaces.has(element.uri)
          ),
          contentEnd: null,
          contentName: null,
          contentNamespace: "",
          contentPrefix: "",
          contentSelfClosing: false,
          contentStart: null,
          contentTokenStart: null,
          inPropertiesDepth: 0,
          picture: false,
          relationshipIds: [],
          showingPlaceholderRanges: [],
          showingPlaceholderStarts: [],
          tag: null,
        });
      } else {
        const frame = controls.at(-1);
        if (frame) {
          if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
            frame.inPropertiesDepth = 1;
          } else if (frame.inPropertiesDepth > 0) {
            if (
              frame.inPropertiesDepth === 1 &&
              tag.local === "tag" &&
              templateWordNamespaces.has(tag.uri)
            ) {
              frame.tag =
                templateControlAttribute(templateXmlElement(tag), "val") ??
                null;
            }
            if (
              frame.inPropertiesDepth === 1 &&
              tag.local === "picture" &&
              templateWordNamespaces.has(tag.uri)
            ) {
              frame.picture = true;
            }
            if (
              frame.inPropertiesDepth === 1 &&
              tag.local === "showingPlcHdr" &&
              templateWordNamespaces.has(tag.uri)
            ) {
              frame.showingPlaceholderStarts.push(
                xml.lastIndexOf("<", parser.position - 1)
              );
            }
            frame.inPropertiesDepth += 1;
          } else if (
            tag.local === "sdtContent" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            frame.contentName = tag.name;
            frame.contentNamespace = tag.uri;
            frame.contentPrefix = tag.name.includes(":")
              ? tag.name.slice(0, tag.name.indexOf(":") + 1)
              : "";
            frame.contentSelfClosing = tag.isSelfClosing;
            frame.contentStart = tag.isSelfClosing ? null : parser.position;
            frame.contentTokenStart = tag.isSelfClosing ? start : null;
          } else if (frame.picture) {
            const relationshipId = responsePictureRelationshipId(
              templateXmlElement(tag)
            );
            if (relationshipId) {
              frame.relationshipIds.push(relationshipId);
            }
          }
        }
      }
      elements.push({ local: tag.local, uri: tag.uri });
    });
    parser.on("closetag", (tag) => {
      const frame = controls.at(-1);
      if (frame) {
        if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
          frame.inPropertiesDepth = 0;
        } else if (
          frame.inPropertiesDepth === 2 &&
          tag.local === "showingPlcHdr" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          const start = frame.showingPlaceholderStarts.pop();
          if (start !== undefined) {
            frame.showingPlaceholderRanges.push({
              end: parser.position,
              start,
            });
          }
          frame.inPropertiesDepth -= 1;
        } else if (frame.inPropertiesDepth > 0) {
          frame.inPropertiesDepth -= 1;
        } else if (
          tag.local === "sdtContent" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          frame.contentEnd = frame.contentSelfClosing
            ? parser.position
            : xml.lastIndexOf("</", parser.position - 1);
          const fieldTag = frame.tag?.trim();
          const upload = fieldTag ? pictures.get(fieldTag) : undefined;
          if (frame.picture && fieldTag && upload) {
            if (placedTags.has(fieldTag)) {
              invalidResponsePicture(fieldTag, "content control is duplicated");
            }
            placedTags.add(fieldTag);
            for (const range of frame.showingPlaceholderRanges) {
              patches.push({ ...range, replacement: "" });
            }
            let drawingId = 1;
            while (drawingIds.has(drawingId)) {
              drawingId += 1;
            }
            drawingIds.add(drawingId);
            const relationshipId = addNativePicturePackageParts(
              archive,
              archivePath,
              upload,
              frame.relationshipIds
            );
            if (!frame.contentName) {
              invalidResponsePicture(fieldTag, "content control is incomplete");
            }
            const content = nativePictureDrawing(
              fieldTag,
              frame.block,
              frame.contentPrefix,
              frame.contentNamespace,
              relationshipId,
              drawingId,
              upload.dimensions
            );
            if (
              frame.contentSelfClosing &&
              frame.contentTokenStart !== null &&
              frame.contentEnd !== null
            ) {
              const openingTag = xml.slice(
                frame.contentTokenStart,
                frame.contentEnd
              );
              patches.push({
                end: frame.contentEnd,
                replacement: `${openingTag.replace(/\/\s*>$/u, ">")}${content}</${frame.contentName}>`,
                start: frame.contentTokenStart,
              });
            } else if (
              frame.contentStart !== null &&
              frame.contentEnd !== null
            ) {
              patches.push({
                end: frame.contentEnd,
                replacement: content,
                start: frame.contentStart,
              });
            } else {
              invalidResponsePicture(fieldTag, "content control is incomplete");
            }
          }
        } else if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
          controls.pop();
        }
      }
      elements.pop();
    });
    try {
      parser.write(xml).close();
    } catch {
      fail(422, "invalid_template", "The DOCX package contains invalid XML");
    }
    let updatedXml = xml;
    for (const patch of patches.toSorted(
      (left, right) => right.start - left.start
    )) {
      updatedXml =
        updatedXml.slice(0, patch.start) +
        patch.replacement +
        updatedXml.slice(patch.end);
    }
    if (updatedXml !== originalXml) {
      archive[archivePath] = strToU8(
        updatedXml.replace(
          /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
          'encoding="UTF-8"'
        )
      );
    }
  }
  for (const tag of selectedTags) {
    if (!placedTags.has(tag)) {
      invalidResponsePicture(tag, "content control is missing");
    }
  }
  return zipSync(archive);
}

function manifestFieldValueMatches(
  field: { options: unknown; type: FieldType },
  value: unknown
): boolean {
  if (field.type === FieldType.checkbox) {
    return typeof value === "boolean";
  }
  if (field.type === FieldType.dropdown) {
    const optionValues = Array.isArray(field.options)
      ? field.options.flatMap((option) => {
          if (!option || typeof option !== "object" || Array.isArray(option)) {
            return [];
          }
          const optionValue = (option as Record<string, unknown>).value;
          return typeof optionValue === "string" ? [optionValue] : [];
        })
      : [];
    return typeof value === "string" && optionValues.includes(value);
  }
  if (field.type === FieldType.date) {
    return typeof value === "string" && isValidDateFieldValue(value);
  }
  if (field.type === FieldType.combo || field.type === FieldType.text) {
    return typeof value === "string" && value.length <= maxResponseTextLength;
  }
  return false;
}

function validatePrefillValuesAgainstManifest(
  values: JsonRecord,
  fields: readonly { options: unknown; tag: string; type: FieldType }[]
): void {
  const serialized = JSON.stringify(values);
  if (new TextEncoder().encode(serialized).byteLength > maxResponseDataBytes) {
    handoffUnavailable();
  }
  const fieldsByTag = new Map(fields.map((field) => [field.tag, field]));
  for (const [tag, value] of Object.entries(values)) {
    const field = fieldsByTag.get(tag);
    if (!field || !manifestFieldValueMatches(field, value)) {
      handoffUnavailable();
    }
  }
}

interface OnlyOfficeFieldDisplayMetadata {
  calendar: string | null;
  checkedState: NativeCheckboxState | null;
  dateFormat: string | null;
  dateLanguage: string | null;
  uncheckedState: NativeCheckboxState | null;
}
function onlyOfficeFieldDisplayMetadata(
  documentBytes: Uint8Array
): Map<string, OnlyOfficeFieldDisplayMetadata> {
  const { archive, xmlPaths } = safeTemplateArchive(documentBytes);
  const metadataByTag = new Map<string, OnlyOfficeFieldDisplayMetadata>();
  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    const controls: {
      dateFormat: string | null;
      dateLanguage: string | null;
      calendar: string | null;
      checkedState: NativeCheckboxState | null;
      inPropertiesDepth: number;
      tag: string | null;
      uncheckedState: NativeCheckboxState | null;
    }[] = [];
    let alternateFallbackDepth = 0;
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      close: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth -= 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 0;
        } else if (frame.inPropertiesDepth > 0) {
          frame.inPropertiesDepth -= 1;
        } else if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.pop();
          if (frame.tag) {
            metadataByTag.set(frame.tag, {
              calendar: frame.calendar,
              checkedState: frame.checkedState,
              dateFormat: frame.dateFormat,
              dateLanguage: frame.dateLanguage,
              uncheckedState: frame.uncheckedState,
            });
          }
        }
      },
      open: (element) => {
        if (
          element.local === "Fallback" &&
          element.uri === templateMarkupCompatibilityNamespace
        ) {
          alternateFallbackDepth += 1;
          return;
        }
        if (alternateFallbackDepth > 0) {
          return;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controls.push({
            calendar: null,
            checkedState: null,
            dateFormat: null,
            dateLanguage: null,
            inPropertiesDepth: 0,
            tag: null,
            uncheckedState: null,
          });
          return;
        }
        const frame = controls.at(-1);
        if (!frame) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.inPropertiesDepth = 1;
          return;
        }
        if (frame.inPropertiesDepth === 0) {
          return;
        }
        if (
          frame.inPropertiesDepth === 1 &&
          element.local === "tag" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.tag = templateControlAttribute(element, "val")?.trim() ?? null;
        } else if (
          element.local === "checkedState" &&
          element.uri === templateCheckboxNamespace
        ) {
          frame.checkedState = nativeCheckboxState(element);
        } else if (
          element.local === "uncheckedState" &&
          element.uri === templateCheckboxNamespace
        ) {
          frame.uncheckedState = nativeCheckboxState(element);
        } else if (
          element.local === "dateFormat" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.dateFormat = templateControlAttribute(element, "val") ?? null;
        } else if (
          element.local === "lid" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.dateLanguage = templateControlAttribute(element, "val") ?? null;
        } else if (
          element.local === "calendar" &&
          templateWordNamespaces.has(element.uri)
        ) {
          frame.calendar = templateControlAttribute(element, "val") ?? null;
        }
        frame.inPropertiesDepth += 1;
      },
    });
  }
  return metadataByTag;
}
function onlyOfficeFieldTagAliases(
  documentBytes: Uint8Array
): Record<string, string> {
  const { archive, xmlPaths } = safeTemplateArchive(documentBytes);
  const aliases = new Map<string, string>();
  const ambiguousAliases = new Set<string>();
  const fieldTags = new Set<string>();

  for (const archivePath of reachableTemplateControlParts(archive, xmlPaths)) {
    let controlDepth = 0;
    let propertiesDepth = 0;
    parseTemplateXml(templateArchiveText(archive, archivePath), {
      close: (element) => {
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          propertiesDepth = 0;
          return;
        }
        if (propertiesDepth > 0) {
          propertiesDepth -= 1;
        }
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controlDepth -= 1;
        }
      },
      open: (element) => {
        if (
          element.local === "sdt" &&
          templateWordNamespaces.has(element.uri)
        ) {
          controlDepth += 1;
          return;
        }
        if (controlDepth === 0) {
          return;
        }
        if (
          element.local === "sdtPr" &&
          templateWordNamespaces.has(element.uri)
        ) {
          propertiesDepth = 1;
          return;
        }
        if (propertiesDepth === 0) {
          return;
        }
        if (
          propertiesDepth === 1 &&
          element.local === "tag" &&
          templateWordNamespaces.has(element.uri)
        ) {
          const fieldTag = templateControlAttribute(element, "val")?.trim();
          if (fieldTag) {
            fieldTags.add(fieldTag);
            for (const attribute of element.attributes) {
              if (attribute.local !== "val" || attribute.uri === element.uri) {
                continue;
              }
              const alias = attribute.value.trim();
              if (!alias || alias === fieldTag) {
                continue;
              }
              const previous = aliases.get(alias);
              if (previous && previous !== fieldTag) {
                aliases.delete(alias);
                ambiguousAliases.add(alias);
              } else if (!ambiguousAliases.has(alias)) {
                aliases.set(alias, fieldTag);
              }
            }
          }
        }
        propertiesDepth += 1;
      },
    });
  }

  return Object.fromEntries(
    [...aliases].filter(
      ([alias]) => !fieldTags.has(alias) && !ambiguousAliases.has(alias)
    )
  );
}
interface OnlyOfficeDateNames {
  monthsLong: string[];
  monthsShort: string[];
  weekdaysLong: string[];
  weekdaysShort: string[];
}
const onlyOfficeDateNames = new Map<string, OnlyOfficeDateNames>();
function onlyOfficeDateNamesFor(locale: string): OnlyOfficeDateNames | null {
  const cached = onlyOfficeDateNames.get(locale);
  if (cached) {
    return cached;
  }
  try {
    const date = (year: number, month: number, day: number) =>
      new Date(Date.UTC(year, month, day));
    const names = {
      monthsLong: Array.from({ length: 12 }, (_, month) =>
        new Intl.DateTimeFormat(locale, {
          month: "long",
          timeZone: "UTC",
        }).format(date(2024, month, 1))
      ),
      monthsShort: Array.from({ length: 12 }, (_, month) =>
        new Intl.DateTimeFormat(locale, {
          month: "short",
          timeZone: "UTC",
        }).format(date(2024, month, 1))
      ),
      weekdaysLong: Array.from({ length: 7 }, (_, day) =>
        new Intl.DateTimeFormat(locale, {
          timeZone: "UTC",
          weekday: "long",
        }).format(date(2024, 0, 7 + day))
      ),
      weekdaysShort: Array.from({ length: 7 }, (_, day) =>
        new Intl.DateTimeFormat(locale, {
          timeZone: "UTC",
          weekday: "short",
        }).format(date(2024, 0, 7 + day))
      ),
    };
    onlyOfficeDateNames.set(locale, names);
    return names;
  } catch {
    return null;
  }
}
function escapeDateMaskLiteral(value: string): string {
  return value.replaceAll(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}
function onlyOfficeDateDisplayValue(
  value: string,
  metadata: OnlyOfficeFieldDisplayMetadata | undefined
): string | null {
  const format = metadata?.dateFormat;
  const calendar = metadata?.calendar;
  if (
    !format ||
    /(^|[^y])yy([^y]|$)/u.test(format.replaceAll(/'[^']*'|"[^"]*"/gu, "")) ||
    (calendar !== null &&
      calendar !== undefined &&
      calendar !== "gregorian" &&
      calendar !== "thai")
  ) {
    return null;
  }
  const locale = metadata?.dateLanguage?.split(/[-_]/u)[0] || "en";
  const names = onlyOfficeDateNamesFor(locale);
  if (!names) {
    return null;
  }
  const captures: string[] = [];
  const tokenPattern = /'[^']*'|"[^"]*"|dddd|ddd|yyyy|MMMM|MMM|yy|MM|dd|M|d/gu;
  let pattern = "^";
  let offset = 0;
  for (const match of format.matchAll(tokenPattern)) {
    const index = match.index ?? 0;
    pattern += escapeDateMaskLiteral(format.slice(offset, index));
    const token = match[0];
    offset = index + token.length;
    if (token.startsWith("'") || token.startsWith('"')) {
      pattern += escapeDateMaskLiteral(token.slice(1, -1));
      continue;
    }
    if (
      token === "yyyy" ||
      token === "yy" ||
      token === "dd" ||
      token === "d" ||
      token === "MM" ||
      token === "M"
    ) {
      captures.push(token);
      pattern +=
        token === "yyyy"
          ? "(\\d{4})"
          : token === "yy"
            ? "(\\d{2})"
            : token === "dd" || token === "MM"
              ? "(\\d{2})"
              : "(\\d{1,2})";
      continue;
    }
    const tokenNames =
      token === "MMMM"
        ? names.monthsLong
        : token === "MMM"
          ? names.monthsShort
          : token === "dddd"
            ? names.weekdaysLong
            : token === "ddd"
              ? names.weekdaysShort
              : [];
    if (tokenNames.length > 0) {
      captures.push(token);
      pattern += `(${[...tokenNames]
        .sort((left, right) => right.length - left.length)
        .map(escapeDateMaskLiteral)
        .join("|")})`;
    } else {
      pattern += escapeDateMaskLiteral(token);
    }
  }
  pattern += `${escapeDateMaskLiteral(format.slice(offset))}$`;
  const result = new RegExp(pattern, "iu").exec(value);
  if (!result) {
    return null;
  }
  let year: number | null = null;
  let month: number | null = null;
  let day: number | null = null;
  const weekdayValues: { token: string; value: string }[] = [];
  for (const [index, token] of captures.entries()) {
    const capture = result[index + 1];
    if (capture === undefined) {
      return null;
    }
    if (token === "yyyy") {
      year = Number(capture);
    } else if (token === "yy") {
      year = 2000 + Number(capture);
    } else if (token === "MM" || token === "M") {
      month = Number(capture);
    } else if (token === "dd" || token === "d") {
      day = Number(capture);
    } else if (token === "MMMM" || token === "MMM") {
      const monthNames =
        token === "MMMM" ? names.monthsLong : names.monthsShort;
      month =
        monthNames.findIndex(
          (name) => name.toLowerCase() === capture.toLowerCase()
        ) + 1;
    } else {
      weekdayValues.push({ token, value: capture });
    }
  }
  if (year === null || month === null || day === null) {
    return null;
  }
  if (calendar === "thai") {
    year -= 543;
  }
  const isoValue = `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  if (!isValidDateFieldValue(isoValue)) {
    return null;
  }
  if (weekdayValues.length > 0) {
    const weekday = new Date(`${isoValue}T00:00:00Z`).getUTCDay();
    if (
      weekdayValues.some(({ token, value: actual }) => {
        const weekdays =
          token === "dddd" ? names.weekdaysLong : names.weekdaysShort;
        return weekdays[weekday]?.toLowerCase() !== actual.toLowerCase();
      })
    ) {
      return null;
    }
  }
  return isoValue;
}
function onlyOfficeFieldValueNeedsMetadata(
  field: { options: unknown; tag: string; type: FieldType },
  value: unknown
): boolean {
  if (typeof value !== "string") {
    return false;
  }
  if (field.type === FieldType.checkbox) {
    return true;
  }
  if (field.type === FieldType.date) {
    return true;
  }
  if (field.type !== FieldType.dropdown && field.type !== FieldType.combo) {
    return false;
  }
  const options = Array.isArray(field.options)
    ? field.options.flatMap((option) => {
        if (!option || typeof option !== "object" || Array.isArray(option)) {
          return [];
        }
        const record = option as Record<string, unknown>;
        return typeof record.displayText === "string" &&
          typeof record.value === "string"
          ? [{ displayText: record.displayText, value: record.value }]
          : [];
      })
    : [];
  return (
    options.some(
      (option) => option.displayText === value && option.value !== value
    ) || value === ""
  );
}
function normalizeOnlyOfficeDisplayValue(
  field: { options: unknown; tag: string; type: FieldType },
  value: unknown,
  metadata: OnlyOfficeFieldDisplayMetadata | undefined
): unknown {
  if (typeof value !== "string") {
    return value;
  }
  if (field.type === FieldType.checkbox) {
    if (value === "") {
      return false;
    }
    const checkedGlyph = metadata?.checkedState?.glyph ?? "☒";
    const uncheckedGlyph = metadata?.uncheckedState?.glyph ?? "☐";
    if (checkedGlyph !== uncheckedGlyph) {
      if (value === checkedGlyph) {
        return true;
      }
      if (value === uncheckedGlyph) {
        return false;
      }
    }
    return value;
  }
  if (field.type === FieldType.date) {
    if (value === "") {
      return null;
    }
    return onlyOfficeDateDisplayValue(value, metadata) ?? value;
  }
  if (field.type !== FieldType.dropdown && field.type !== FieldType.combo) {
    return value;
  }
  const options = Array.isArray(field.options)
    ? field.options.flatMap((option) => {
        if (!option || typeof option !== "object" || Array.isArray(option)) {
          return [];
        }
        const record = option as Record<string, unknown>;
        return typeof record.displayText === "string" &&
          typeof record.value === "string"
          ? [{ displayText: record.displayText, value: record.value }]
          : [];
      })
    : [];
  if (value === "") {
    return null;
  }
  const selectedOption = options.find((option) => option.displayText === value);
  return selectedOption?.value ?? value;
}
function callbackFieldMetadata(
  data: JsonRecord,
  suppliedData: unknown,
  snapshot: PrefillSnapshot | null
): Pick<OperationMetadata, "callbackEmptyFieldTags" | "serverHeldFieldTags"> {
  const suppliedFields = jsonRecord(suppliedData);
  const lockedFields = jsonRecord(snapshot?.lockedFields);
  const tags = Object.keys(data);
  return {
    callbackEmptyFieldTags: tags.filter((tag) => suppliedFields[tag] === ""),
    serverHeldFieldTags: tags.filter(
      (tag) => lockedFields[tag] === true || !Object.hasOwn(suppliedFields, tag)
    ),
  };
}

async function normalizeResponseData(
  form: FormWithDocuments,
  response: ResponseWithSnapshot,
  inputData: unknown,
  requireRequired = false,
  onlyOfficeCompatibility = false,
  canonicalDateFieldsInput: unknown = undefined
): Promise<JsonRecord> {
  const suppliedData = { ...jsonRecord(inputData) };
  const canonicalDateFields = new Set<string>();
  if (canonicalDateFieldsInput !== undefined) {
    if (
      !Array.isArray(canonicalDateFieldsInput) ||
      !canonicalDateFieldsInput.every(
        (field): field is string => typeof field === "string"
      )
    ) {
      fail(
        400,
        "invalid_request",
        "canonicalDateFields must be an array of field tags"
      );
    }
    for (const field of canonicalDateFieldsInput as string[]) {
      if (canonicalDateFields.has(field)) {
        fail(
          400,
          "invalid_request",
          "canonicalDateFields must not contain duplicates"
        );
      }
      canonicalDateFields.add(field);
    }
  }
  const publishedTemplate = form.publishedTemplate;
  if (!publishedTemplate?.objectKey) {
    fail(409, "not_published", "This form has not been published");
  }
  const manifest = await prisma.fieldManifest.findUnique({
    include: { fields: { orderBy: { tag: "asc" } } },
    where: { publishedTemplateId: publishedTemplate.id },
  });
  if (!manifest) {
    fail(500, "internal_error", "The published Field Manifest is unavailable");
  }
  const fieldsByTag = new Map(
    manifest.fields.map((field) => [field.tag, field])
  );
  for (const field of canonicalDateFields) {
    const definition = fieldsByTag.get(field);
    if (
      !definition ||
      definition.type !== FieldType.date ||
      !Object.hasOwn(suppliedData, field) ||
      typeof suppliedData[field] !== "string" ||
      !isValidDateFieldValue(suppliedData[field] as string)
    ) {
      fail(
        422,
        "invalid_response_data",
        `${field} is not a supplied canonical date field`
      );
    }
  }
  const unknownFields = Object.keys(suppliedData).filter(
    (field) => !fieldsByTag.has(field)
  );
  if (unknownFields.length > 0) {
    fail(
      422,
      "invalid_response_data",
      `Unknown form field(s): ${unknownFields.join(", ")}`
    );
  }
  let onlyOfficeMetadataByTag = new Map<
    string,
    OnlyOfficeFieldDisplayMetadata
  >();
  if (onlyOfficeCompatibility) {
    const fieldsNeedingMetadata = manifest.fields.filter(
      (field) =>
        !canonicalDateFields.has(field.tag) &&
        onlyOfficeFieldValueNeedsMetadata(field, suppliedData[field.tag])
    );
    if (fieldsNeedingMetadata.length > 0) {
      onlyOfficeMetadataByTag = onlyOfficeFieldDisplayMetadata(
        await readObject(publishedTemplate.objectKey)
      );
      for (const field of manifest.fields) {
        if (
          Object.hasOwn(suppliedData, field.tag) &&
          !canonicalDateFields.has(field.tag)
        ) {
          suppliedData[field.tag] = normalizeOnlyOfficeDisplayValue(
            field,
            suppliedData[field.tag],
            onlyOfficeMetadataByTag.get(field.tag)
          );
        }
      }
    }
  }
  let data = onlyOfficeCompatibility
    ? { ...jsonRecord(response.draftData ?? {}), ...suppliedData }
    : { ...suppliedData };
  const snapshot = response.prefillSnapshot;
  if (snapshot) {
    const snapshotData = jsonRecord(snapshot.values);
    const lockedFields = jsonRecord(snapshot.lockedFields);
    for (const [field, value] of Object.entries(snapshotData)) {
      if (lockedFields[field] === true) {
        data[field] = value;
      }
    }
  }
  const pictureTags = new Set(
    manifest.fields
      .filter((field) => field.type === FieldType.picture)
      .map((field) => field.tag)
  );
  data = Object.fromEntries(
    Object.entries(data).filter(([fieldTag]) => !pictureTags.has(fieldTag))
  );
  const serialized = JSON.stringify(data);
  if (new TextEncoder().encode(serialized).byteLength > maxResponseDataBytes) {
    fail(413, "response_too_large", "Response data exceeds the size limit");
  }

  for (const [fieldTag, value] of Object.entries(data)) {
    const field = fieldsByTag.get(fieldTag);
    if (field && value !== null && !manifestFieldValueMatches(field, value)) {
      fail(
        422,
        "invalid_response_data",
        `${fieldTag} does not match the published Field Manifest`
      );
    }
  }
  if (requireRequired) {
    for (const field of manifest.fields) {
      if (field.type === FieldType.picture) {
        continue;
      }
      const value = data[field.tag];
      const blankOptionSelected =
        typeof value === "string" &&
        value.trim().length === 0 &&
        (field.type === FieldType.dropdown || field.type === FieldType.combo) &&
        Array.isArray(field.options) &&
        field.options.some(
          (option) =>
            option !== null &&
            typeof option === "object" &&
            !Array.isArray(option) &&
            (option as Record<string, unknown>).value === value
        );
      if (
        field.required &&
        (!Object.hasOwn(data, field.tag) ||
          value === null ||
          (typeof value === "string" &&
            value.trim().length === 0 &&
            !blankOptionSelected) ||
          (field.type === FieldType.checkbox && value !== true))
      ) {
        fail(
          422,
          "invalid_response_data",
          `${field.tag} is required by the published Field Manifest`
        );
      }
    }
  }
  if (snapshot) {
    const snapshotData = jsonRecord(snapshot.values);
    const lockedFields = jsonRecord(snapshot.lockedFields);
    for (const [field, value] of Object.entries(snapshotData)) {
      if (lockedFields[field] === true && Object.hasOwn(suppliedData, field)) {
        const suppliedValue = suppliedData[field];
        if (suppliedValue !== value) {
          fail(422, "invalid_response_data", `${field} is locked by Prefill`);
        }
      }
    }
  }
  return data;
}
function requireCurrentFillMethod(
  form: FormWithDocuments,
  input: JsonRecord
): void {
  const requestedFillMethod = input.fillMethod ?? FillMethod.onlyoffice;
  if (
    requestedFillMethod !== FillMethod.native &&
    requestedFillMethod !== FillMethod.onlyoffice
  ) {
    fail(400, "invalid_request", "fillMethod must be native or onlyoffice");
  }
  if (requestedFillMethod !== form.fillMethod) {
    fail(409, "fill_method_changed", "The form Fill Method changed");
  }
}
function changedResponseData(
  previous: JsonRecord,
  next: JsonRecord
): JsonRecord {
  const keys = new Set([...Object.keys(previous), ...Object.keys(next)]);
  return Object.fromEntries(
    [...keys]
      .filter((key) => !isDeepStrictEqual(previous[key], next[key]))
      .map((key) => [key, next[key] ?? null])
  );
}

async function activeAfterRecovery(
  operation: Operation | null
): Promise<boolean> {
  if (!operation) {
    return false;
  }
  const current = await expireOperationIfNeeded(operation);
  return (
    current.status === OperationStatus.pending ||
    current.status === OperationStatus.processing
  );
}

async function activeOperationForForm(formId: string): Promise<boolean> {
  return activeAfterRecovery(
    await prisma.operation.findFirst({
      where: {
        formId,
        status: { in: [OperationStatus.pending, OperationStatus.processing] },
        targetType: OperationTargetType.template_draft,
      },
    })
  );
}

async function activeOperationForResponse(
  responseId: string
): Promise<boolean> {
  return activeAfterRecovery(
    await prisma.operation.findFirst({
      where: {
        responseId,
        status: { in: [OperationStatus.pending, OperationStatus.processing] },
      },
    })
  );
}

async function createOperation(input: {
  actorId: string;
  authorization: ActionEditorAuthorization;
  capabilityScope: Omit<EditorCapabilityScope, "action" | "operationId">;
  documentKey: string;
  formId: string;
  expectedFillMethod?: FillMethod;
  metadata: OperationMetadata;
  ownerUserId: string;
  responseId?: string;
  stagingObjectKey: string;
  submissionId?: string;
  targetId: string;
  targetType: OperationTargetType;
  type: OperationType;
}): Promise<Operation> {
  try {
    return await prisma.$transaction(async (tx) => {
      const [lockedForm] = await tx.$queryRaw<
        {
          fillMethod: FillMethod;
          id: string;
        }[]
      >(
        Prisma.sql`
          SELECT "id", "fill_method" AS "fillMethod"
          FROM "forms"
          WHERE "id" = ${input.formId}::uuid
          FOR UPDATE
        `
      );
      if (!lockedForm) {
        fail(404, "not_found", "Form was not found");
      }
      if (
        input.expectedFillMethod !== undefined &&
        lockedForm.fillMethod !== input.expectedFillMethod
      ) {
        fail(409, "fill_method_changed", "The form Fill Method changed");
      }
      await lockActiveEditorLease(
        tx,
        input.authorization,
        input.capabilityScope
      );
      const activeOperation = await tx.operation.findFirst({
        select: { id: true },
        where: {
          status: { in: [OperationStatus.pending, OperationStatus.processing] },
          targetId: input.targetId,
          targetType: input.targetType,
        },
      });
      if (activeOperation) {
        fail(
          409,
          "operation_in_progress",
          "Another document operation is already in progress"
        );
      }
      return tx.operation.create({
        data: {
          actorId: input.actorId,
          documentKey: input.documentKey,
          errorCode: null,
          formId: input.formId,
          metadata: jsonValue(input.metadata),
          ownerUserId: input.ownerUserId,
          responseId: input.responseId,
          stagingObjectKey: input.stagingObjectKey,
          status: OperationStatus.pending,
          submissionId: input.submissionId,
          targetId: input.targetId,
          targetType: input.targetType,
          type: input.type,
        },
      });
    });
  } catch (error) {
    if (databaseErrorCode(error) === "P2002") {
      fail(
        409,
        "operation_in_progress",
        "Another document operation is already in progress"
      );
    }
    throw error;
  }
}

async function updateOperationFailed(
  operationId: string,
  errorCode: OperationErrorCode,
  updatedBefore?: Date
): Promise<boolean> {
  const operation = await prisma.$transaction(async (tx) => {
    const current = await tx.operation.findUnique({
      select: { actorId: true, metadata: true, stagingObjectKey: true },
      where: { id: operationId },
    });
    if (!current) {
      return null;
    }
    const metadata = operationMetadata(current.metadata);
    const failed = await tx.operation.updateMany({
      data: {
        errorCode,
        status: OperationStatus.failed,
        updatedAt: new Date(),
      },
      where: {
        id: operationId,
        ...(updatedBefore ? { updatedAt: { lte: updatedBefore } } : {}),
        status: { in: [OperationStatus.pending, OperationStatus.processing] },
      },
    });
    if (failed.count !== 1) {
      return null;
    }
    if (metadata.action === "save-template" || metadata.action === "publish") {
      let targetId =
        typeof metadata.publicId === "string" &&
        publicIdPattern.test(metadata.publicId)
          ? metadata.publicId
          : null;
      if (!targetId) {
        const targetForm = await tx.form.findUnique({
          select: { publicId: true },
          where: { id: metadata.formId },
        });
        targetId = targetForm?.publicId ?? null;
      }
      await createFormAudit(tx, {
        action:
          metadata.action === "publish"
            ? "publish_form"
            : "save_template_draft",
        actorId: current.actorId,
        outcome: AuditOutcome.failure,
        safeMetadata: { errorCode: formAuditErrorCodeFromCode(errorCode) },
        targetId,
      });
    }
    if (
      metadata.action === "save-correction" &&
      metadata.responseId &&
      current.actorId
    ) {
      await tx.auditEvent.create({
        data: {
          action: "create_correction",
          actorId: current.actorId,
          outcome: AuditOutcome.failure,
          safeMetadata: jsonValue({ errorCode, state: "submitted" }),
          targetId: metadata.responseId,
          targetType: "response",
        },
      });
    }
    if (metadata.action === "submit" && metadata.responseId) {
      await tx.response.updateMany({
        data: { status: ResponseStatus.draft, updatedAt: new Date() },
        where: {
          id: metadata.responseId,
          status: ResponseStatus.submitting,
        },
      });
    }
    return { ...current, metadata };
  });
  if (!operation) {
    return false;
  }
  await deleteObjects([operation.stagingObjectKey]);
  await deleteObjectUnlessCanonical(operation.metadata.finalObjectKey);
  return true;
}

async function expireOperationIfNeeded(
  operation: Operation
): Promise<Operation> {
  const active =
    operation.status === OperationStatus.pending ||
    operation.status === OperationStatus.processing;
  const staleBefore = new Date(Date.now() - operationTimeoutMs);
  const stale = operation.updatedAt <= staleBefore;
  if (!active || !stale) {
    return operation;
  }
  await updateOperationFailed(operation.id, "operation_timeout", staleBefore);
  return (
    (await prisma.operation.findUnique({ where: { id: operation.id } })) ??
    operation
  );
}

type CallbackClaimConsumption = "claimed" | "invalid" | "replayed";

async function consumeCallbackClaim(
  operationId: string,
  userdata: string
): Promise<CallbackClaimConsumption> {
  const digest = tokenDigest(userdata);
  const consumed = await prisma.callbackClaim.updateMany({
    data: { consumedAt: new Date() },
    where: {
      consumedAt: null,
      expiresAt: { gt: new Date() },
      operationId,
      tokenDigest: digest,
    },
  });
  if (consumed.count === 1) {
    return "claimed";
  }
  const existing = await prisma.callbackClaim.findFirst({
    select: { consumedAt: true },
    where: { operationId, tokenDigest: digest },
  });
  return existing?.consumedAt ? "replayed" : "invalid";
}
async function expireDuePrefillHandoffs(now: Date): Promise<void> {
  await prisma.$transaction(async (tx) => {
    const dueHandoffs = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`
        SELECT "id"
        FROM "handoffs"
        WHERE "status" IN ('pending', 'reserved')
          AND "expires_at" <= ${now}
        ORDER BY "expires_at" ASC
        LIMIT ${handoffExpirySweepBatchSize}
      `
    );
    if (dueHandoffs.length === 0) {
      return;
    }
    for (const handoff of dueHandoffs) {
      await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "pending_claims"
          WHERE "handoff_id" = ${handoff.id}::uuid
          FOR UPDATE
        `
      );
      const expired = await tx.handoff.updateMany({
        data: {
          codeDigest: null,
          configurationHash: null,
          filteredValues: Prisma.JsonNull,
          formId: null,
          normalizedEmail: null,
          responseId: null,
          status: HandoffStatus.expired,
        },
        where: {
          expiresAt: { lte: now },
          id: handoff.id,
          status: { in: [HandoffStatus.pending, HandoffStatus.reserved] },
        },
      });
      if (expired.count === 1) {
        await tx.pendingClaim.deleteMany({
          where: { handoffId: handoff.id },
        });
      }
    }
  });
}

export async function reconcileRecoverableState(): Promise<void> {
  await drainObjectCleanupIntents();
  const now = new Date();
  await expireDuePrefillHandoffs(now);
  const staleBefore = new Date(now.getTime() - operationTimeoutMs);
  const staleOperations = await prisma.operation.findMany({
    select: { id: true },
    where: {
      status: { in: [OperationStatus.pending, OperationStatus.processing] },
      updatedAt: { lte: staleBefore },
    },
  });
  for (const operation of staleOperations) {
    await updateOperationFailed(operation.id, "operation_timeout", staleBefore);
  }
  const expiredLeases = await prisma.editorLease.findMany({
    select: {
      id: true,
      targetId: true,
      targetType: true,
    },
    where: {
      OR: [
        { expiresAt: { lte: now } },
        { holderSession: { expiresAt: { lte: now } } },
      ],
    },
  });
  for (const lease of expiredLeases) {
    const workspaceObjectKey = await prisma.$transaction(async (tx) => {
      const [locked] = await tx.$queryRaw<
        {
          expiresAt: Date;
          sessionExpiresAt: Date;
          workspaceObjectKey: string | null;
        }[]
      >(
        Prisma.sql`
          SELECT
            "editor_leases"."expires_at" AS "expiresAt",
            "session"."expires_at" AS "sessionExpiresAt",
            "editor_leases"."workspace_object_key" AS "workspaceObjectKey"
          FROM "editor_leases"
          INNER JOIN "session"
            ON "session"."id" = "editor_leases"."holder_session_id"
          WHERE "editor_leases"."id" = ${lease.id}::uuid
          FOR UPDATE
        `
      );
      if (
        !locked ||
        (locked.expiresAt > now && locked.sessionExpiresAt > now)
      ) {
        return null;
      }
      const activeOperation = await tx.operation.findFirst({
        select: { id: true },
        where: {
          status: { in: [OperationStatus.pending, OperationStatus.processing] },
          targetId: lease.targetId,
          targetType: lease.targetType,
        },
      });
      if (activeOperation) {
        return null;
      }
      const deleted = await tx.editorLease.deleteMany({
        where: { id: lease.id },
      });
      return deleted.count === 1 ? locked.workspaceObjectKey : null;
    });
    if (workspaceObjectKey) {
      await deleteObjectUnlessCanonical(workspaceObjectKey);
    }
  }
  await prisma.callbackClaim.deleteMany({
    where: { expiresAt: { lte: now } },
  });
  await drainObjectCleanupIntents();
}

function launchForceSave(
  operation: Operation,
  onlyOffice: OnlyOfficeClient,
  allowedCallbackOrigins: ReadonlySet<string>
): void {
  void (async () => {
    try {
      if (!operation.documentKey) {
        fail(500, "invalid_operation", "Operation has no document key");
      }
      const expiresAtSeconds =
        Math.floor(Date.now() / 1000) + callbackClaimLifetimeSeconds;
      const userdata = createCallbackUserdata({
        documentKey: operation.documentKey,
        expiresAt: expiresAtSeconds,
        operationId: operation.id,
        operationType: operation.type,
      });
      const claimed = await prisma.$transaction(async (tx) => {
        const activated = await tx.operation.updateMany({
          data: { status: OperationStatus.processing, updatedAt: new Date() },
          where: { id: operation.id, status: OperationStatus.pending },
        });
        if (activated.count !== 1) {
          return false;
        }
        await tx.callbackClaim.create({
          data: {
            expiresAt: new Date(expiresAtSeconds * 1000),
            operationId: operation.id,
            tokenDigest: tokenDigest(userdata),
          },
        });
        return true;
      });
      if (!claimed) {
        return;
      }
      const hasChanges = await onlyOffice.forceSave(
        operation.documentKey,
        userdata
      );
      if (!hasChanges) {
        const consumption = await consumeCallbackClaim(operation.id, userdata);
        if (consumption === "replayed") {
          return;
        }
        if (consumption === "invalid") {
          await updateOperationFailed(operation.id, "callback_claim_invalid");
          return;
        }
        const currentObjectKey = await operationDocumentKey(
          operation.documentKey
        );
        if (!currentObjectKey) {
          fail(
            500,
            "document_unavailable",
            "The current document snapshot is unavailable"
          );
        }
        await finalizeCallback(
          operation.id,
          { key: operation.documentKey, status: 6 },
          await readObject(currentObjectKey),
          allowedCallbackOrigins
        );
      }
    } catch (error) {
      await updateOperationFailed(
        operation.id,
        error instanceof HttpError && error.code === "invalid_template"
          ? "invalid_template"
          : "force_save_failed"
      );
    }
  })();
}

async function operationDocumentKey(
  documentKey: string
): Promise<string | null> {
  const pending = await prisma.operation.findMany({
    orderBy: { updatedAt: "desc" },
    select: { metadata: true },
    take: 10,
    where: {
      documentKey,
      status: { in: [OperationStatus.pending, OperationStatus.processing] },
    },
  });
  for (const operation of pending) {
    const metadata = operationMetadata(operation.metadata);
    if (await objectExists(metadata.stagedObjectKey)) {
      return metadata.stagedObjectKey;
    }
  }

  const templateDraft = await prisma.templateDraft.findUnique({
    where: { documentKey },
  });
  if (templateDraft) {
    return templateDraft.objectKey;
  }
  const publishedTemplate = await prisma.publishedTemplate.findUnique({
    where: { documentKey },
  });
  if (publishedTemplate) {
    return publishedTemplate.objectKey;
  }
  const response = await prisma.response.findUnique({
    select: { draftObjectKey: true },
    where: { draftDocumentKey: documentKey },
  });
  if (response?.draftObjectKey) {
    return response.draftObjectKey;
  }
  const correctionWorkspace = await prisma.editorLease.findUnique({
    select: { workspaceObjectKey: true },
    where: { workspaceDocumentKey: documentKey },
  });
  if (correctionWorkspace?.workspaceObjectKey) {
    return correctionWorkspace.workspaceObjectKey;
  }
  const correction = await prisma.correction.findUnique({
    select: { objectKey: true },
    where: { documentKey },
  });
  if (correction) {
    return correction.objectKey;
  }
  const submission = await prisma.submission.findUnique({
    select: { objectKey: true },
    where: { documentKey },
  });
  return submission?.objectKey ?? null;
}

async function markOperationCompleted(
  tx: Prisma.TransactionClient,
  operationId: string,
  result: JsonRecord,
  metadata: OperationMetadata,
  cleanupObjectKeys: string[],
  submissionId?: string
): Promise<void> {
  const completed = await tx.operation.updateMany({
    data: {
      errorCode: null,
      metadata: jsonValue({ ...metadata, cleanupObjectKeys }),
      result: jsonValue(result),
      status: OperationStatus.completed,
      ...(submissionId ? { submissionId } : {}),
      updatedAt: new Date(),
    },
    where: { id: operationId, status: OperationStatus.processing },
  });
  if (completed.count !== 1) {
    fail(409, "stale_operation", "The document operation is no longer active");
  }
}

async function completeTemplateOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  if (!documentKey) {
    fail(500, "invalid_operation", "Template operation has no document key");
  }
  const form = await prisma.form.findUnique({
    include: { publishedTemplate: true, templateDraft: true },
    where: { id: metadata.formId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  if (form.status === FormStatus.published || form.publishedTemplate) {
    fail(
      409,
      "published_immutable",
      "Published forms cannot be structurally edited"
    );
  }
  const templateDraft = form.templateDraft;
  if (!templateDraft || templateDraft.documentKey !== documentKey) {
    fail(
      409,
      "stale_operation",
      "The template changed while this operation was running"
    );
  }
  const nextDocumentKey = metadata.nextDocumentKey ?? documentKey;
  const result = { documentKey: nextDocumentKey, publicId: form.publicId };
  const cleanupObjectKeys = [metadata.stagedObjectKey, templateDraft.objectKey];
  await prisma.$transaction(
    async (tx) => {
      const updated = await tx.templateDraft.updateMany({
        data: {
          contentHash: contentHash(bytes),
          documentKey: nextDocumentKey,
          objectKey: metadata.finalObjectKey,
          updatedAt: new Date(),
        },
        where: { documentKey, id: templateDraft.id },
      });
      if (updated.count !== 1) {
        fail(
          409,
          "stale_operation",
          "The template changed while this operation was running"
        );
      }
      const formUpdated = await tx.form.updateMany({
        data: { updatedAt: new Date() },
        where: { id: form.id },
      });
      if (formUpdated.count !== 1) {
        fail(
          409,
          "stale_operation",
          "The form changed while this operation was running"
        );
      }
      await markOperationCompleted(
        tx,
        operation.id,
        result,
        metadata,
        cleanupObjectKeys
      );
      await createFormAudit(tx, {
        action: "save_template_draft",
        actorId: operation.actorId,
        outcome: AuditOutcome.success,
        safeMetadata: {},
        targetId: form.publicId,
      });
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}

function publishedContractFields(
  controls: ParsedTemplateField[],
  draftRules: DraftFieldRule[]
): {
  manifestFields: {
    label: string;
    options?: Prisma.InputJsonValue;
    pictureMaxBytes: number | null;
    pictureMaxHeight: number | null;
    pictureMaxWidth: number | null;
    placeholder: string | null;
    position: number;
    prefillPolicy: PrefillPolicy;
    required: boolean;
    tag: string;
    type: FieldType;
  }[];
  prefillFields: {
    pointer: string;
    policy: PrefillPolicy;
    tag: string;
  }[];
} {
  const controlsByTag = new Map(controls.map((field) => [field.tag, field]));
  const rulesByTag = new Map<string, DraftFieldRule>();
  for (const rule of draftRules) {
    if (rulesByTag.has(rule.tag) || !controlsByTag.has(rule.tag)) {
      fail(
        422,
        "invalid_template",
        `Field policy does not match a published content control: ${rule.tag}`
      );
    }
    const control = controlsByTag.get(rule.tag);
    if (control?.type === FieldType.picture && rule.prefillPointer !== null) {
      fail(
        422,
        "invalid_template",
        `Picture fields cannot use prefillPointer: ${rule.tag}`
      );
    }
    validateFieldRulePointer(rule.prefillPointer);
    if (
      rule.prefillPolicy === PrefillPolicy.lock_when_available &&
      rule.prefillPointer === null
    ) {
      fail(
        422,
        "invalid_template",
        `Locked Prefill policy requires a pointer: ${rule.tag}`
      );
    }
    rulesByTag.set(rule.tag, rule);
  }
  return {
    manifestFields: controls.map((field, position) => {
      const rule = rulesByTag.get(field.tag);
      return {
        ...(field.options ? { options: jsonValue(field.options) } : {}),
        label: field.label,
        pictureMaxBytes: field.pictureMaxBytes,
        pictureMaxHeight: field.pictureMaxHeight,
        pictureMaxWidth: field.pictureMaxWidth,
        placeholder: field.placeholder,
        position,
        prefillPolicy: rule?.prefillPolicy ?? PrefillPolicy.editable,
        required: rule?.required ?? false,
        tag: field.tag,
        type: field.type,
      };
    }),
    prefillFields: draftRules
      .filter((rule) => rule.prefillPointer !== null)
      .map((rule) => ({
        pointer: rule.prefillPointer as string,
        policy: rule.prefillPolicy,
        tag: rule.tag,
      })),
  };
}
async function completePublishOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  if (!documentKey) {
    fail(500, "invalid_operation", "Publish operation has no document key");
  }
  const form = await prisma.form.findUnique({
    include: { publishedTemplate: true, templateDraft: true },
    where: { id: metadata.formId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  if (form.status === FormStatus.published || form.publishedTemplate) {
    fail(
      409,
      "published_immutable",
      "A Published Template already exists for this Form"
    );
  }
  if (!form.templateDraft || form.templateDraft.documentKey !== documentKey) {
    fail(409, "stale_operation", "The template changed while publishing");
  }
  const { publishedVersion, publishedKey } = metadata;
  if (
    typeof publishedVersion !== "number" ||
    typeof publishedKey !== "string"
  ) {
    fail(500, "invalid_operation", "Publish metadata is incomplete");
  }
  const controls = parseTemplateFields(bytes);
  const hash = contentHash(bytes);
  const result = {
    documentKey: publishedKey,
    publicId: form.publicId,
    version: publishedVersion,
  };
  const cleanupObjectKeys = [metadata.stagedObjectKey];
  await prisma.$transaction(
    async (tx) => {
      const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "forms"
          WHERE "id" = ${form.id}::uuid
          FOR UPDATE
        `
      );
      if (!lockedForm) {
        fail(404, "not_found", "Form was not found");
      }
      const currentForm = await tx.form.findUnique({
        include: { publishedTemplate: true, templateDraft: true },
        where: { id: form.id },
      });
      if (!currentForm) {
        fail(404, "not_found", "Form was not found");
      }
      if (
        currentForm.status === FormStatus.published ||
        currentForm.publishedTemplate
      ) {
        fail(
          409,
          "published_immutable",
          "A Published Template already exists for this Form"
        );
      }
      if (
        currentForm.version !== form.version ||
        !currentForm.templateDraft ||
        currentForm.templateDraft.documentKey !== documentKey
      ) {
        fail(409, "stale_operation", "The form changed while publishing");
      }
      const draftRules = await tx.draftFieldRule.findMany({
        orderBy: { tag: "asc" },
        where: { templateDraftId: currentForm.templateDraft.id },
      });
      const { manifestFields, prefillFields } = publishedContractFields(
        controls,
        draftRules
      );
      const publishedTemplate = await tx.publishedTemplate.create({
        data: {
          contentHash: hash,
          documentKey: publishedKey,
          form: { connect: { id: currentForm.id } },
          id: crypto.randomUUID(),
          manifest: {
            create: {
              configurationHash: hash,
              fields: { create: manifestFields },
            },
          },
          objectKey: metadata.finalObjectKey,
          version: publishedVersion,
        },
      });
      const prefillConfiguration = await tx.prefillConfiguration.create({
        data: {
          configurationHash: hash,
          formId: currentForm.id,
        },
      });
      if (prefillFields.length > 0) {
        await tx.prefillField.createMany({
          data: prefillFields.map((field) => ({
            ...field,
            configurationId: prefillConfiguration.id,
          })),
        });
      }
      await tx.prefillConfiguration.update({
        data: { publishedTemplateId: publishedTemplate.id },
        where: { id: prefillConfiguration.id },
      });
      const updated = await tx.form.updateMany({
        data: {
          status: FormStatus.published,
          updatedAt: new Date(),
          version: publishedVersion,
        },
        where: {
          id: currentForm.id,
          status: FormStatus.draft,
          version: currentForm.version,
        },
      });
      if (updated.count !== 1) {
        fail(409, "stale_operation", "The form changed while publishing");
      }
      await markOperationCompleted(
        tx,
        operation.id,
        result,
        metadata,
        cleanupObjectKeys
      );
      await createFormAudit(tx, {
        action: "publish_form",
        actorId: operation.actorId,
        outcome: AuditOutcome.success,
        safeMetadata: {},
        targetId: currentForm.publicId,
      });
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}
async function validateResponseDocument(
  publishedTemplateId: string,
  bytes: Uint8Array,
  enforceRequired: boolean
): Promise<void> {
  const fields = parseTemplateFields(bytes);
  const manifest = await prisma.fieldManifest.findUnique({
    include: { fields: { orderBy: { tag: "asc" } } },
    where: { publishedTemplateId },
  });
  if (!manifest || manifest.fields.length !== fields.length) {
    fail(
      422,
      "invalid_template",
      "The response document does not match the published manifest"
    );
  }
  const fieldsByTag = new Map(fields.map((field) => [field.tag, field]));
  for (const manifestField of manifest.fields) {
    const field = fieldsByTag.get(manifestField.tag);
    const fieldOptions = field?.options;
    const manifestOptions = manifestField.options;
    const optionsMatch =
      (fieldOptions === null && manifestOptions === null) ||
      (Array.isArray(fieldOptions) &&
        Array.isArray(manifestOptions) &&
        manifestOptions.length === fieldOptions.length &&
        fieldOptions.every((option, index) => {
          const manifestOption = manifestOptions[index];
          return (
            typeof manifestOption === "object" &&
            manifestOption !== null &&
            !Array.isArray(manifestOption) &&
            manifestOption.displayText === option.displayText &&
            manifestOption.value === option.value
          );
        }));
    if (
      !field ||
      field.type !== manifestField.type ||
      !optionsMatch ||
      field.pictureMaxBytes !== manifestField.pictureMaxBytes ||
      field.pictureMaxHeight !== manifestField.pictureMaxHeight ||
      field.pictureMaxWidth !== manifestField.pictureMaxWidth
    ) {
      fail(
        422,
        "invalid_template",
        "The response document does not match the published manifest"
      );
    }
  }
  validateResponsePictureControls(bytes, manifest.fields, enforceRequired);
}
function templatePartsHaveNestedControls(
  archive: Record<string, Uint8Array>,
  controlParts: ReadonlySet<string>
): boolean {
  for (const archivePath of controlParts) {
    let controlDepth = 0;
    let alternateFallbackDepth = 0;
    let nested = false;
    const parser = new SaxesParser({ xmlns: true });
    parser.on("opentag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth += 1;
        return;
      }
      if (alternateFallbackDepth > 0) {
        return;
      }
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        nested ||= controlDepth > 0;
        controlDepth += 1;
      }
    });
    parser.on("closetag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth -= 1;
        return;
      }
      if (
        alternateFallbackDepth === 0 &&
        tag.local === "sdt" &&
        templateWordNamespaces.has(tag.uri) &&
        controlDepth > 0
      ) {
        controlDepth -= 1;
      }
    });
    parser.write(templateArchiveText(archive, archivePath)).close();
    if (nested) {
      return true;
    }
  }
  return false;
}

function hasUnsupportedNativeDateFormatTokens(format: string): boolean {
  const unquoted = format.replaceAll(/'[^']*'|"[^"]*"/gu, "");
  const unrecognized = unquoted.replaceAll(
    /dddd|ddd|yyyy|MMMM|MMM|yy|MM|dd|M|d/gu,
    ""
  );
  return /[A-Za-z]/u.test(unrecognized);
}

function templatePartsHaveUnsupportedDateSettings(
  archive: Record<string, Uint8Array>,
  controlParts: ReadonlySet<string>
): boolean {
  for (const archivePath of controlParts) {
    let alternateFallbackDepth = 0;
    let inDateControl = false;
    let propertiesDepth = 0;
    let unsupportedDateSettings = false;
    const parser = new SaxesParser({ xmlns: true });
    parser.on("opentag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth += 1;
        return;
      }
      if (alternateFallbackDepth > 0) {
        return;
      }
      if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
        propertiesDepth = 1;
        inDateControl = false;
        return;
      }
      if (propertiesDepth === 0) {
        return;
      }
      if (
        propertiesDepth === 1 &&
        tag.local === "date" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        inDateControl = true;
      } else if (
        inDateControl &&
        tag.local === "calendar" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        const calendar = templateControlAttribute(
          templateXmlElement(tag),
          "val"
        );
        unsupportedDateSettings ||=
          calendar !== undefined && calendar !== "gregorian";
      } else if (
        inDateControl &&
        tag.local === "dateFormat" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        const format = templateControlAttribute(templateXmlElement(tag), "val");
        unsupportedDateSettings ||=
          format !== undefined && hasUnsupportedNativeDateFormatTokens(format);
      } else if (
        inDateControl &&
        tag.local === "lid" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        const language = templateControlAttribute(
          templateXmlElement(tag),
          "val"
        );
        unsupportedDateSettings ||=
          language !== undefined && !/^en(?:-|$)/iu.test(language);
      }
      propertiesDepth += 1;
    });
    parser.on("closetag", (tag) => {
      if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternateFallbackDepth -= 1;
        return;
      }
      if (alternateFallbackDepth > 0) {
        return;
      }
      if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
        propertiesDepth = 0;
        inDateControl = false;
      } else if (propertiesDepth > 0) {
        if (tag.local === "date" && templateWordNamespaces.has(tag.uri)) {
          inDateControl = false;
        }
        propertiesDepth -= 1;
      }
    });
    parser.write(templateArchiveText(archive, archivePath)).close();
    if (unsupportedDateSettings) {
      return true;
    }
  }
  return false;
}

interface NativeFontAttributeNames {
  ascii: string | null;
  asciiTheme: string | null;
  elementName: string;
  hAnsi: string | null;
  hAnsiTheme: string | null;
}
interface NativeFieldControlFrame {
  block: boolean;
  checkboxElement: {
    attributeDeclaration: string;
    attributeName: string;
    end: number;
    name: string;
    prefix: string;
    start: number;
  } | null;
  checkboxHasChecked: boolean;
  checkedState: NativeCheckboxState | null;
  contentEnd: number | null;
  contentName: string | null;
  contentPrefix: string;
  contentNamespace: string;
  contentSelfClosing: boolean;
  contentStart: number | null;
  contentTokenStart: number | null;
  paragraphProperties: string | null;
  paragraphPropertiesStart: number | null;
  runProperties: string | null;
  runPropertiesStart: number | null;
  contentFontAttributeNames: (NativeFontAttributeNames | null)[];
  runPropertiesFontAttributeNames: NativeFontAttributeNames | null | undefined;
  inPropertiesDepth: number;
  dateFormat: string | null;
  tag: string | null;
  uncheckedState: NativeCheckboxState | null;
}
interface NativeCheckboxState {
  font: string | null;
  glyph: string | null;
}
function nativeCheckboxState(element: TemplateXmlElement): NativeCheckboxState {
  const value = templateControlAttribute(element, "val");
  const codePoint =
    value && /^[\da-f]{1,6}$/iu.test(value)
      ? Number.parseInt(value, 16)
      : Number.NaN;
  return {
    font: templateControlAttribute(element, "font") ?? null,
    glyph:
      Number.isInteger(codePoint) &&
      codePoint <= 0x10_ff_ff &&
      !(codePoint >= 0xd8_00 && codePoint <= 0xdf_ff)
        ? String.fromCodePoint(codePoint)
        : null,
  };
}
function nativeNamespacedAttribute(
  elementName: string,
  namespaceBindings: ReadonlyMap<string, string>,
  namespaceUri: string,
  localName: string,
  openingTag: string
): { declaration: string; name: string } {
  let prefix = elementName.includes(":")
    ? elementName.slice(0, elementName.indexOf(":"))
    : [...namespaceBindings].find(
        ([candidate, uri]) => candidate && uri === namespaceUri
      )?.[0];
  if (!prefix) {
    prefix = "word";
    for (
      let suffix = 1;
      namespaceBindings.has(prefix) || openingTag.includes(`${prefix}:`);
      suffix += 1
    ) {
      prefix = `word${suffix}`;
    }
    return {
      declaration: ` xmlns:${prefix}="${namespaceUri}"`,
      name: `${prefix}:${localName}`,
    };
  }
  return { declaration: "", name: `${prefix}:${localName}` };
}
function nativeRunPropertiesXml(
  runProperties: string | null,
  contentPrefix: string,
  font: string | null,
  fontAttributeNames: NativeFontAttributeNames | null | undefined = undefined,
  wordNamespace = templateWordMainNamespace
): string {
  if (!font) {
    return runProperties ?? "";
  }
  const fontXml =
    contentPrefix.length > 0
      ? `<${contentPrefix}rFonts ${contentPrefix}ascii="${htmlEscape(font)}" ${contentPrefix}hAnsi="${htmlEscape(font)}"/>`
      : `<rFonts xmlns:word="${wordNamespace}" word:ascii="${htmlEscape(font)}" word:hAnsi="${htmlEscape(font)}"/>`;
  if (!runProperties) {
    return `<${contentPrefix}rPr>${fontXml}</${contentPrefix}rPr>`;
  }
  let fontElementStart: number | null = null;
  let fontElementOpenEnd: number | null = null;
  let fontElementEnd: number | null = null;
  let fontElementName: string | null = null;
  let fontOpeningTag = "";
  const parser = new SaxesParser({ position: true, xmlns: false });
  parser.on("opentag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (
      fontElementStart !== null ||
      localName !== "rFonts" ||
      fontAttributeNames === null ||
      (fontAttributeNames && tag.name !== fontAttributeNames.elementName)
    ) {
      return;
    }
    fontElementStart = runProperties.lastIndexOf("<", parser.position - 1);
    if (fontElementStart < 0) {
      return;
    }
    fontElementOpenEnd = parser.position;
    fontElementName = tag.name;
    const prefix = tag.name.includes(":")
      ? tag.name.slice(0, tag.name.lastIndexOf(":") + 1)
      : "";
    const setAttribute = (
      openingTag: string,
      attributeName: string
    ): string => {
      let replaced = false;
      const updated = openingTag.replaceAll(
        /(\s)([^\s="'<>/]+)(\s*=\s*)(["'])[\s\S]*?\4/gu,
        (
          match,
          whitespace: string,
          name: string,
          assignment: string,
          quote: string
        ) => {
          if (name !== attributeName) {
            return match;
          }
          replaced = true;
          return `${whitespace}${name}${assignment}${quote}${htmlEscape(font)}${quote}`;
        }
      );
      return replaced
        ? updated
        : updated.replace(
            /(\/?)>$/u,
            ` ${attributeName}="${htmlEscape(font)}"$1>`
          );
    };
    const removeAttribute = (
      openingTag: string,
      attributeName: string
    ): string =>
      openingTag.replaceAll(
        /(\s)([^\s="'<>/]+)(\s*=\s*)(["'])[\s\S]*?\4/gu,
        (match, _whitespace: string, name: string) =>
          name === attributeName ? "" : match
      );
    fontOpeningTag = runProperties.slice(fontElementStart, fontElementOpenEnd);
    let fontAttributePrefix =
      fontAttributeNames?.ascii?.slice(
        0,
        fontAttributeNames.ascii.lastIndexOf(":") + 1
      ) ??
      fontAttributeNames?.hAnsi?.slice(
        0,
        fontAttributeNames.hAnsi.lastIndexOf(":") + 1
      ) ??
      "";
    if (!fontAttributePrefix) {
      let prefix = "word";
      for (let suffix = 1; fontOpeningTag.includes(`${prefix}:`); suffix += 1) {
        prefix = `word${suffix}`;
      }
      fontAttributePrefix = `${prefix}:`;
      fontOpeningTag = fontOpeningTag.replace(
        /(\/?)>$/u,
        ` xmlns:${prefix}="${wordNamespace}"$1>`
      );
    }
    fontOpeningTag = setAttribute(
      fontOpeningTag,
      fontAttributeNames?.ascii ?? `${fontAttributePrefix}ascii`
    );
    fontOpeningTag = setAttribute(
      fontOpeningTag,
      fontAttributeNames?.hAnsi ?? `${fontAttributePrefix}hAnsi`
    );
    if (fontAttributeNames) {
      for (const attributeName of [
        fontAttributeNames.asciiTheme,
        fontAttributeNames.hAnsiTheme,
      ]) {
        if (attributeName) {
          fontOpeningTag = removeAttribute(fontOpeningTag, attributeName);
        }
      }
    } else if (fontAttributeNames === undefined) {
      fontOpeningTag = removeAttribute(fontOpeningTag, `${prefix}asciiTheme`);
      fontOpeningTag = removeAttribute(fontOpeningTag, `${prefix}hAnsiTheme`);
    }
    if (tag.isSelfClosing) {
      fontElementEnd = parser.position;
    }
  });
  parser.on("closetag", (tag) => {
    if (
      fontElementStart !== null &&
      fontElementEnd === null &&
      tag.name === fontElementName
    ) {
      fontElementEnd = parser.position;
    }
  });
  try {
    parser.write(runProperties).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  if (
    fontElementStart !== null &&
    fontElementOpenEnd !== null &&
    fontElementEnd !== null
  ) {
    return (
      runProperties.slice(0, fontElementStart) +
      fontOpeningTag +
      runProperties.slice(fontElementOpenEnd, fontElementEnd) +
      runProperties.slice(fontElementEnd)
    );
  }
  if (/\/\s*>$/u.test(runProperties)) {
    const propertiesName = runProperties.match(/^<([^\s/>]+)/u)?.[1];
    if (propertiesName) {
      return `${runProperties.replace(/\/\s*>$/u, ">")}${fontXml}</${propertiesName}>`;
    }
  }
  const closingTagStart = runProperties.lastIndexOf("</");
  if (closingTagStart === -1) {
    return runProperties;
  }
  return `${runProperties.slice(0, closingTagStart)}${fontXml}${runProperties.slice(closingTagStart)}`;
}

function nativeApplyTextControlFontXml(
  xml: string,
  font: string,
  fontAttributeNamesByElement: readonly (NativeFontAttributeNames | null)[],
  wordNamespace = templateWordMainNamespace
): string {
  interface RunFrame {
    fontAttributeNames: NativeFontAttributeNames | null | undefined;
    hasText: boolean;
    name: string;
    openingEnd: number;
    propertiesEnd: number | null;
    propertiesName: string | null;
    propertiesStart: number | null;
  }
  const runStack: RunFrame[] = [];
  const patches: { end: number; replacement: string; start: number }[] = [];
  const parser = new SaxesParser({
    fragment: true,
    position: true,
    xmlns: false,
  });
  let fontElementIndex = 0;
  parser.on("opentag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "rFonts") {
      const fontAttributeNames =
        fontAttributeNamesByElement[fontElementIndex] ?? null;
      fontElementIndex += 1;
      const run = runStack.at(-1);
      if (run && run.fontAttributeNames === undefined) {
        run.fontAttributeNames = fontAttributeNames;
      }
    } else if (localName === "r") {
      runStack.push({
        fontAttributeNames: undefined,
        hasText: false,
        name: tag.name,
        openingEnd: parser.position,
        propertiesEnd: null,
        propertiesName: null,
        propertiesStart: null,
      });
    } else if (localName === "rPr" && runStack.length > 0) {
      const run = runStack.at(-1);
      if (run && run.propertiesStart === null) {
        run.propertiesName = tag.name;
        run.propertiesStart = xml.lastIndexOf("<", parser.position - 1);
        run.propertiesEnd = tag.isSelfClosing ? parser.position : null;
      }
    } else if (localName === "t" && runStack.length > 0) {
      const run = runStack.at(-1);
      if (run) {
        run.hasText = true;
      }
    }
  });
  parser.on("closetag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "rPr") {
      const run = runStack.at(-1);
      if (run?.propertiesName === tag.name && run.propertiesEnd === null) {
        run.propertiesEnd = parser.position;
      }
    } else if (localName === "r" && runStack.at(-1)?.name === tag.name) {
      const run = runStack.pop();
      if (!run?.hasText) {
        return;
      }
      if (
        run.propertiesStart !== null &&
        run.propertiesEnd !== null &&
        run.propertiesName
      ) {
        const properties = xml.slice(run.propertiesStart, run.propertiesEnd);
        const prefix = run.propertiesName.includes(":")
          ? run.propertiesName.slice(0, run.propertiesName.lastIndexOf(":") + 1)
          : "";
        const replacement = nativeRunPropertiesXml(
          properties,
          prefix,
          font,
          run.fontAttributeNames,
          wordNamespace
        );
        if (replacement !== properties) {
          patches.push({
            end: run.propertiesEnd,
            replacement,
            start: run.propertiesStart,
          });
        }
        return;
      }
      const prefix = run.name.includes(":")
        ? run.name.slice(0, run.name.lastIndexOf(":") + 1)
        : "";
      patches.push({
        end: run.openingEnd,
        replacement: nativeRunPropertiesXml(
          null,
          prefix,
          font,
          undefined,
          wordNamespace
        ),
        start: run.openingEnd,
      });
    }
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  let updated = xml;
  for (const patch of patches.sort((left, right) => right.start - left.start)) {
    updated =
      updated.slice(0, patch.start) +
      patch.replacement +
      updated.slice(patch.end);
  }
  return updated;
}

const nativeDateMonths = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
] as const;
const nativeDateWeekdays = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
] as const;
function nativeDateDisplayValue(value: string, format: string | null): string {
  if (!format) {
    return value;
  }
  const [yearText, monthText, dayText] = value.split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  const day = Number(dayText);
  if (
    !yearText ||
    !monthText ||
    !dayText ||
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day)
  ) {
    return value;
  }
  const monthName = nativeDateMonths[month - 1];
  const weekday =
    nativeDateWeekdays[new Date(`${value}T00:00:00Z`).getUTCDay()];
  if (!monthName || !weekday) {
    return value;
  }
  const monthNumber = String(month);
  const dayNumber = String(day);
  return format.replaceAll(
    /'[^']*'|"[^"]*"|dddd|ddd|yyyy|MMMM|MMM|yy|MM|dd|M|d/gu,
    (token) => {
      if (token.startsWith("'") || token.startsWith('"')) {
        return token.slice(1, -1);
      }
      switch (token) {
        case "yyyy": {
          return yearText;
        }
        case "yy": {
          return yearText.slice(-2);
        }
        case "dddd": {
          return weekday;
        }
        case "ddd": {
          return weekday.slice(0, 3);
        }
        case "MMMM": {
          return monthName;
        }
        case "MMM": {
          return monthName.slice(0, 3);
        }
        case "MM": {
          return monthNumber.padStart(2, "0");
        }
        case "M": {
          return monthNumber;
        }
        case "dd": {
          return dayNumber.padStart(2, "0");
        }
        case "d": {
          return dayNumber;
        }
        default: {
          return token;
        }
      }
    }
  );
}

function nativeRewriteTextControlContentXml(
  xml: string,
  value: string
): string | null {
  interface TextNode {
    contentEnd: number;
    line: number;
    name: string;
    openingStart: number;
    openingTag: string;
    paragraph: number;
    selfClosing: boolean;
    start: number;
    tagEnd: number;
    text: string;
  }
  interface TabNode {
    end: number;
    line: number;
    name: string;
    paragraph: number;
    start: number;
  }
  interface Paragraph {
    end: number;
    index: number;
    start: number;
    startLine: number;
  }
  interface LineBreak {
    end: number;
    line: number;
    name: string;
    paragraph: number;
    start: number;
  }
  const textNodes: TextNode[] = [];
  const tabNodes: TabNode[] = [];
  const paragraphs: Paragraph[] = [];
  const lineBreaks: LineBreak[] = [];
  const paragraphStack: Omit<Paragraph, "end">[] = [];
  const breakStack: Omit<LineBreak, "end">[] = [];
  const tabStack: Omit<TabNode, "end">[] = [];
  let activeTextNode: Omit<TextNode, "contentEnd" | "tagEnd"> | null = null;
  let currentLine = 0;
  let currentParagraph = -1;
  let visibleText = "";
  const parser = new SaxesParser({
    fragment: true,
    position: true,
    xmlns: false,
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "p") {
      if (currentParagraph >= 0) {
        currentLine += 1;
        visibleText += "\n";
      }
      currentParagraph += 1;
      paragraphStack.push({
        index: currentParagraph,
        start: xml.lastIndexOf("<", parser.position - 1),
        startLine: currentLine,
      });
    } else if (localName === "br" || localName === "cr") {
      const lineBreak = {
        line: currentLine,
        name: tag.name,
        paragraph: currentParagraph,
        start: xml.lastIndexOf("<", parser.position - 1),
      };
      if (tag.isSelfClosing) {
        lineBreaks.push({ ...lineBreak, end: parser.position });
      } else {
        breakStack.push(lineBreak);
      }
      currentLine += 1;
      visibleText += "\n";
    } else if (localName === "tab") {
      const tab = {
        line: currentLine,
        name: tag.name,
        paragraph: currentParagraph,
        start: xml.lastIndexOf("<", parser.position - 1),
      };
      if (tag.isSelfClosing) {
        tabNodes.push({ ...tab, end: parser.position });
      } else {
        tabStack.push(tab);
      }
      visibleText += "\t";
    } else if (localName === "t") {
      const openingStart = xml.lastIndexOf("<", parser.position - 1);
      const textNode = {
        line: currentLine,
        name: tag.name,
        openingStart,
        openingTag: xml.slice(openingStart, parser.position),
        paragraph: currentParagraph,
        selfClosing: tag.isSelfClosing,
        start: parser.position,
        text: "",
      };
      if (tag.isSelfClosing) {
        textNodes.push({
          ...textNode,
          contentEnd: parser.position,
          tagEnd: parser.position,
        });
      } else {
        activeTextNode = textNode;
      }
    }
  });
  parser.on("text", (text) => {
    if (activeTextNode) {
      activeTextNode.text += text;
      visibleText += text;
    }
  });
  parser.on("closetag", (tag) => {
    const localName = tag.name.slice(tag.name.lastIndexOf(":") + 1);
    if (localName === "t" && activeTextNode?.name === tag.name) {
      const contentEnd = xml.lastIndexOf("</", parser.position - 1);
      textNodes.push({
        ...activeTextNode,
        contentEnd,
        tagEnd: parser.position,
      });
      activeTextNode = null;
    } else if (
      (localName === "br" || localName === "cr") &&
      breakStack.at(-1)?.name === tag.name
    ) {
      const lineBreak = breakStack.pop();
      if (lineBreak) {
        lineBreaks.push({ ...lineBreak, end: parser.position });
      }
    } else if (localName === "tab" && tabStack.at(-1)?.name === tag.name) {
      const tab = tabStack.pop();
      if (tab) {
        tabNodes.push({ ...tab, end: parser.position });
      }
    } else if (localName === "p") {
      const paragraph = paragraphStack.pop();
      if (paragraph) {
        paragraphs.push({ ...paragraph, end: parser.position });
      }
    }
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  if (textNodes.length === 0 || visibleText === value) {
    return textNodes.length === 0 ? null : xml;
  }
  const lines = value.split(/\r\n|\r|\n/u);
  const existingLineCount = currentLine + 1;
  const removedParagraphs = new Set(
    lines.length < existingLineCount
      ? paragraphs
          .filter(({ startLine }) => startLine >= lines.length)
          .map(({ index }) => index)
      : []
  );
  const removedParagraphRanges = paragraphs.filter(({ index }) =>
    removedParagraphs.has(index)
  );
  const keptTextNodes = textNodes.filter(
    ({ paragraph }) => !removedParagraphs.has(paragraph)
  );
  const keptTabNodes = tabNodes.filter(
    ({ paragraph }) => !removedParagraphs.has(paragraph)
  );
  type ContentSegment =
    | { kind: "text"; node: TextNode }
    | { kind: "tab"; node: TabNode };
  const segmentsByLine = new Map<number, ContentSegment[]>();
  for (const node of keptTextNodes) {
    const segments = segmentsByLine.get(node.line) ?? [];
    segments.push({ kind: "text", node });
    segmentsByLine.set(node.line, segments);
  }
  for (const node of keptTabNodes) {
    const segments = segmentsByLine.get(node.line) ?? [];
    segments.push({ kind: "tab", node });
    segmentsByLine.set(node.line, segments);
  }
  for (const segments of segmentsByLine.values()) {
    segments.sort((left, right) => {
      const leftStart =
        left.kind === "text" ? left.node.openingStart : left.node.start;
      const rightStart =
        right.kind === "text" ? right.node.openingStart : right.node.start;
      return leftStart - rightStart;
    });
  }
  for (
    let line = 0;
    line < Math.min(lines.length, existingLineCount);
    line += 1
  ) {
    if (lines[line] && !segmentsByLine.has(line)) {
      return null;
    }
  }
  const extraLines = lines.slice(existingLineCount);
  const appendSegment = extraLines.length
    ? segmentsByLine.get(existingLineCount - 1)?.at(-1)
    : undefined;
  if (extraLines.length > 0 && !appendSegment) {
    return null;
  }
  const assignments = new Map<TextNode, string>();
  const tabAssignments = new Map<TabNode, string>();
  for (const [line, segments] of segmentsByLine) {
    const characters = [...(lines[line] ?? "")];
    const originalLengths = segments.map((segment) =>
      segment.kind === "text" ? [...segment.node.text].length : 1
    );
    const originalLength = originalLengths.reduce(
      (total, length) => total + length,
      0
    );
    const assign = (segment: ContentSegment, text: string): void => {
      if (segment.kind === "text") {
        assignments.set(segment.node, text);
      } else {
        tabAssignments.set(segment.node, text);
      }
    };
    if (originalLength === 0) {
      const [first, ...rest] = segments;
      if (first) {
        assign(first, characters.join(""));
      }
      for (const segment of rest) {
        assign(segment, "");
      }
      continue;
    }
    let offset = 0;
    for (const [index, segment] of segments.entries()) {
      const end =
        index === segments.length - 1
          ? characters.length
          : Math.min(characters.length, offset + (originalLengths[index] ?? 0));
      assign(segment, characters.slice(offset, end).join(""));
      offset = end;
    }
  }
  const textOpeningTag = (openingTag: string, text: string): string => {
    const opening = openingTag.replace(/\/\s*>$/u, ">");
    return text.trim() === text || /\bxml:space\s*=/u.test(opening)
      ? opening
      : opening.replace(/>$/u, ' xml:space="preserve">');
  };
  const textContentXml = (
    name: string,
    openingTag: string,
    text: string
  ): string => {
    const prefix = name.includes(":")
      ? name.slice(0, name.lastIndexOf(":") + 1)
      : "";
    const [firstPart = "", ...tabParts] = text.split("\t");
    let replacement = htmlEscape(firstPart);
    for (const part of tabParts) {
      replacement += `</${name}><${prefix}tab/>${textOpeningTag(openingTag, part)}${htmlEscape(part)}`;
    }
    return replacement;
  };
  const patches: { end: number; replacement: string; start: number }[] = [];
  for (const paragraph of removedParagraphRanges) {
    patches.push({
      end: paragraph.end,
      replacement: "",
      start: paragraph.start,
    });
  }
  if (lines.length < existingLineCount) {
    for (const lineBreak of lineBreaks) {
      if (
        lineBreak.line >= lines.length - 1 &&
        !removedParagraphs.has(lineBreak.paragraph)
      ) {
        patches.push({
          end: lineBreak.end,
          replacement: "",
          start: lineBreak.start,
        });
      }
    }
  }
  for (const node of keptTextNodes) {
    const assigned = assignments.get(node) ?? "";
    const appended =
      appendSegment?.kind === "text" && node === appendSegment.node
        ? extraLines
        : [];
    if (node.selfClosing && !assigned && appended.length === 0) {
      continue;
    }
    const opening = textOpeningTag(node.openingTag, assigned);
    const prefix = node.name.includes(":")
      ? node.name.slice(0, node.name.lastIndexOf(":") + 1)
      : "";
    let replacement = textContentXml(node.name, opening, assigned);
    for (const line of appended) {
      replacement += `</${node.name}><${prefix}br/>${textOpeningTag(node.openingTag, line)}${textContentXml(node.name, opening, line)}`;
    }
    if (node.selfClosing) {
      patches.push({
        end: node.tagEnd,
        replacement: `${opening}${replacement}</${node.name}>`,
        start: node.openingStart,
      });
    } else {
      if (opening !== node.openingTag) {
        patches.push({
          end: node.start,
          replacement: opening,
          start: node.openingStart,
        });
      }
      if (replacement !== xml.slice(node.start, node.contentEnd)) {
        patches.push({
          end: node.contentEnd,
          replacement,
          start: node.start,
        });
      }
    }
  }
  for (const node of keptTabNodes) {
    const assigned = tabAssignments.get(node) ?? "";
    const appended =
      appendSegment?.kind === "tab" && node === appendSegment.node
        ? extraLines
        : [];
    if (assigned === "\t" && appended.length === 0) {
      continue;
    }
    const prefix = node.name.includes(":")
      ? node.name.slice(0, node.name.lastIndexOf(":") + 1)
      : "";
    if (!assigned && appended.length === 0) {
      patches.push({ end: node.end, replacement: "", start: node.start });
      continue;
    }
    const name = `${prefix}t`;
    const opening = textOpeningTag(`<${name} xml:space="preserve">`, assigned);
    let replacement = `${opening}${textContentXml(name, opening, assigned)}</${name}>`;
    for (const line of appended) {
      const lineOpening = textOpeningTag(opening, line);
      replacement += `<${prefix}br/>${lineOpening}${textContentXml(name, lineOpening, line)}</${name}>`;
    }
    patches.push({ end: node.end, replacement, start: node.start });
  }
  let updated = xml;
  for (const patch of patches.sort((left, right) => right.start - left.start)) {
    updated =
      updated.slice(0, patch.start) +
      patch.replacement +
      updated.slice(patch.end);
  }
  return updated;
}

function nativeFieldControlXml(
  xml: string,
  data: JsonRecord,
  fieldsByTag: Map<string, { options: unknown; type: FieldType }>,
  selectedTags?: ReadonlySet<string>
): string {
  const controlTags: (string | null)[] = [];
  let alternateTagFallbackDepth = 0;
  let controlTag: string | null = null;
  let inControlProperties = false;
  const tagParser = new SaxesParser({ xmlns: true });
  tagParser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  tagParser.on("opentag", (tag) => {
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateTagFallbackDepth += 1;
      return;
    }
    if (alternateTagFallbackDepth > 0) {
      return;
    }
    if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      controlTag = null;
    } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      inControlProperties = true;
    } else if (
      inControlProperties &&
      tag.local === "tag" &&
      templateWordNamespaces.has(tag.uri)
    ) {
      controlTag =
        templateControlAttribute(templateXmlElement(tag), "val") ?? null;
    }
  });
  tagParser.on("closetag", (tag) => {
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateTagFallbackDepth -= 1;
      return;
    }
    if (alternateTagFallbackDepth > 0) {
      return;
    }
    if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      inControlProperties = false;
    } else if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      controlTags.push(controlTag?.trim() ?? null);
    }
  });
  try {
    tagParser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  let controlIndex = 0;
  const controls: NativeFieldControlFrame[] = [];
  const elements: {
    local: string;
    namespaces: Map<string, string>;
    uri: string;
  }[] = [];
  const patches: { end: number; replacement: string; start: number }[] = [];
  const showingPlcHdrStarts: { remove: boolean; start: number }[] = [];
  let alternateFallbackDepth = 0;
  const parser = new SaxesParser({ position: true, xmlns: true });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    const namespaceBindings = new Map(elements.at(-1)?.namespaces);
    for (const [attributeName, attribute] of Object.entries(tag.attributes)) {
      if (attribute.uri === templateXmlnsNamespace) {
        namespaceBindings.set(
          attributeName === "xmlns"
            ? ""
            : attributeName.slice(attributeName.indexOf(":") + 1),
          attribute.value
        );
      }
    }
    const element = {
      local: tag.local,
      namespaces: namespaceBindings,
      uri: tag.uri,
    };
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateFallbackDepth += 1;
      elements.push(element);
      return;
    }
    if (alternateFallbackDepth > 0) {
      elements.push(element);
      return;
    }
    if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      if (controls.length > 0) {
        fail(422, "invalid_template", "Nested native fields are unsupported");
      }
      controls.push({
        block: !elements.some(
          (element) =>
            element.local === "p" && templateWordNamespaces.has(element.uri)
        ),
        checkboxElement: null,
        checkboxHasChecked: false,
        checkedState: null,
        contentEnd: null,
        contentFontAttributeNames: [],
        contentName: null,
        contentNamespace: "",
        contentPrefix: "",
        contentSelfClosing: false,
        contentStart: null,
        contentTokenStart: null,
        dateFormat: null,
        inPropertiesDepth: 0,
        paragraphProperties: null,
        paragraphPropertiesStart: null,
        runProperties: null,
        runPropertiesFontAttributeNames: undefined,
        runPropertiesStart: null,
        tag: controlTags[controlIndex++] ?? null,
        uncheckedState: null,
      });
    } else {
      const frame = controls.at(-1);
      if (frame) {
        if (
          tag.local === "rFonts" &&
          frame.contentStart !== null &&
          frame.contentEnd === null
        ) {
          let fontAttributeNames: NativeFontAttributeNames | null = null;
          if (templateWordNamespaces.has(tag.uri)) {
            fontAttributeNames = {
              ascii: null,
              asciiTheme: null,
              elementName: tag.name,
              hAnsi: null,
              hAnsiTheme: null,
            };
            for (const [attributeName, attribute] of Object.entries(
              tag.attributes
            )) {
              if (attribute.uri !== tag.uri) {
                continue;
              }
              if (attribute.local === "ascii") {
                fontAttributeNames.ascii = attributeName;
              } else if (attribute.local === "hAnsi") {
                fontAttributeNames.hAnsi = attributeName;
              } else if (attribute.local === "asciiTheme") {
                fontAttributeNames.asciiTheme = attributeName;
              } else if (attribute.local === "hAnsiTheme") {
                fontAttributeNames.hAnsiTheme = attributeName;
              }
            }
          }
          frame.contentFontAttributeNames.push(fontAttributeNames);
          if (
            frame.runPropertiesStart !== null &&
            frame.runPropertiesFontAttributeNames === undefined
          ) {
            frame.runPropertiesFontAttributeNames = fontAttributeNames;
          }
        }
        if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
          frame.inPropertiesDepth = 1;
        } else if (frame.inPropertiesDepth > 0) {
          if (
            frame.inPropertiesDepth === 1 &&
            tag.local === "tag" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            frame.tag =
              templateControlAttribute(templateXmlElement(tag), "val") ?? null;
            if (
              start !== -1 &&
              frame.tag !== null &&
              (!selectedTags || selectedTags.has(frame.tag.trim()))
            ) {
              const openingTag = xml.slice(start, parser.position);
              const normalizedTag = frame.tag.trim();
              if (normalizedTag !== frame.tag) {
                const tagValueAttributeName = Object.entries(
                  tag.attributes
                ).find(
                  ([, attribute]) =>
                    attribute.local === "val" && attribute.uri === tag.uri
                )?.[0];
                if (tagValueAttributeName) {
                  patches.push({
                    end: parser.position,
                    replacement: openingTag.replaceAll(
                      /(\s)([^\s="'<>/]+)(\s*=\s*)(["'])[\s\S]*?\4/gu,
                      (match, whitespace, attributeName, assignment, quote) =>
                        attributeName === tagValueAttributeName
                          ? `${whitespace}${attributeName}${assignment}${quote}${htmlEscape(normalizedTag)}${quote}`
                          : match
                    ),
                    start,
                  });
                }
              }
            }
          }
          if (
            tag.local === "dateFormat" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            frame.dateFormat =
              templateControlAttribute(templateXmlElement(tag), "val") ?? null;
          }
          if (
            tag.local === "showingPlcHdr" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            const remove =
              !selectedTags || selectedTags.has(frame.tag?.trim() ?? "");
            if (tag.isSelfClosing && remove && start !== -1) {
              patches.push({ end: parser.position, replacement: "", start });
            }
            showingPlcHdrStarts.push({
              remove: remove && !tag.isSelfClosing,
              start,
            });
          }
          const field = frame.tag
            ? fieldsByTag.get(frame.tag.trim())
            : undefined;
          const propertyName =
            tag.local === "date" && field?.type === FieldType.date
              ? "fullDate"
              : (tag.local === "dropDownList" &&
                    field?.type === FieldType.dropdown) ||
                  (tag.local === "comboBox" && field?.type === FieldType.combo)
                ? "lastValue"
                : null;
          if (propertyName) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            const openingTag =
              start === -1 ? "" : xml.slice(start, parser.position);
            const propertyAttributeName = Object.entries(tag.attributes).find(
              ([, attribute]) =>
                attribute.local === propertyName && attribute.uri === tag.uri
            )?.[0];
            const storedValue = frame.tag ? data[frame.tag.trim()] : undefined;
            const replacementValue =
              typeof storedValue === "string" &&
              !(propertyName === "fullDate" && storedValue === "")
                ? propertyName === "fullDate"
                  ? `${storedValue}T00:00:00Z`
                  : storedValue
                : null;
            if (
              start !== -1 &&
              (propertyAttributeName || replacementValue !== null)
            ) {
              let replacement: string;
              if (propertyAttributeName) {
                replacement = openingTag.replaceAll(
                  /(\s)([^\s="'<>/]+)(\s*=\s*)(["'])[\s\S]*?\4/gu,
                  (match, whitespace, name, assignment, quote) =>
                    name === propertyAttributeName
                      ? replacementValue === null
                        ? ""
                        : `${whitespace}${name}${assignment}${quote}${htmlEscape(replacementValue)}${quote}`
                      : match
                );
              } else {
                let prefix = tag.name.includes(":")
                  ? tag.name.slice(0, tag.name.indexOf(":") + 1)
                  : (Object.entries(tag.attributes)
                      .find(
                        ([name, attribute]) =>
                          name.includes(":") && attribute.uri === tag.uri
                      )?.[0]
                      .replace(/[^:]+$/u, ":") ?? "");
                let namespaceDeclaration = "";
                if (!prefix) {
                  let localPrefix = "word";
                  for (
                    let suffix = 1;
                    openingTag.includes(`${localPrefix}:`);
                    suffix += 1
                  ) {
                    localPrefix = `word${suffix}`;
                  }
                  prefix = `${localPrefix}:`;
                  namespaceDeclaration = ` xmlns:${localPrefix}="${tag.uri}"`;
                }
                replacement = openingTag.replace(
                  /(\/?)>$/u,
                  ` ${prefix}${propertyName}="${htmlEscape(replacementValue ?? "")}"${namespaceDeclaration}$1>`
                );
              }
              patches.push({
                end: parser.position,
                replacement,
                start,
              });
            }
          }
          if (
            tag.local === "checkbox" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            const start = xml.lastIndexOf("<", parser.position - 1);
            const openingTag =
              start === -1 ? "" : xml.slice(start, parser.position);
            const prefix = tag.name.includes(":")
              ? tag.name.slice(0, tag.name.indexOf(":") + 1)
              : "";
            const valueAttribute = nativeNamespacedAttribute(
              tag.name,
              namespaceBindings,
              tag.uri,
              "val",
              openingTag
            );
            const checkedValue = data[frame.tag.trim()] === true ? "1" : "0";
            if (/\/\s*>$/u.test(openingTag)) {
              patches.push({
                end: parser.position,
                replacement: `${openingTag.replace(/\/\s*>$/u, ">")}<${prefix}checked ${valueAttribute.name}="${checkedValue}"${valueAttribute.declaration}/></${tag.name}>`,
                start,
              });
            } else {
              frame.checkboxElement = {
                attributeDeclaration: valueAttribute.declaration,
                attributeName: valueAttribute.name,
                end: parser.position,
                name: tag.name,
                prefix,
                start,
              };
              frame.checkboxHasChecked = false;
            }
          }
          if (
            tag.local === "checkedState" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            frame.checkedState = nativeCheckboxState(templateXmlElement(tag));
          }
          if (
            tag.local === "uncheckedState" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            frame.uncheckedState = nativeCheckboxState(templateXmlElement(tag));
          }
          if (
            tag.local === "checked" &&
            tag.uri === templateCheckboxNamespace &&
            frame.tag &&
            fieldsByTag.get(frame.tag.trim())?.type === FieldType.checkbox
          ) {
            frame.checkboxHasChecked = true;
            const start = xml.lastIndexOf("<", parser.position - 1);
            const checkedAttributeName = Object.entries(tag.attributes).find(
              ([, attribute]) =>
                attribute.local === "val" && attribute.uri === tag.uri
            )?.[0];
            const unqualifiedValueName = Object.entries(tag.attributes).find(
              ([, attribute]) =>
                attribute.local === "val" && attribute.uri === ""
            )?.[0];
            if (start !== -1) {
              const openingTag = xml.slice(start, parser.position);
              const checkedValue = data[frame.tag.trim()] === true ? "1" : "0";
              const qualifiedAttribute = checkedAttributeName
                ? { declaration: "", name: checkedAttributeName }
                : nativeNamespacedAttribute(
                    tag.name,
                    namespaceBindings,
                    tag.uri,
                    "val",
                    openingTag
                  );
              const replacement =
                checkedAttributeName || unqualifiedValueName
                  ? openingTag.replaceAll(
                      /(\s)([^\s="'<>/]+)(\s*=\s*)(["'])[\s\S]*?\4/gu,
                      (match, whitespace, name, assignment, quote) =>
                        name === checkedAttributeName ||
                        name === unqualifiedValueName
                          ? `${whitespace}${qualifiedAttribute.name}${assignment}${quote}${checkedValue}${quote}${qualifiedAttribute.declaration}`
                          : match
                    )
                  : openingTag.replace(
                      /(\/?)>$/u,
                      ` ${qualifiedAttribute.name}="${checkedValue}"${qualifiedAttribute.declaration}$1>`
                    );
              patches.push({
                end: parser.position,
                replacement,
                start,
              });
            }
          }
          frame.inPropertiesDepth += 1;
        } else if (
          frame.contentStart !== null &&
          frame.inPropertiesDepth === 0 &&
          tag.local === "pPr" &&
          templateWordNamespaces.has(tag.uri) &&
          frame.paragraphProperties === null &&
          frame.paragraphPropertiesStart === null
        ) {
          frame.paragraphPropertiesStart = xml.lastIndexOf(
            "<",
            parser.position - 1
          );
        } else if (
          frame.contentStart !== null &&
          frame.inPropertiesDepth === 0 &&
          tag.local === "rPr" &&
          templateWordNamespaces.has(tag.uri) &&
          frame.runProperties === null &&
          frame.runPropertiesStart === null
        ) {
          frame.runPropertiesStart = xml.lastIndexOf("<", parser.position - 1);
        } else if (
          tag.local === "sdtContent" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          const start = xml.lastIndexOf("<", parser.position - 1);
          const openingTag =
            start === -1 ? "" : xml.slice(start, parser.position);
          frame.contentNamespace = tag.uri;
          frame.contentName = tag.name;
          frame.contentPrefix = tag.name.includes(":")
            ? tag.name.slice(0, tag.name.indexOf(":") + 1)
            : "";
          frame.contentSelfClosing = /\/\s*>$/u.test(openingTag);
          frame.contentStart = frame.contentSelfClosing
            ? null
            : parser.position;
          frame.contentTokenStart = frame.contentSelfClosing ? start : null;
        }
      }
    }
    elements.push(element);
  });
  parser.on("closetag", (tag) => {
    if (
      tag.local === "Fallback" &&
      tag.uri === templateMarkupCompatibilityNamespace
    ) {
      alternateFallbackDepth -= 1;
      elements.pop();
      return;
    }
    if (alternateFallbackDepth > 0) {
      elements.pop();
      return;
    }
    if (tag.local === "showingPlcHdr" && templateWordNamespaces.has(tag.uri)) {
      const placeholder = showingPlcHdrStarts.pop();
      if (placeholder?.remove && placeholder.start >= 0) {
        patches.push({
          end: parser.position,
          replacement: "",
          start: placeholder.start,
        });
      }
    }
    const frame = controls.at(-1);
    if (frame) {
      if (
        frame.contentStart !== null &&
        frame.paragraphPropertiesStart !== null &&
        tag.local === "pPr" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        frame.paragraphProperties = xml.slice(
          frame.paragraphPropertiesStart,
          parser.position
        );
        frame.paragraphPropertiesStart = null;
      } else if (
        frame.contentStart !== null &&
        frame.runPropertiesStart !== null &&
        tag.local === "rPr" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        frame.runProperties = xml.slice(
          frame.runPropertiesStart,
          parser.position
        );
        frame.runPropertiesStart = null;
      } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
        frame.inPropertiesDepth = 0;
      } else if (frame.inPropertiesDepth > 0) {
        if (
          tag.local === "checkbox" &&
          tag.uri === templateCheckboxNamespace &&
          frame.checkboxElement
        ) {
          if (!frame.checkboxHasChecked) {
            const { attributeDeclaration, attributeName, end, prefix, start } =
              frame.checkboxElement;
            const checkedValue =
              data[frame.tag?.trim() ?? ""] === true ? "1" : "0";
            patches.push({
              end,
              replacement: `${xml.slice(start, end)}<${prefix}checked ${attributeName}="${checkedValue}"${attributeDeclaration}/>`,
              start,
            });
          }
          frame.checkboxElement = null;
        }
        frame.inPropertiesDepth -= 1;
      } else if (
        tag.local === "sdtContent" &&
        templateWordNamespaces.has(tag.uri)
      ) {
        frame.contentEnd = frame.contentSelfClosing
          ? parser.position
          : xml.lastIndexOf("</", parser.position - 1);
        const completed = controls.pop();
        if (!completed?.tag || !completed.contentName) {
          fail(422, "invalid_template", "A native field is incomplete");
        }
        if (selectedTags && !selectedTags.has(completed.tag.trim())) {
          elements.pop();
          return;
        }
        const tagName = completed.tag.trim();
        const field = fieldsByTag.get(tagName);
        const storedValue = data[tagName];
        const checkboxChecked = storedValue === true;
        const checkboxState = checkboxChecked
          ? completed.checkedState
          : completed.uncheckedState;
        const value =
          field?.type === FieldType.checkbox
            ? (checkboxState?.glyph ?? (checkboxChecked ? "☒" : "☐"))
            : typeof storedValue === "string"
              ? storedValue
              : "";
        const option =
          typeof storedValue === "string" &&
          (field?.type === FieldType.dropdown ||
            field?.type === FieldType.combo)
            ? (Array.isArray(field.options) ? field.options : []).find(
                (item) =>
                  item &&
                  typeof item === "object" &&
                  !Array.isArray(item) &&
                  (item as Record<string, unknown>).value === storedValue
              )
            : null;
        const displayValue =
          field?.type === FieldType.date
            ? nativeDateDisplayValue(value, completed.dateFormat)
            : option &&
                typeof (option as Record<string, unknown>).displayText ===
                  "string"
              ? ((option as Record<string, unknown>).displayText as string)
              : value;
        const existingContent =
          !completed.contentSelfClosing &&
          completed.contentStart !== null &&
          completed.contentEnd !== null
            ? xml.slice(completed.contentStart, completed.contentEnd)
            : null;
        let preservedContent =
          existingContent === null
            ? null
            : nativeRewriteTextControlContentXml(existingContent, displayValue);
        if (preservedContent !== null && checkboxState?.font) {
          preservedContent = nativeApplyTextControlFontXml(
            preservedContent,
            checkboxState.font,
            completed.contentFontAttributeNames,
            completed.contentNamespace
          );
        }
        if (preservedContent === null) {
          const lines = displayValue.split(/\r\n|\r|\n/u);
          const text = lines
            .map(
              (line, index) =>
                `${index > 0 ? `<${completed.contentPrefix}br/>` : ""}<${completed.contentPrefix}t xml:space="preserve">${htmlEscape(line)}</${completed.contentPrefix}t>`
            )
            .join("");
          const runProperties = nativeRunPropertiesXml(
            completed.runProperties,
            completed.contentPrefix,
            checkboxState?.font ?? null,
            completed.runPropertiesFontAttributeNames,
            completed.contentNamespace
          );
          const run = `<${completed.contentPrefix}r>${runProperties}${text}</${completed.contentPrefix}r>`;
          const content = completed.block
            ? `<${completed.contentPrefix}p>${completed.paragraphProperties ?? ""}${run}</${completed.contentPrefix}p>`
            : run;
          if (completed.contentSelfClosing) {
            if (
              completed.contentTokenStart === null ||
              completed.contentEnd === null
            ) {
              fail(422, "invalid_template", "A native field is incomplete");
            }
            const openingTag = xml.slice(
              completed.contentTokenStart,
              completed.contentEnd
            );
            const expandedOpeningTag = openingTag.replace(/\/\s*>$/u, ">");
            patches.push({
              end: completed.contentEnd,
              replacement: `${expandedOpeningTag}${content}</${completed.contentName}>`,
              start: completed.contentTokenStart,
            });
          } else if (
            completed.contentStart !== null &&
            completed.contentEnd !== null &&
            completed.contentEnd >= completed.contentStart
          ) {
            patches.push({
              end: completed.contentEnd,
              replacement: content,
              start: completed.contentStart,
            });
          } else {
            fail(422, "invalid_template", "A native field is incomplete");
          }
        } else {
          if (preservedContent !== existingContent) {
            const contentStart = completed.contentStart;
            const contentEnd = completed.contentEnd;
            if (contentStart === null || contentEnd === null) {
              fail(422, "invalid_template", "A native field is incomplete");
            }
            patches.push({
              end: contentEnd,
              replacement: preservedContent,
              start: contentStart,
            });
          }
        }
      }
    }
    elements.pop();
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  let result = xml;
  for (const patch of patches.sort((left, right) => right.start - left.start)) {
    result =
      result.slice(0, patch.start) +
      patch.replacement +
      result.slice(patch.end);
  }
  return result;
}
function nativePreserveAlternateNamespaceDeclarations(
  xml: string,
  declarations: readonly { name: string; value: string }[]
): string {
  if (declarations.length === 0) {
    return xml;
  }
  let depth = 0;
  const patches: { end: number; replacement: string; start: number }[] = [];
  const parser = new SaxesParser({
    fragment: true,
    position: true,
    xmlns: false,
  });
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    if (depth === 0) {
      const start = xml.lastIndexOf("<", parser.position - 1);
      if (start !== -1) {
        const originalTag = xml.slice(start, parser.position);
        let replacement = originalTag;
        for (const declaration of declarations) {
          if (!Object.hasOwn(tag.attributes, declaration.name)) {
            replacement = replacement.replace(
              /(\/?)>$/u,
              ` ${declaration.name}="${htmlEscape(declaration.value)}"$1>`
            );
          }
        }
        if (replacement !== originalTag) {
          patches.push({ end: parser.position, replacement, start });
        }
      }
    }
    if (!tag.isSelfClosing) {
      depth += 1;
    }
  });
  parser.on("closetag", (tag) => {
    if (!tag.isSelfClosing) {
      depth -= 1;
    }
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "DOCX XML is malformed");
  }
  let result = xml;
  for (const patch of patches.sort((left, right) => right.start - left.start)) {
    result =
      result.slice(0, patch.start) +
      patch.replacement +
      result.slice(patch.end);
  }
  return result;
}

// PDF conversion omits AlternateContent fields; retain compatible Choice content and namespace bindings.
function nativeFlattenAlternateFieldsXml(
  xml: string,
  fieldTags: ReadonlySet<string>,
  validateOnly = false
): { xml: string; supported: boolean } {
  let result = xml;
  while (result.includes("AlternateContent")) {
    interface AlternateContentFrame {
      branch: "choice" | "fallback" | null;
      choiceContentStart: number | null;
      choiceFieldCount: number;
      choiceFieldEnd: number | null;
      choiceFieldStart: number | null;
      choiceHasField: boolean;
      choiceNamespaceDeclarations: Map<string, string> | null;
      choiceNamespaces: Map<string, string> | null;
      choiceRequiresSupported: boolean;
      fallbackHasField: boolean;
      hasField: boolean;
      hasNestedField: boolean;
      namespaceDeclarations: Map<string, string>;
      parent: AlternateContentFrame | null;
      start: number;
      unsupported: boolean;
      firstSupportedChoiceHasField: boolean | null;
    }
    interface AlternateControlFrame {
      alternate: AlternateContentFrame | null;
      branch: "choice" | "fallback" | null;
      inPropertiesDepth: number;
      start: number;
      tag: string | null;
    }
    const alternates: AlternateContentFrame[] = [];
    const controls: AlternateControlFrame[] = [];
    const patches: { end: number; replacement: string; start: number }[] = [];
    const namespaceBindings = new Map<string, string>([
      ["xml", "http://www.w3.org/XML/1998/namespace"],
    ]);
    const namespaceChanges: {
      prefix: string;
      previous: string | undefined;
    }[][] = [];
    const parser = new SaxesParser({ position: true, xmlns: true });
    let unsupported = false;
    parser.on("doctype", () => {
      fail(422, "invalid_template", "DOCX XML document types are not allowed");
    });
    parser.on("opentag", (tag) => {
      const changes: { prefix: string; previous: string | undefined }[] = [];
      for (const attribute of Object.values(tag.attributes)) {
        if (attribute.uri !== templateXmlnsNamespace) {
          continue;
        }
        const prefix = attribute.name === "xmlns" ? "" : attribute.local;
        changes.push({ prefix, previous: namespaceBindings.get(prefix) });
        namespaceBindings.set(prefix, attribute.value);
      }
      namespaceChanges.push(changes);
      if (
        tag.local === "AlternateContent" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const declarations = new Map<string, string>();
        for (const attribute of Object.values(tag.attributes)) {
          if (attribute.uri === templateXmlnsNamespace) {
            declarations.set(attribute.name, attribute.value);
          }
        }
        alternates.push({
          branch: null,
          choiceContentStart: null,
          choiceFieldCount: 0,
          choiceFieldEnd: null,
          choiceFieldStart: null,
          choiceHasField: false,
          choiceNamespaceDeclarations: null,
          choiceNamespaces: null,
          choiceRequiresSupported: false,
          fallbackHasField: false,
          firstSupportedChoiceHasField: null,
          hasField: false,
          hasNestedField: false,
          namespaceDeclarations: declarations,
          parent: alternates.at(-1) ?? null,
          start: result.lastIndexOf("<", parser.position - 1),
          unsupported: false,
        });
      } else if (
        tag.local === "Choice" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const alternate = alternates.at(-1);
        if (alternate) {
          const requires = Object.values(tag.attributes).find(
            (attribute) =>
              attribute.local === "Requires" && attribute.uri === ""
          )?.value;
          const requiredPrefixes =
            requires?.trim().split(/\s+/u).filter(Boolean) ?? [];
          alternate.choiceRequiresSupported =
            requiredPrefixes.length > 0 &&
            requiredPrefixes.every((prefix) =>
              templateControlNamespaces.has(namespaceBindings.get(prefix) ?? "")
            );
          alternate.choiceNamespaceDeclarations = new Map(
            alternate.namespaceDeclarations
          );
          for (const attribute of Object.values(tag.attributes)) {
            if (attribute.uri === templateXmlnsNamespace) {
              alternate.choiceNamespaceDeclarations.set(
                attribute.name,
                attribute.value
              );
            }
          }
          alternate.branch = "choice";
          alternate.choiceContentStart = parser.position;
          alternate.choiceHasField = false;
        }
      } else if (
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const alternate = alternates.at(-1);
        if (alternate) {
          alternate.branch = "fallback";
        }
      }
      const control = controls.at(-1);
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        const alternate = alternates.at(-1) ?? null;
        controls.push({
          alternate,
          branch: alternate?.branch ?? null,
          inPropertiesDepth: 0,
          start: result.lastIndexOf("<", parser.position - 1),
          tag: null,
        });
      } else if (control) {
        if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
          control.inPropertiesDepth = 1;
        } else if (control.inPropertiesDepth > 0) {
          if (
            control.inPropertiesDepth === 1 &&
            tag.local === "tag" &&
            templateWordNamespaces.has(tag.uri)
          ) {
            control.tag =
              templateControlAttribute(templateXmlElement(tag), "val") ?? null;
          }
          control.inPropertiesDepth += 1;
        }
      }
    });
    parser.on("closetag", (tag) => {
      if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
        const completed = controls.pop();
        const fieldTag = completed?.tag?.trim();
        if (completed?.alternate && fieldTag && fieldTags.has(fieldTag)) {
          completed.alternate.hasField = true;
          if (completed.branch === "choice") {
            completed.alternate.choiceHasField = true;
          } else if (completed.branch === "fallback") {
            completed.alternate.fallbackHasField = true;
          }
        }
      } else {
        const control = controls.at(-1);
        if (
          control &&
          tag.local === "sdtPr" &&
          templateWordNamespaces.has(tag.uri)
        ) {
          control.inPropertiesDepth = 0;
        } else if (control && control.inPropertiesDepth > 0) {
          control.inPropertiesDepth -= 1;
        }
      }
      const alternate = alternates.at(-1);
      if (
        alternate &&
        tag.local === "Choice" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        if (alternate.choiceRequiresSupported) {
          if (alternate.firstSupportedChoiceHasField === null) {
            alternate.firstSupportedChoiceHasField = alternate.choiceHasField;
          } else if (
            !alternate.firstSupportedChoiceHasField &&
            alternate.choiceHasField
          ) {
            alternate.unsupported = true;
          }
        }
        if (alternate.choiceHasField) {
          alternate.choiceFieldCount += 1;
          alternate.unsupported ||= !alternate.choiceRequiresSupported;
          if (
            alternate.choiceFieldStart === null &&
            alternate.choiceContentStart !== null
          ) {
            alternate.choiceFieldStart = alternate.choiceContentStart;
            alternate.choiceFieldEnd = result.lastIndexOf(
              "</",
              parser.position - 1
            );
            alternate.choiceNamespaces = alternate.choiceNamespaceDeclarations;
          }
        }
        alternate.branch = null;
        alternate.choiceContentStart = null;
        alternate.choiceHasField = false;
        alternate.choiceNamespaceDeclarations = null;
      } else if (
        alternate &&
        tag.local === "Fallback" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        alternate.branch = null;
      } else if (
        alternate &&
        tag.local === "AlternateContent" &&
        tag.uri === templateMarkupCompatibilityNamespace
      ) {
        const completed = alternates.pop();
        if (completed) {
          completed.unsupported ||= completed.choiceFieldCount > 1;
          completed.unsupported ||= completed.fallbackHasField;
          unsupported ||= completed.unsupported;
          if (completed.hasField && completed.parent) {
            completed.parent.hasNestedField = true;
            completed.parent.hasField = true;
            if (completed.parent.branch === "choice") {
              completed.parent.choiceHasField = true;
            } else if (completed.parent.branch === "fallback") {
              completed.parent.fallbackHasField = true;
            }
          }
          if (
            !validateOnly &&
            completed.hasField &&
            !completed.hasNestedField &&
            completed.choiceFieldStart !== null &&
            completed.choiceFieldEnd !== null
          ) {
            const declarations = Array.from(
              completed.choiceNamespaces ?? [],
              ([name, value]) => ({ name, value })
            );
            patches.push({
              end: parser.position,
              replacement: nativePreserveAlternateNamespaceDeclarations(
                result.slice(
                  completed.choiceFieldStart,
                  completed.choiceFieldEnd
                ),
                declarations
              ),
              start: completed.start,
            });
          }
        }
      }
      const changes = namespaceChanges.pop() ?? [];
      for (const { prefix, previous } of changes) {
        if (previous === undefined) {
          namespaceBindings.delete(prefix);
        } else {
          namespaceBindings.set(prefix, previous);
        }
      }
    });
    try {
      parser.write(result).close();
    } catch {
      fail(422, "invalid_template", "The DOCX package contains invalid XML");
    }
    if (unsupported) {
      return { supported: false, xml: result };
    }
    if (validateOnly) {
      return { supported: true, xml: result };
    }
    let flattened = result;
    for (const patch of patches.sort(
      (left, right) => right.start - left.start
    )) {
      flattened =
        flattened.slice(0, patch.start) +
        patch.replacement +
        flattened.slice(patch.end);
    }
    if (flattened === result) {
      return { supported: true, xml: result };
    }
    result = flattened;
  }
  return { supported: true, xml: result };
}

function overlayNativeResponseDocument(
  callbackBytes: Uint8Array,
  fields: readonly { options: unknown; tag: string; type: FieldType }[],
  data: JsonRecord,
  selectedTags: ReadonlySet<string>
): Uint8Array {
  const { archive, xmlPaths } = safeTemplateArchive(callbackBytes);
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  if (templatePartsHaveNestedControls(archive, controlParts)) {
    fail(422, "invalid_template", "Native fields cannot be nested");
  }
  if (controlParts.size === 0) {
    fail(422, "invalid_template", "The document has no fields");
  }
  const fieldsByTag = new Map(
    fields
      .filter(
        ({ tag, type }) => type !== FieldType.picture && selectedTags.has(tag)
      )
      .map(({ options, tag, type }) => [tag, { options, type }])
  );
  const fieldTags = new Set(fields.map(({ tag }) => tag));
  for (const archivePath of controlParts) {
    const originalXml = templateArchiveText(archive, archivePath);
    const flattened = nativeFlattenAlternateFieldsXml(
      nativeFieldControlXml(originalXml, data, fieldsByTag, selectedTags),
      fieldTags
    );
    if (!flattened.supported) {
      fail(
        422,
        "invalid_template",
        "AlternateContent field branches are unsupported"
      );
    }
    const xml = flattened.xml;
    if (xml !== originalXml) {
      archive[archivePath] = strToU8(
        xml.replace(/encoding=(["'])UTF-16(?:LE|BE)?\1/iu, 'encoding="UTF-8"')
      );
    }
  }
  return zipSync(archive);
}

function nativePdfScalarControlXml(xml: string): string {
  interface ScalarControlFrame {
    contentCloseStart: number | null;
    contentDepth: number | null;
    contentHasContent: boolean;
    contentIsBlock: boolean;
    contentOpenEnd: number | null;
    contentSelfClosing: boolean;
    namespaceAttributeNames: ReadonlySet<string>;
    namespaceDeclarations: { name: string; value: string }[];
    parentContentControls: ScalarControlFrame[] | null;
    parentIsBlockContainer: boolean;
    scalarControl: boolean;
    start: number;
    startEnd: number;
    wordNamespace: string;
    wordPrefix: string;
  }
  const controls: ScalarControlFrame[] = [];
  const elementStack: { local: string; uri: string }[] = [];
  const propertyControls: ScalarControlFrame[] = [];
  const patches: { end: number; replacement: string; start: number }[] = [];
  const parser = new SaxesParser({ position: true, xmlns: true });
  const preserveNamespaces = (
    start: number,
    end: number,
    declarations: readonly { name: string; value: string }[],
    existingAttributes: ReadonlySet<string>
  ) => {
    if (declarations.length === 0) {
      return;
    }
    const original = xml.slice(start, end);
    let replacement = original;
    for (const declaration of declarations) {
      if (existingAttributes.has(declaration.name)) {
        continue;
      }
      const insertionIndex = replacement.endsWith("/>")
        ? replacement.length - 2
        : replacement.length - 1;
      replacement = `${replacement.slice(0, insertionIndex)} ${declaration.name}="${htmlEscape(declaration.value)}"${replacement.slice(insertionIndex)}`;
    }
    if (replacement !== original) {
      patches.push({ end, replacement, start });
    }
  };
  const activeScalarNamespaceDeclarations = () => {
    const declarations = new Map<string, { name: string; value: string }>();
    for (const control of controls) {
      if (!control.scalarControl || control.contentDepth === null) {
        continue;
      }
      for (const declaration of control.namespaceDeclarations) {
        declarations.set(declaration.name, declaration);
      }
    }
    return [...declarations.values()];
  };
  parser.on("doctype", () => {
    fail(422, "invalid_template", "DOCX XML document types are not allowed");
  });
  parser.on("opentag", (tag) => {
    const start = xml.lastIndexOf("<", parser.position - 1);
    let hasContentControl = false;
    let parentContentControls: ScalarControlFrame[] | null = null;
    for (const control of controls) {
      if (control.contentDepth === null) {
        continue;
      }
      if (control.contentDepth === 0) {
        hasContentControl = true;
        control.contentHasContent = true;
        if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
          (parentContentControls ??= []).push(control);
        }
      }
      if (
        templateWordNamespaces.has(tag.uri) &&
        (tag.local === "p" ||
          tag.local === "tbl" ||
          (control.contentDepth === 0 &&
            (tag.local === "altChunk" ||
              tag.local === "customXml" ||
              tag.local === "oMathPara")))
      ) {
        control.contentIsBlock = true;
      }
    }
    if (hasContentControl && tag.local !== "sdt") {
      preserveNamespaces(
        start,
        parser.position,
        activeScalarNamespaceDeclarations(),
        new Set(Object.keys(tag.attributes))
      );
    }
    for (const control of controls) {
      if (control.contentDepth !== null) {
        control.contentDepth += 1;
      }
    }
    if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      const parent = elementStack.at(-1);
      controls.push({
        contentCloseStart: null,
        contentDepth: null,
        contentHasContent: false,
        contentIsBlock: false,
        contentOpenEnd: null,
        contentSelfClosing: false,
        namespaceAttributeNames: new Set(Object.keys(tag.attributes)),
        namespaceDeclarations: Object.values(tag.attributes)
          .filter((attribute) => attribute.uri === templateXmlnsNamespace)
          .map((attribute) => ({
            name: attribute.name,
            value: attribute.value,
          })),
        parentContentControls,
        parentIsBlockContainer:
          parent !== undefined &&
          templateWordNamespaces.has(parent.uri) &&
          (parent.local === "body" ||
            parent.local === "comment" ||
            parent.local === "docPartBody" ||
            parent.local === "endnote" ||
            parent.local === "ftr" ||
            parent.local === "footnote" ||
            parent.local === "hdr" ||
            parent.local === "txbxContent"),
        scalarControl: false,
        start,
        startEnd: parser.position,
        wordNamespace: tag.uri,
        wordPrefix: tag.name.includes(":")
          ? tag.name.slice(0, tag.name.indexOf(":"))
          : "",
      });
    } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      const control = controls.at(-1);
      if (control) {
        propertyControls.push(control);
      }
    } else if (
      propertyControls.length > 0 &&
      ((tag.local === "checkbox" && tag.uri === templateCheckboxNamespace) ||
        (templateWordNamespaces.has(tag.uri) &&
          (tag.local === "comboBox" ||
            tag.local === "date" ||
            tag.local === "dropDownList")))
    ) {
      const control = propertyControls.at(-1);
      if (control) {
        control.scalarControl = true;
      }
    } else if (
      tag.local === "sdtContent" &&
      templateWordNamespaces.has(tag.uri)
    ) {
      const control = controls.at(-1);
      if (control) {
        control.contentOpenEnd = parser.position;
        control.contentSelfClosing = tag.isSelfClosing;
        control.contentDepth = 0;
        for (const attribute of Object.values(tag.attributes)) {
          if (attribute.uri === templateXmlnsNamespace) {
            control.namespaceDeclarations.push({
              name: attribute.name,
              value: attribute.value,
            });
          }
        }
      }
    }
    elementStack.push({ local: tag.local, uri: tag.uri });
  });
  parser.on("closetag", (tag) => {
    if (tag.local === "sdtContent" && templateWordNamespaces.has(tag.uri)) {
      const control = controls.at(-1);
      if (control && control.contentOpenEnd !== null) {
        control.contentCloseStart = xml.lastIndexOf("<", parser.position - 1);
        control.contentDepth = null;
      }
    } else if (tag.local === "sdtPr" && templateWordNamespaces.has(tag.uri)) {
      propertyControls.pop();
    } else if (tag.local === "sdt" && templateWordNamespaces.has(tag.uri)) {
      const control = controls.pop();
      if (control?.contentIsBlock && control.parentContentControls) {
        for (const parentControl of control.parentContentControls) {
          parentControl.contentIsBlock = true;
        }
      }
      if (
        control?.scalarControl &&
        control.contentOpenEnd !== null &&
        control.contentCloseStart !== null
      ) {
        if (control.contentSelfClosing) {
          patches.push({
            end: parser.position,
            replacement: "",
            start: control.start,
          });
        } else {
          const needsParagraph =
            control.parentIsBlockContainer &&
            control.contentHasContent &&
            !control.contentIsBlock;
          const paragraphName = control.wordPrefix
            ? `${control.wordPrefix}:p`
            : "p";
          const namespaceAttribute = control.wordPrefix
            ? `xmlns:${control.wordPrefix}`
            : "xmlns";
          const paragraphOpen = needsParagraph
            ? `<${paragraphName} ${namespaceAttribute}="${htmlEscape(control.wordNamespace)}">`
            : "";
          const paragraphClose = needsParagraph ? `</${paragraphName}>` : "";
          patches.push(
            {
              end: control.contentOpenEnd,
              replacement: paragraphOpen,
              start: control.start,
            },
            {
              end: parser.position,
              replacement: paragraphClose,
              start: control.contentCloseStart,
            }
          );
        }
      } else if (control) {
        preserveNamespaces(
          control.start,
          control.startEnd,
          activeScalarNamespaceDeclarations(),
          control.namespaceAttributeNames
        );
      }
    }
    for (const control of controls) {
      if (control.contentDepth !== null && control.contentDepth > 0) {
        control.contentDepth -= 1;
      }
    }
    elementStack.pop();
  });
  try {
    parser.write(xml).close();
  } catch {
    fail(422, "invalid_template", "The DOCX package contains invalid XML");
  }
  let transformed = xml;
  for (const patch of patches.toSorted(
    (left, right) => right.start - left.start
  )) {
    transformed =
      transformed.slice(0, patch.start) +
      patch.replacement +
      transformed.slice(patch.end);
  }
  return transformed;
}

async function nativePdfDocument(documentKey: string): Promise<Uint8Array> {
  const sourceObjectKey = await operationDocumentKey(documentKey);
  if (!sourceObjectKey || !(await objectExists(sourceObjectKey))) {
    fail(404, "not_found", "Document was not found");
  }
  const { archive, xmlPaths } = safeTemplateArchive(
    await readObject(sourceObjectKey)
  );
  const controlParts = reachableTemplateControlParts(archive, xmlPaths);
  for (const archivePath of controlParts) {
    const originalXml = templateArchiveText(archive, archivePath);
    const transformedXml = nativePdfScalarControlXml(originalXml);
    if (transformedXml !== originalXml) {
      archive[archivePath] = strToU8(
        transformedXml.replace(
          /encoding=(?<quote>["'])UTF-16(?:LE|BE)?\k<quote>/iu,
          'encoding="UTF-8"'
        )
      );
    }
  }
  return zipSync(archive);
}
async function nativeResponseDocument(
  publishedTemplateId: string,
  data: JsonRecord,
  baseDocumentBytes?: Uint8Array,
  pictures: ReadonlyMap<string, NativePictureUpload> = new Map()
): Promise<Uint8Array> {
  const [publishedTemplate, manifest] = await Promise.all([
    prisma.publishedTemplate.findUnique({
      select: { objectKey: true },
      where: { id: publishedTemplateId },
    }),
    prisma.fieldManifest.findUnique({
      include: {
        fields: {
          select: {
            options: true,
            pictureMaxBytes: true,
            pictureMaxHeight: true,
            pictureMaxWidth: true,
            required: true,
            tag: true,
            type: true,
          },
        },
      },
      where: { publishedTemplateId },
    }),
  ]);
  if (!publishedTemplate?.objectKey || !manifest) {
    fail(409, "document_unavailable", "The published document is unavailable");
  }
  const scalarTags = new Set(
    manifest.fields
      .filter(({ type }) => type !== FieldType.picture)
      .map(({ tag }) => tag)
  );
  const baseBytes =
    baseDocumentBytes ?? (await readObject(publishedTemplate.objectKey));
  return overlayNativeResponsePictures(
    overlayNativeResponseDocument(baseBytes, manifest.fields, data, scalarTags),
    manifest.fields,
    pictures
  );
}

async function completeDraftOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  if (!documentKey) {
    fail(500, "invalid_operation", "Draft operation has no document key");
  }
  if (!metadata.responseId || !metadata.data) {
    fail(500, "invalid_operation", "Draft metadata is incomplete");
  }
  const response = await prisma.response.findUnique({
    where: { id: metadata.responseId },
  });
  if (
    !response ||
    response.status !== ResponseStatus.draft ||
    response.draftDocumentKey !== documentKey
  ) {
    fail(409, "stale_operation", "The response is no longer editable");
  }
  await validateResponseDocument(response.publishedTemplateId, bytes, false);
  const nextDocumentKey = metadata.nextDocumentKey ?? documentKey;
  const result = {
    documentKey: nextDocumentKey,
    publicId: metadata.publicId,
    responseId: response.id,
  };
  const cleanupObjectKeys = [
    metadata.stagedObjectKey,
    ...(response.draftObjectKey ? [response.draftObjectKey] : []),
  ];
  await prisma.$transaction(
    async (tx) => {
      const updated = await tx.response.updateMany({
        data: {
          draftData: jsonValue(metadata.data),
          draftDocumentKey: nextDocumentKey,
          draftObjectKey: metadata.finalObjectKey,
          updatedAt: new Date(),
        },
        where: {
          draftDocumentKey: documentKey,
          id: response.id,
          status: ResponseStatus.draft,
        },
      });
      if (updated.count !== 1) {
        fail(409, "stale_operation", "The response is no longer editable");
      }
      await markOperationCompleted(
        tx,
        operation.id,
        result,
        metadata,
        cleanupObjectKeys
      );
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}

async function completeSubmitOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  if (!documentKey) {
    fail(500, "invalid_operation", "Submit operation has no document key");
  }
  if (!metadata.responseId || !metadata.submissionId || !metadata.data) {
    fail(500, "invalid_operation", "Submit metadata is incomplete");
  }
  const { responseId, submissionId, data } = metadata;
  const response = await prisma.response.findUnique({
    where: { id: responseId },
  });
  if (
    !response ||
    response.status !== ResponseStatus.submitting ||
    response.draftDocumentKey !== documentKey
  ) {
    fail(
      409,
      "stale_operation",
      "The response is no longer pending submission"
    );
  }
  await validateResponseDocument(response.publishedTemplateId, bytes, true);

  const submissionDocumentKey = metadata.submissionDocumentKey ?? documentKey;
  const result = {
    publicId: metadata.publicId,
    responseId: response.id,
    submissionId,
  };
  const cleanupObjectKeys = [
    metadata.stagedObjectKey,
    ...(response.draftObjectKey ? [response.draftObjectKey] : []),
  ];
  await prisma.$transaction(
    async (tx) => {
      const claimed = await tx.operation.updateMany({
        data: { updatedAt: new Date() },
        where: { id: operation.id, status: OperationStatus.processing },
      });
      if (claimed.count !== 1) {
        fail(
          409,
          "stale_operation",
          "The submission operation is no longer active"
        );
      }
      await tx.submission.create({
        data: {
          data: jsonValue(data),
          documentKey: submissionDocumentKey,
          form: { connect: { id: response.formId } },
          id: submissionId,
          objectKey: metadata.finalObjectKey,
          owner: { connect: { id: response.userId } },
          response: { connect: { id: response.id } },
        },
      });
      const updatedResponse = await tx.response.updateMany({
        data: {
          draftData: Prisma.DbNull,
          draftDocumentKey: null,
          draftObjectKey: null,
          status: ResponseStatus.submitted,
          updatedAt: new Date(),
        },
        where: { id: response.id, status: ResponseStatus.submitting },
      });
      if (updatedResponse.count !== 1) {
        fail(
          409,
          "stale_operation",
          "The response is no longer pending submission"
        );
      }
      await markOperationCompleted(
        tx,
        operation.id,
        result,
        metadata,
        cleanupObjectKeys,
        submissionId
      );
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}
async function processNativeResponseOperation(
  operation: Operation,
  storeObject: typeof putObject
): Promise<void> {
  const metadata = operationMetadata(operation.metadata);
  try {
    if (
      (metadata.action !== "save-draft" && metadata.action !== "submit") ||
      !operation.documentKey ||
      !metadata.responseId ||
      !metadata.data ||
      !metadata.finalObjectKey ||
      !metadata.stagedObjectKey
    ) {
      fail(500, "invalid_operation", "Native response operation is incomplete");
    }
    const claimed = await prisma.operation.updateMany({
      data: { status: OperationStatus.processing, updatedAt: new Date() },
      where: { id: operation.id, status: OperationStatus.pending },
    });
    if (claimed.count !== 1) {
      return;
    }
    const response = await prisma.response.findUnique({
      select: {
        draftDocumentKey: true,
        publishedTemplateId: true,
        status: true,
      },
      where: { id: metadata.responseId },
    });
    const expectedStatus =
      metadata.action === "submit"
        ? ResponseStatus.submitting
        : ResponseStatus.draft;
    if (
      !response ||
      response.status !== expectedStatus ||
      response.draftDocumentKey !== operation.documentKey
    ) {
      fail(409, "stale_operation", "The response is no longer editable");
    }
    const bytes = metadata.nativeDocumentStaged
      ? await readObject(metadata.stagedObjectKey)
      : await nativeResponseDocument(
          response.publishedTemplateId,
          metadata.data
        );
    if (!metadata.nativeDocumentStaged) {
      await storeObject(metadata.stagedObjectKey, bytes, DOCX_CONTENT_TYPE);
    }
    await storeObject(metadata.finalObjectKey, bytes, DOCX_CONTENT_TYPE);
    const completion =
      metadata.action === "submit"
        ? await completeSubmitOperation(operation, metadata, bytes)
        : await completeDraftOperation(operation, metadata, bytes);
    await deleteObjects(completion.cleanupObjectKeys);
  } catch (error) {
    await deleteObjects([metadata.stagedObjectKey]);
    if (metadata.finalObjectKey) {
      await deleteObjectUnlessCanonical(metadata.finalObjectKey);
    }
    await updateOperationFailed(
      operation.id,
      error instanceof HttpError && error.code === "invalid_template"
        ? "invalid_template"
        : "document_save_failed"
    );
  }
}
async function completeCorrectionOperation(
  operation: Operation,
  metadata: OperationMetadata,
  bytes: Uint8Array
): Promise<OperationCompletion> {
  const { documentKey } = operation;
  const {
    baseDocumentKey,
    baseRevision,
    data,
    nextDocumentKey,
    reason,
    responseId,
    submissionId,
  } = metadata;
  const actorId = operation.actorId;
  if (
    !documentKey ||
    !responseId ||
    !submissionId ||
    !data ||
    typeof baseDocumentKey !== "string" ||
    typeof baseRevision !== "number" ||
    !nextDocumentKey ||
    !reason ||
    !actorId
  ) {
    fail(500, "invalid_operation", "Correction metadata is incomplete");
  }
  const response = await prisma.response.findUnique({
    include: { submission: true },
    where: { id: responseId },
  });
  const submission = response?.submission;
  if (
    !response ||
    response.status !== ResponseStatus.submitted ||
    !submission ||
    submission.id !== submissionId
  ) {
    fail(409, "stale_operation", "The Submission is no longer correctable");
  }
  await validateResponseDocument(response.publishedTemplateId, bytes, true);
  const cleanupObjectKeys = [
    metadata.stagedObjectKey,
    metadata.workspaceObjectKey,
  ].filter((key): key is string => Boolean(key));
  await prisma.$transaction(
    async (tx) => {
      const [lockedResponse] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "responses"
          WHERE "id" = ${response.id}::uuid
          FOR UPDATE
        `
      );
      if (!lockedResponse) {
        fail(409, "stale_operation", "The Submission is no longer correctable");
      }
      const current = await tx.response.findUnique({
        include: {
          corrections: {
            orderBy: { revision: "desc" },
            take: 1,
          },
          submission: true,
        },
        where: { id: response.id },
      });
      const currentSubmission = current?.submission;
      const latest = current?.corrections[0];
      const currentRevision = latest?.revision ?? 0;
      const currentDocumentKey =
        latest?.documentKey ?? currentSubmission?.documentKey;
      if (
        !current ||
        current.status !== ResponseStatus.submitted ||
        !currentSubmission ||
        currentSubmission.id !== submissionId ||
        currentRevision !== baseRevision ||
        currentDocumentKey !== baseDocumentKey
      ) {
        fail(409, "stale_operation", "A newer Correction is already effective");
      }
      const previousData = jsonRecord(latest?.data ?? currentSubmission.data);
      const correction = await tx.correction.create({
        data: {
          actorId,
          changedData: jsonValue(changedResponseData(previousData, data)),
          data: jsonValue(data),
          documentKey: nextDocumentKey,
          objectKey: metadata.finalObjectKey,
          reason,
          responseId: current.id,
          revision: currentRevision + 1,
          submissionId,
        },
      });
      await tx.response.update({
        data: { updatedAt: new Date() },
        where: { id: current.id },
      });
      await tx.editorLease.updateMany({
        data: {
          workspaceBaseDocumentKey: null,
          workspaceBaseRevision: null,
          workspaceDocumentKey: null,
          workspaceObjectKey: null,
        },
        where: {
          targetId: current.id,
          targetType: OperationTargetType.correction,
          workspaceDocumentKey: metadata.workspaceDocumentKey,
        },
      });
      const completedResult = {
        correctionId: correction.id,
        publicId: metadata.publicId,
        responseId: current.id,
        revision: correction.revision,
        submissionId,
      };
      await markOperationCompleted(
        tx,
        operation.id,
        completedResult,
        metadata,
        cleanupObjectKeys
      );
      await tx.auditEvent.create({
        data: {
          action: "create_correction",
          actorId,
          outcome: AuditOutcome.success,
          safeMetadata: jsonValue({
            revision: correction.revision,
            state: "submitted",
          }),
          targetId: correction.id,
          targetType: "correction",
        },
      });
      return completedResult;
    },
    { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
  );
  return { cleanupObjectKeys };
}

async function finalizeCallback(
  operationId: string,
  payload: CallbackPayload,
  snapshot: Uint8Array | undefined,
  allowedCallbackOrigins: ReadonlySet<string>,
  maximumBytes = maxCallbackDocumentBytes
): Promise<void> {
  const operation = await prisma.operation.findUnique({
    where: { id: operationId },
  });
  if (!operation) {
    return;
  }
  if (
    operation.status === OperationStatus.completed ||
    operation.status === OperationStatus.failed
  ) {
    await cleanupTerminalOperationObjects(operation);
    return;
  }
  const metadata = operationMetadata(operation.metadata);
  const callbackUrl = resolveCallbackDocumentUrl(
    payload.url,
    allowedCallbackOrigins
  );
  if (
    typeof payload.key !== "string" ||
    payload.key !== operation.documentKey
  ) {
    await updateOperationFailed(operation.id, "callback_key_mismatch");
    return;
  }
  if (!snapshot && !callbackUrl) {
    await updateOperationFailed(operation.id, "callback_document_unavailable");
    return;
  }

  let bytes = snapshot;
  if (!bytes) {
    if (!callbackUrl) {
      throw new Error("ONLYOFFICE callback document URL is unavailable");
    }
    bytes = await readCallbackDocument(callbackUrl, maximumBytes);
  }
  const claimed = await prisma.operation.updateMany({
    data: { status: OperationStatus.processing, updatedAt: new Date() },
    where: {
      id: operation.id,
      status: { in: [OperationStatus.pending, OperationStatus.processing] },
    },
  });
  if (claimed.count !== 1) {
    return;
  }
  let completion: OperationCompletion | undefined;

  try {
    if (
      (metadata.action === "save-draft" ||
        metadata.action === "save-correction" ||
        metadata.action === "submit") &&
      metadata.responseId &&
      metadata.data
    ) {
      const response = await prisma.response.findUnique({
        select: { publishedTemplateId: true },
        where: { id: metadata.responseId },
      });
      if (response) {
        const [publishedTemplate, manifest] = await Promise.all([
          prisma.publishedTemplate.findUnique({
            select: { objectKey: true },
            where: { id: response.publishedTemplateId },
          }),
          prisma.fieldManifest.findUnique({
            include: {
              fields: { select: { options: true, tag: true, type: true } },
            },
            where: { publishedTemplateId: response.publishedTemplateId },
          }),
        ]);
        if (publishedTemplate?.objectKey && manifest) {
          const templateBytes = await readObject(publishedTemplate.objectKey);
          if (
            await supportsNativeTemplate(
              publishedTemplate.objectKey,
              manifest.fields,
              templateBytes
            )
          ) {
            const callbackValues =
              metadata.callbackEmptyFieldTags &&
              metadata.callbackEmptyFieldTags.length > 0
                ? responseTextControlValues(bytes)
                : null;
            const callbackMetadata =
              callbackValues === null
                ? null
                : onlyOfficeFieldDisplayMetadata(bytes);
            const serverHeldFieldTags = new Set(metadata.serverHeldFieldTags);
            const callbackEmptyFieldTags = new Set(
              metadata.callbackEmptyFieldTags
            );
            for (const field of manifest.fields) {
              if (
                !callbackEmptyFieldTags.has(field.tag) ||
                serverHeldFieldTags.has(field.tag)
              ) {
                continue;
              }
              const callbackValue = callbackValues?.get(field.tag);
              if (callbackValue === undefined || !callbackMetadata) {
                continue;
              }
              const normalizedValue = normalizeOnlyOfficeDisplayValue(
                field,
                callbackValue,
                callbackMetadata.get(field.tag)
              );
              if (manifestFieldValueMatches(field, normalizedValue)) {
                metadata.data[field.tag] = normalizedValue;
              }
            }
            bytes = await overlayNativeResponseDocument(
              bytes,
              manifest.fields,
              metadata.data,
              serverHeldFieldTags
            );
          }
        }
      }
    }

    if (metadata.action === "save-template") {
      validateTemplatePackage(bytes);
    } else if (metadata.action === "publish") {
      validateTemplateControls(bytes);
    } else {
      validateOfficeRelationshipsBytes(bytes);
    }
    await putObject(metadata.stagedObjectKey, bytes, DOCX_CONTENT_TYPE);
    await putObject(metadata.finalObjectKey, bytes, DOCX_CONTENT_TYPE);

    if (metadata.action === "save-template") {
      completion = await completeTemplateOperation(operation, metadata, bytes);
    } else if (metadata.action === "publish") {
      completion = await completePublishOperation(operation, metadata, bytes);
    } else if (metadata.action === "save-draft") {
      completion = await completeDraftOperation(operation, metadata, bytes);
    } else if (metadata.action === "save-correction") {
      completion = await completeCorrectionOperation(
        operation,
        metadata,
        bytes
      );
    } else {
      completion = await completeSubmitOperation(operation, metadata, bytes);
    }
  } catch (error) {
    await deleteObjects([metadata.stagedObjectKey]);
    await deleteObjectUnlessCanonical(metadata.finalObjectKey);
    throw error;
  }

  await deleteObjects(completion?.cleanupObjectKeys ?? []);
}

type FieldRulePolicy = "editable" | "lock-when-available";
interface FieldRuleInput {
  documentKey: string;
  previousTag: string | null;
  prefillPointer: string | null;
  prefillPolicy: FieldRulePolicy;
  required: boolean;
  tag: string;
}

function fieldRuleDto(rule: {
  prefillPointer: string | null;
  prefillPolicy: PrefillPolicy;
  required: boolean;
  tag: string;
}): {
  prefillPointer: string | null;
  prefillPolicy: FieldRulePolicy;
  required: boolean;
  tag: string;
} {
  return {
    prefillPointer: rule.prefillPointer,
    prefillPolicy:
      rule.prefillPolicy === PrefillPolicy.lock_when_available
        ? "lock-when-available"
        : "editable",
    required: rule.required,
    tag: rule.tag,
  };
}

function fieldRuleTag(value: unknown, field: string): string {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.trim().length === 0 ||
    value.length > fieldTagMaximumLength
  ) {
    fail(400, "invalid_field_selection", `${field} is invalid`);
  }
  return value;
}

function fieldRuleInput(input: JsonRecord): FieldRuleInput {
  const expectedKeys = new Set([
    "documentKey",
    "previousTag",
    "tag",
    "required",
    "prefillPointer",
    "prefillPolicy",
  ]);
  const keys = Object.keys(input);
  if (
    keys.length !== expectedKeys.size ||
    keys.some((key) => !expectedKeys.has(key))
  ) {
    fail(
      400,
      "invalid_field_config",
      "documentKey, previousTag, tag, required, prefillPointer, and prefillPolicy are required"
    );
  }
  if (
    typeof input.documentKey !== "string" ||
    input.documentKey.trim().length === 0
  ) {
    fail(400, "invalid_field_config", "documentKey is required");
  }
  if (typeof input.required !== "boolean") {
    fail(400, "invalid_field_config", "required must be a boolean");
  }
  const tag = fieldRuleTag(input.tag, "tag");
  const previousTag =
    input.previousTag === null
      ? null
      : fieldRuleTag(input.previousTag, "previousTag");
  let prefillPointer: string | null;
  if (input.prefillPointer === null) {
    prefillPointer = null;
  } else if (
    typeof input.prefillPointer !== "string" ||
    input.prefillPointer.length === 0 ||
    input.prefillPointer.length > fieldPointerMaximumLength
  ) {
    fail(400, "invalid_field_config", "prefillPointer is invalid");
  } else {
    prefillPointer = input.prefillPointer;
  }
  if (
    input.prefillPolicy !== "editable" &&
    input.prefillPolicy !== "lock-when-available"
  ) {
    fail(400, "invalid_field_config", "prefillPolicy is invalid");
  }
  if (
    input.prefillPolicy === "lock-when-available" &&
    prefillPointer === null
  ) {
    fail(
      400,
      "invalid_field_config",
      "lock-when-available requires prefillPointer"
    );
  }
  return {
    documentKey: input.documentKey,
    prefillPointer,
    prefillPolicy: input.prefillPolicy,
    previousTag,
    required: input.required,
    tag,
  };
}

function fieldRulePrefillPolicy(policy: FieldRulePolicy): PrefillPolicy {
  return policy === "lock-when-available"
    ? PrefillPolicy.lock_when_available
    : PrefillPolicy.editable;
}

function requireTemplateDraft(
  form: FormWithDocuments,
  missingMessage = "No template DOCX is configured"
): TemplateDraft {
  if (!form.templateDraft) {
    fail(409, "document_unavailable", missingMessage);
  }
  return form.templateDraft;
}

function fieldRuleCapabilityScope(
  form: FormWithDocuments,
  templateDraft: TemplateDraft
): Omit<EditorCapabilityScope, "action" | "operationId"> {
  return {
    documentKey: templateDraft.documentKey,
    formId: form.id,
    targetId: templateDraft.id,
    targetType: "template-draft",
  };
}

function validateFieldRulePointer(prefillPointer: string | null): void {
  if (
    prefillPointer !== null &&
    !externalSchemaItems.some((item) => item.pointer === prefillPointer)
  ) {
    fail(
      400,
      "invalid_field_selection",
      "prefillPointer is not a selectable schema field"
    );
  }
}
async function findFormByPublicId(
  publicId: string
): Promise<FormWithDocuments> {
  if (!publicIdPattern.test(publicId)) {
    fail(404, "not_found", "Form was not found");
  }
  const form = await prisma.form.findUnique({
    include: { publishedTemplate: true, templateDraft: true },
    where: { publicId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  return form;
}

interface PrefillHandoffLaunch {
  claimToken: string;
  publicId: string;
}

interface PrefillHandoffRedeemResult {
  cleanupObjectKeys: string[];
  response: ResponseWithSnapshot;
}

function publishedPrefillConfiguration(
  form: {
    publishedTemplate: {
      contentHash: string;
      id: string;
      prefillConfiguration: {
        configurationHash: string;
        fields: { pointer: string; policy: PrefillPolicy; tag: string }[];
        publishedTemplateId: string | null;
      } | null;
    } | null;
    status: FormStatus;
  } | null
): {
  configurationHash: string;
  fields: { pointer: string; policy: PrefillPolicy; tag: string }[];
  publishedTemplateId: string;
  templateId: string;
} | null {
  const template = form?.publishedTemplate;
  const configuration = template?.prefillConfiguration;
  if (
    !template ||
    !configuration ||
    configuration.fields.length === 0 ||
    configuration.publishedTemplateId !== template.id ||
    configuration.configurationHash !== template.contentHash ||
    form.status !== FormStatus.published
  ) {
    return null;
  }
  return {
    configurationHash: configuration.configurationHash,
    fields: configuration.fields,
    publishedTemplateId: configuration.publishedTemplateId,
    templateId: template.id,
  };
}

async function createPrefillHandoff(
  input: PrefillHandoffCreateInput
): Promise<{ code: string; launchPath: string }> {
  if (!publicIdPattern.test(input.publicId)) {
    fail(404, "not_found", "Form was not found");
  }
  const form = await prisma.form.findUnique({
    include: {
      publishedTemplate: {
        include: {
          manifest: { include: { fields: true } },
          prefillConfiguration: { include: { fields: true } },
        },
      },
    },
    where: { publicId: input.publicId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  const configuration = publishedPrefillConfiguration(form);
  if (!configuration) {
    fail(404, "not_found", "Form was not found");
  }
  const filteredValues = filteredPrefillValues(
    input.values,
    configuration.fields
  );
  const manifest = form.publishedTemplate?.manifest;
  if (
    !manifest ||
    manifest.configurationHash !== configuration.configurationHash
  ) {
    handoffUnavailable();
  }
  validatePrefillValuesAgainstManifest(filteredValues, manifest.fields);
  const code = randomBytes(32).toString("base64url");
  const codeDigest = tokenDigest(code);
  const externalReferenceDigest = tokenDigest(input.externalReference);
  const expiresAt = new Date(Date.now() + handoffCodeLifetimeMs);

  try {
    await prisma.$transaction(
      async (tx) => {
        const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "forms"
            WHERE "id" = ${form.id}::uuid
            FOR UPDATE
          `
        );
        if (!lockedForm) {
          fail(404, "not_found", "Form was not found");
        }
        const current = await tx.form.findUnique({
          include: {
            publishedTemplate: {
              include: {
                prefillConfiguration: { include: { fields: true } },
              },
            },
          },
          where: { id: lockedForm.id },
        });
        const currentConfiguration = publishedPrefillConfiguration(current);
        if (!current || !currentConfiguration) {
          fail(404, "not_found", "Form was not found");
        }
        if (
          currentConfiguration.configurationHash !==
          configuration.configurationHash
        ) {
          fail(404, "not_found", "Form was not found");
        }
        const deletedReference = await tx.deletionTombstone.findUnique({
          select: { id: true },
          where: { externalReferenceDigest },
        });
        if (deletedReference) {
          handoffUnavailable();
        }
        const existingReference = await tx.handoff.findUnique({
          select: { id: true },
          where: { externalReferenceDigest },
        });
        if (existingReference) {
          handoffUnavailable();
        }
        await tx.handoff.create({
          data: {
            codeDigest,
            configurationHash: currentConfiguration.configurationHash,
            expiresAt,
            externalReferenceDigest,
            filteredValues: jsonValue(filteredValues),
            form: { connect: { id: current.id } },
            id: crypto.randomUUID(),
            normalizedEmail: input.email,
          },
        });
        await createFormAudit(tx, {
          action: "create_handoff",
          actorId: null,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: current.publicId,
        });
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
  } catch (error) {
    try {
      await createFormFailureAudit({
        action: "create_handoff",
        actorId: null,
        error,
        targetId: publicIdPattern.test(input.publicId) ? input.publicId : null,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw error;
  }
  return { code, launchPath: "/prefill/handoff" };
}
async function launchPrefillHandoff(
  code: string,
  clock: () => Date = () => new Date()
): Promise<PrefillHandoffLaunch> {
  const codeDigest = tokenDigest(code);
  let auditTargetId: string | null = null;
  try {
    const launch = await prisma.$transaction(
      async (tx) => {
        const [lockedHandoff] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "handoffs"
            WHERE "code_digest" = ${codeDigest}
            FOR UPDATE
          `
        );
        if (!lockedHandoff) {
          handoffUnavailable();
        }
        const handoff = await tx.handoff.findUnique({
          where: { id: lockedHandoff.id },
        });
        const form = handoff?.formId
          ? await tx.form.findUnique({
              include: {
                publishedTemplate: {
                  include: {
                    prefillConfiguration: { include: { fields: true } },
                  },
                },
              },
              where: { id: handoff.formId },
            })
          : null;
        auditTargetId =
          form && publicIdPattern.test(form.publicId) ? form.publicId : null;
        const configuration = publishedPrefillConfiguration(form);
        const now = clock();
        if (
          !handoff ||
          !form ||
          !configuration ||
          handoff.status !== HandoffStatus.pending ||
          handoff.expiresAt <= now
        ) {
          handoffUnavailable();
        }
        const claimToken = randomBytes(32).toString("base64url");
        const claimExpiresAt = new Date(
          now.getTime() + pendingClaimLifetimeSeconds * 1000
        );
        await tx.pendingClaim.create({
          data: {
            claimDigest: tokenDigest(claimToken),
            expiresAt: claimExpiresAt,
            handoff: { connect: { id: handoff.id } },
            id: crypto.randomUUID(),
          },
        });
        const reserved = await tx.handoff.updateMany({
          data: {
            expiresAt: claimExpiresAt,
            reservedAt: now,
            status: HandoffStatus.reserved,
          },
          where: {
            expiresAt: { gt: now },
            id: handoff.id,
            status: HandoffStatus.pending,
          },
        });
        if (reserved.count !== 1) {
          handoffUnavailable();
        }
        await createFormAudit(tx, {
          action: "launch_handoff",
          actorId: null,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: form.publicId,
        });
        return { claimToken, publicId: form.publicId };
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
    return launch;
  } catch (error) {
    try {
      await createFormFailureAudit({
        action: "launch_handoff",
        actorId: null,
        error,
        targetId: auditTargetId,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw error;
  }
}

async function redeemPrefillHandoff(
  form: FormWithDocuments,
  identity: Identity,
  claimToken: string,
  clock: () => Date = () => new Date()
): Promise<PrefillHandoffRedeemResult> {
  const responseHint = await prisma.response.findUnique({
    select: { id: true, status: true },
    where: { formId_userId: { formId: form.id, userId: identity.id } },
  });
  if (responseHint?.status === ResponseStatus.submitted) {
    handoffUnavailable();
  }
  const responseId = responseHint?.id ?? crypto.randomUUID();
  const publishedTemplate = form.publishedTemplate;
  if (!publishedTemplate?.objectKey || !publishedTemplate.documentKey) {
    handoffUnavailable();
  }
  if (!(await objectExists(publishedTemplate.objectKey))) {
    handoffUnavailable();
  }
  const document = await readObject(publishedTemplate.objectKey);
  const draftObjectKey = objectKey(
    "responses",
    responseId,
    "draft",
    crypto.randomUUID(),
    "docx"
  );
  const draftDocumentKey = `response-${responseId}-${crypto.randomUUID()}`;
  await putObject(draftObjectKey, document, DOCX_CONTENT_TYPE);
  const claimDigest = tokenDigest(claimToken);
  let oldDraftObjectKey: string | null = null;
  let unusedDraftObjectKey: string | null = null;
  try {
    const response = await prisma.$transaction(
      async (tx) => {
        const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "forms"
            WHERE "id" = ${form.id}::uuid
            FOR UPDATE
          `
        );
        if (!lockedForm) {
          handoffUnavailable();
        }
        const currentForm = await tx.form.findUnique({
          include: {
            publishedTemplate: {
              include: {
                prefillConfiguration: { include: { fields: true } },
              },
            },
          },
          where: { id: lockedForm.id },
        });
        const configuration = publishedPrefillConfiguration(currentForm);
        if (
          !currentForm ||
          currentForm.publicId !== form.publicId ||
          !configuration ||
          !currentForm.publishedTemplate
        ) {
          handoffUnavailable();
        }
        const [lockedClaim] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "pending_claims"
            WHERE "claim_digest" = ${claimDigest}
            FOR UPDATE
          `
        );
        if (!lockedClaim) {
          handoffUnavailable();
        }
        const pendingClaim = await tx.pendingClaim.findUnique({
          include: { handoff: true },
          where: { id: lockedClaim.id },
        });
        const handoff = pendingClaim?.handoff;
        const now = clock();
        if (
          !pendingClaim ||
          !handoff ||
          pendingClaim.consumedAt ||
          pendingClaim.expiresAt <= now ||
          handoff.formId !== currentForm.id ||
          handoff.status !== HandoffStatus.reserved ||
          handoff.consumedAt ||
          handoff.expiresAt <= now ||
          handoff.normalizedEmail !== normalizeEmail(identity.email) ||
          handoff.configurationHash !== configuration.configurationHash ||
          configuration.publishedTemplateId !==
            currentForm.publishedTemplate.id ||
          configuration.configurationHash !==
            currentForm.publishedTemplate.contentHash
        ) {
          handoffUnavailable();
        }
        const deletedReference = await tx.deletionTombstone.findUnique({
          select: { id: true },
          where: {
            externalReferenceDigest: handoff.externalReferenceDigest,
          },
        });
        if (deletedReference) {
          handoffUnavailable();
        }
        const currentResponse = await tx.response.findUnique({
          include: { prefillSnapshot: true, submission: true },
          where: {
            formId_userId: { formId: currentForm.id, userId: identity.id },
          },
        });
        if (
          currentResponse?.status === ResponseStatus.submitted ||
          currentResponse?.status === ResponseStatus.submitting ||
          (currentResponse &&
            (await tx.operation.findFirst({
              select: { id: true },
              where: {
                responseId: currentResponse.id,
                status: {
                  in: [OperationStatus.pending, OperationStatus.processing],
                },
              },
            })))
        ) {
          handoffUnavailable();
        }
        const reuseExistingDraft =
          currentResponse?.status === ResponseStatus.draft &&
          currentResponse.publishedVersion === currentForm.version &&
          currentResponse.draftDocumentKey !== null &&
          currentResponse.draftObjectKey !== null;
        const storedValues = jsonRecord(
          handoff.filteredValues,
          "The prefill handoff values are invalid"
        );
        const values: JsonRecord = {};
        const lockedFields: JsonRecord = {};
        for (const field of configuration.fields) {
          const value = storedValues[field.tag];
          if (value === undefined) {
            continue;
          }
          if (!externalValueMatchesSchema(field.pointer, value)) {
            handoffUnavailable();
          }
          values[field.tag] = value;
          lockedFields[field.tag] =
            field.policy === PrefillPolicy.lock_when_available;
        }
        const responseTargetId = currentResponse?.id ?? responseId;
        if (reuseExistingDraft) {
          unusedDraftObjectKey = draftObjectKey;
        } else {
          oldDraftObjectKey = currentResponse?.draftObjectKey ?? null;
          if (currentResponse) {
            await tx.editorLease.deleteMany({
              where: {
                targetId: currentResponse.id,
                targetType: OperationTargetType.response,
              },
            });
            await tx.operation.deleteMany({
              where: { responseId: currentResponse.id },
            });
            await tx.prefillSnapshot.deleteMany({
              where: { responseId: currentResponse.id },
            });
            await tx.response.update({
              data: {
                draftData: Prisma.DbNull,
                draftDocumentKey,
                draftObjectKey,
                externalReferenceDigest: handoff.externalReferenceDigest,
                publishedTemplateId: currentForm.publishedTemplate.id,
                publishedVersion: currentForm.version,
                status: ResponseStatus.draft,
                updatedAt: now,
              },
              where: { id: responseTargetId },
            });
          } else {
            await tx.response.create({
              data: {
                draftData: Prisma.DbNull,
                draftDocumentKey,
                draftObjectKey,
                externalReferenceDigest: handoff.externalReferenceDigest,
                form: { connect: { id: currentForm.id } },
                id: responseTargetId,
                owner: { connect: { id: identity.id } },
                publishedTemplate: {
                  connect: { id: currentForm.publishedTemplate.id },
                },
                publishedVersion: currentForm.version,
                status: ResponseStatus.draft,
              },
            });
          }
          await tx.prefillSnapshot.create({
            data: {
              form: { connect: { id: currentForm.id } },
              id: crypto.randomUUID(),
              lockedFields: jsonValue(lockedFields),
              owner: { connect: { id: identity.id } },
              response: { connect: { id: responseTargetId } },
              values: jsonValue(values),
            },
          });
        }
        const consumedClaim = await tx.pendingClaim.updateMany({
          data: { consumedAt: now },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: pendingClaim.id,
          },
        });
        if (consumedClaim.count !== 1) {
          handoffUnavailable();
        }
        const consumedHandoff = await tx.handoff.updateMany({
          data: {
            consumedAt: now,
            responseId: responseTargetId,
            status: HandoffStatus.consumed,
          },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: handoff.id,
            status: HandoffStatus.reserved,
          },
        });
        if (consumedHandoff.count !== 1) {
          handoffUnavailable();
        }
        await createFormAudit(tx, {
          action: "redeem_handoff",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: currentForm.publicId,
        });
        const created = reuseExistingDraft
          ? currentResponse
          : await tx.response.findUnique({
              include: { prefillSnapshot: true },
              where: { id: responseTargetId },
            });
        if (!created) {
          fail(500, "start_failed", "Unable to redeem prefill handoff");
        }
        return created;
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
    return {
      cleanupObjectKeys: uniqueObjectKeys([
        oldDraftObjectKey,
        unusedDraftObjectKey,
      ]),
      response,
    };
  } catch (error) {
    await deleteObjectUnlessCanonical(draftObjectKey);
    const normalizedError =
      databaseErrorCode(error) === "P2002" || isSerializationConflict(error)
        ? new HttpError(
            409,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          )
        : error;
    try {
      await createFormFailureAudit({
        action: "redeem_handoff",
        actorId: identity.id,
        error: normalizedError,
        targetId: publicIdPattern.test(form.publicId) ? form.publicId : null,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw normalizedError;
  }
}
async function consumeSubmittedPrefillHandoff(
  form: FormWithDocuments,
  identity: Identity,
  claimToken: string,
  responseId: string,
  clock: () => Date = () => new Date()
): Promise<void> {
  const claimDigest = tokenDigest(claimToken);
  try {
    await prisma.$transaction(
      async (tx) => {
        const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
          SELECT "id"
          FROM "forms"
          WHERE "id" = ${form.id}::uuid
          FOR UPDATE
        `
        );
        if (!lockedForm) {
          handoffUnavailable();
        }
        const currentForm = await tx.form.findUnique({
          include: {
            publishedTemplate: {
              include: { prefillConfiguration: { include: { fields: true } } },
            },
          },
          where: { id: lockedForm.id },
        });
        const configuration = publishedPrefillConfiguration(currentForm);
        const currentTemplate = currentForm?.publishedTemplate;
        const [lockedClaim] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
          SELECT "id"
          FROM "pending_claims"
          WHERE "claim_digest" = ${claimDigest}
          FOR UPDATE
        `
        );
        if (!lockedClaim) {
          handoffUnavailable();
        }
        const pendingClaim = await tx.pendingClaim.findUnique({
          include: { handoff: true },
          where: { id: lockedClaim.id },
        });
        const handoff = pendingClaim?.handoff;
        const currentResponse = await tx.response.findUnique({
          select: { id: true, status: true },
          where: { id: responseId },
        });
        const now = clock();
        if (
          !currentForm ||
          currentForm.publicId !== form.publicId ||
          !currentTemplate ||
          !configuration ||
          !currentResponse ||
          currentResponse.status !== ResponseStatus.submitted ||
          currentResponse.id !== responseId ||
          !pendingClaim ||
          !handoff ||
          pendingClaim.consumedAt ||
          pendingClaim.expiresAt <= now ||
          handoff.formId !== currentForm.id ||
          handoff.responseId ||
          handoff.status !== HandoffStatus.reserved ||
          handoff.consumedAt ||
          handoff.expiresAt <= now ||
          handoff.normalizedEmail !== normalizeEmail(identity.email) ||
          handoff.configurationHash !== configuration.configurationHash ||
          configuration.publishedTemplateId !== currentTemplate.id ||
          configuration.configurationHash !== currentTemplate.contentHash
        ) {
          handoffUnavailable();
        }
        const consumedClaim = await tx.pendingClaim.updateMany({
          data: { consumedAt: now },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: pendingClaim.id,
          },
        });
        if (consumedClaim.count !== 1) {
          handoffUnavailable();
        }
        const consumedHandoff = await tx.handoff.updateMany({
          data: {
            consumedAt: now,
            responseId,
            status: HandoffStatus.consumed,
          },
          where: {
            consumedAt: null,
            expiresAt: { gt: now },
            id: handoff.id,
            responseId: null,
            status: HandoffStatus.reserved,
          },
        });
        if (consumedHandoff.count !== 1) {
          handoffUnavailable();
        }
        await createFormAudit(tx, {
          action: "redeem_handoff",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: currentForm.publicId,
        });
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
  } catch (error) {
    const normalizedError =
      databaseErrorCode(error) === "P2002" || isSerializationConflict(error)
        ? new HttpError(
            409,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          )
        : error;
    try {
      await createFormFailureAudit({
        action: "redeem_handoff",
        actorId: identity.id,
        error: normalizedError,
        targetId: publicIdPattern.test(form.publicId) ? form.publicId : null,
      });
    } catch {
      // Preserve the retryable response if the failure audit cannot be written.
    }
    throw normalizedError;
  }
}

type ExternalPrefillStatus =
  | "deleted"
  | "draft"
  | "expired"
  | "pending"
  | "submitted";

function externalPrefillStatus(handoff: {
  consumedAt: Date | null;
  createdAt: Date;
  expiresAt: Date;
  reservedAt: Date | null;
  status: HandoffStatus;
  updatedAt: Date;
  response?: {
    corrections: { revision: number }[];
    status: ResponseStatus;
    submission: { createdAt: Date } | null;
  } | null;
}): {
  consumedAt: string | null;
  createdAt: string;
  expiresAt: string;
  latestCorrectionNumber: number | null;
  reservedAt: string | null;
  status: ExternalPrefillStatus;
  submittedAt: string | null;
  updatedAt: string;
} {
  const status =
    handoff.status === HandoffStatus.deleted
      ? "deleted"
      : handoff.status === HandoffStatus.expired
        ? "expired"
        : handoff.response?.status === ResponseStatus.submitted
          ? "submitted"
          : handoff.response
            ? "draft"
            : "pending";
  return {
    consumedAt: handoff.consumedAt?.toISOString() ?? null,
    createdAt: handoff.createdAt.toISOString(),
    expiresAt: handoff.expiresAt.toISOString(),
    latestCorrectionNumber: handoff.response?.corrections[0]?.revision ?? null,
    reservedAt: handoff.reservedAt?.toISOString() ?? null,
    status,
    submittedAt: handoff.response?.submission?.createdAt.toISOString() ?? null,
    updatedAt: handoff.updatedAt.toISOString(),
  };
}
function deletedExternalPrefillStatus(
  createdAt: Date
): ReturnType<typeof externalPrefillStatus> & { deletedAt: string } {
  const timestamp = createdAt.toISOString();
  return {
    consumedAt: null,
    createdAt: timestamp,
    deletedAt: timestamp,
    expiresAt: timestamp,
    latestCorrectionNumber: null,
    reservedAt: null,
    status: "deleted",
    submittedAt: null,
    updatedAt: timestamp,
  };
}

async function pollPrefillHandoffStatus(
  externalReference: string,
  clock: () => Date = () => new Date()
): Promise<ReturnType<typeof externalPrefillStatus>> {
  const relations = {
    response: {
      select: {
        corrections: {
          orderBy: { revision: "desc" as const },
          select: { revision: true },
          take: 1,
        },
        status: true,
        submission: { select: { createdAt: true } },
      },
    },
  } as const;
  const externalReferenceDigest = tokenDigest(externalReference);
  const handoffResult = await prisma.$transaction(async (tx) => {
    const deletionTombstone = await tx.deletionTombstone.findUnique({
      where: { externalReferenceDigest },
    });
    if (deletionTombstone) {
      const pendingCleanup = await tx.objectCleanupIntent.count({
        where: {
          deletionResponseLookupDigest: deletionTombstone.responseLookupDigest,
        },
      });
      return pendingCleanup > 0
        ? { cleanupPending: true }
        : { deletedAt: deletionTombstone.createdAt };
    }
    const current = await tx.handoff.findFirst({
      include: relations,
      orderBy: { createdAt: "desc" },
      where: { externalReferenceDigest },
    });
    if (!current) {
      handoffUnavailable();
    }
    if (
      current.status === HandoffStatus.deleted &&
      current.deletionResponseLookupDigest
    ) {
      const pendingCleanup = await tx.objectCleanupIntent.count({
        where: {
          deletionResponseLookupDigest: current.deletionResponseLookupDigest,
        },
      });
      if (pendingCleanup > 0) {
        return { cleanupPending: true };
      }
    }
    const now = clock();
    if (
      (current.status === HandoffStatus.pending ||
        current.status === HandoffStatus.reserved) &&
      current.expiresAt <= now
    ) {
      await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`
          SELECT "id"
          FROM "pending_claims"
          WHERE "handoff_id" = ${current.id}::uuid
          FOR UPDATE
        `
      );
      const expired = await tx.handoff.updateMany({
        data: {
          codeDigest: null,
          configurationHash: null,
          filteredValues: Prisma.JsonNull,
          formId: null,
          normalizedEmail: null,
          responseId: null,
          status: HandoffStatus.expired,
        },
        where: {
          expiresAt: { lte: now },
          id: current.id,
          status: { in: [HandoffStatus.pending, HandoffStatus.reserved] },
        },
      });
      if (expired.count === 1) {
        await tx.pendingClaim.deleteMany({ where: { handoffId: current.id } });
      }
      const refreshed = await tx.handoff.findUnique({
        include: relations,
        where: { id: current.id },
      });
      if (!refreshed) {
        handoffUnavailable();
      }
      return { handoff: refreshed };
    }
    return { handoff: current };
  });
  if (handoffResult.cleanupPending) {
    handoffUnavailable();
  }
  if (handoffResult.deletedAt instanceof Date) {
    return deletedExternalPrefillStatus(handoffResult.deletedAt);
  }
  if (!handoffResult.handoff) {
    handoffUnavailable();
  }
  if (handoffResult.handoff.status === HandoffStatus.deleted) {
    return deletedExternalPrefillStatus(handoffResult.handoff.updatedAt);
  }
  return externalPrefillStatus(handoffResult.handoff);
}

async function findOwnedResponse(
  responseId: string,
  formId: string,
  userId: string
): Promise<ResponseWithSnapshot> {
  validateId(responseId, "Response");
  const response = await prisma.response.findFirst({
    include: { prefillSnapshot: true },
    where: { formId, id: responseId, userId },
  });
  if (!response) {
    fail(404, "not_found", "Response was not found");
  }
  return response;
}

function canReadSubmission(identity: Identity, submission: Submission): void {
  if (identity.role === "admin") {
    return;
  }
  if (submission.userId !== identity.id) {
    fail(403, "forbidden", "You may only access your own submission");
  }
}
function lockedPrefillForSnapshot(
  snapshot: PrefillSnapshot | null
): { data: JsonRecord; editableFields: JsonRecord } | undefined {
  if (!snapshot) {
    return;
  }
  const values = jsonRecord(snapshot.values);
  const lockedFields = jsonRecord(snapshot.lockedFields);
  const data = Object.fromEntries(
    Object.entries(values).filter(([tag]) => lockedFields[tag] === true)
  );
  return Object.keys(data).length > 0
    ? { data, editableFields: editableFieldsForSnapshot(snapshot) }
    : undefined;
}

async function correctionEditorConfig(
  responseId: string,
  identity: Identity
): Promise<Record<string, unknown>> {
  validateId(responseId, "Response");
  const response = await prisma.response.findUnique({
    include: {
      corrections: {
        orderBy: { revision: "desc" },
        take: 1,
      },
      prefillSnapshot: true,
      submission: true,
    },
    where: { id: responseId },
  });
  if (
    !response ||
    response.status !== ResponseStatus.submitted ||
    !response.submission
  ) {
    fail(404, "not_found", "Submitted Response was not found");
  }
  const form = await prisma.form.findUnique({
    include: { publishedTemplate: true, templateDraft: true },
    where: { id: response.formId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  const latest = response.corrections[0];
  const baseRevision = latest?.revision ?? 0;
  const baseDocumentKey =
    latest?.documentKey ?? response.submission.documentKey;
  const sourceObjectKey = latest?.objectKey ?? response.submission.objectKey;
  if (!(await objectExists(sourceObjectKey))) {
    fail(409, "document_unavailable", "Response document is unavailable");
  }

  const now = new Date();
  const priorLease = await prisma.editorLease.findFirst({
    select: {
      expiresAt: true,
      id: true,
      workspaceBaseDocumentKey: true,
      workspaceBaseRevision: true,
      workspaceDocumentKey: true,
      workspaceObjectKey: true,
    },
    where: {
      holderSessionId: identity.sessionId,
      holderUserId: identity.id,
      targetId: response.id,
      targetType: OperationTargetType.correction,
    },
  });
  let lease: EditorLeaseGrant;
  if (
    priorLease &&
    priorLease.expiresAt > now &&
    priorLease.workspaceBaseDocumentKey === baseDocumentKey &&
    priorLease.workspaceBaseRevision === baseRevision &&
    priorLease.workspaceDocumentKey &&
    priorLease.workspaceObjectKey &&
    (await objectExists(priorLease.workspaceObjectKey))
  ) {
    const renewed = await renewEditorLease(identity, priorLease.id);
    lease = {
      ...renewed,
      proof: editorLeaseProof(identity, "correction", response.id),
      workspaceBaseDocumentKey: priorLease.workspaceBaseDocumentKey,
      workspaceBaseRevision: priorLease.workspaceBaseRevision,
      workspaceDocumentKey: priorLease.workspaceDocumentKey,
      workspaceObjectKey: priorLease.workspaceObjectKey,
    };
  } else {
    if (priorLease) {
      await releaseEditorLease(identity, priorLease.id);
    }
    const workspace: CorrectionWorkspaceInput = {
      baseDocumentKey,
      baseRevision,
      documentKey: `correction-workspace-${response.id}-${crypto.randomUUID()}`,
      objectKey: objectKey(
        "responses",
        response.id,
        "correction-workspaces",
        crypto.randomUUID(),
        "docx"
      ),
    };
    const workspaceCleanupAfter = new Date(Date.now() + operationTimeoutMs);
    await prisma.objectCleanupIntent.upsert({
      create: {
        cleanupAfter: workspaceCleanupAfter,
        objectKey: workspace.objectKey,
      },
      update: { cleanupAfter: workspaceCleanupAfter },
      where: { objectKey: workspace.objectKey },
    });
    await putObject(
      workspace.objectKey,
      await readObject(sourceObjectKey),
      DOCX_CONTENT_TYPE
    );
    try {
      lease = await claimEditorLease(
        identity,
        "correction",
        response.id,
        form.id,
        workspace
      );
    } catch (error) {
      if (await deleteObjectUnlessCanonical(workspace.objectKey)) {
        await prisma.objectCleanupIntent.deleteMany({
          where: { objectKey: workspace.objectKey },
        });
      }
      throw error;
    }
    if (
      lease.workspaceDocumentKey !== workspace.documentKey ||
      lease.workspaceObjectKey !== workspace.objectKey
    ) {
      if (await deleteObjectUnlessCanonical(workspace.objectKey)) {
        await prisma.objectCleanupIntent.deleteMany({
          where: { objectKey: workspace.objectKey },
        });
      }
      fail(
        409,
        "editor_lease_inactive",
        "The correction workspace is unavailable"
      );
    }
    await prisma.objectCleanupIntent.deleteMany({
      where: { objectKey: workspace.objectKey },
    });
    if (
      priorLease?.workspaceObjectKey &&
      priorLease.workspaceObjectKey !== lease.workspaceObjectKey
    ) {
      await deleteObjectUnlessCanonical(priorLease.workspaceObjectKey);
    }
  }
  const currentRevision = await prisma.response.findUnique({
    select: {
      corrections: {
        orderBy: { revision: "desc" },
        select: {
          documentKey: true,
          id: true,
          objectKey: true,
          revision: true,
        },
        take: 1,
      },
      submission: { select: { documentKey: true, objectKey: true } },
    },
    where: { id: response.id },
  });
  const currentLatest = currentRevision?.corrections[0];
  const currentBaseRevision = currentLatest?.revision ?? 0;
  const currentBaseDocumentKey =
    currentLatest?.documentKey ?? currentRevision?.submission?.documentKey;
  const currentSourceObjectKey =
    currentLatest?.objectKey ?? currentRevision?.submission?.objectKey;
  if (
    currentBaseRevision !== baseRevision ||
    currentBaseDocumentKey !== baseDocumentKey ||
    currentSourceObjectKey !== sourceObjectKey
  ) {
    await releaseEditorLease(identity, lease.id);
    fail(
      409,
      "stale_document",
      "The Response changed while the correction editor was opening"
    );
  }
  await createResponseAudit({
    action: currentLatest ? "view_correction" : "view_response",
    actorId: identity.id,
    outcome: AuditOutcome.success,
    safeMetadata: { revision: currentBaseRevision, state: "submitted" },
    targetId: currentLatest?.id ?? response.id,
    targetType: currentLatest ? "correction" : "response",
  });
  const capabilityScope = {
    documentKey: lease.workspaceDocumentKey as string,
    formId: form.id,
    targetId: response.id,
    targetType: "correction",
  } as const;
  return editorConfig(
    {
      action: "correction",
      capabilities: {
        "save-correction": actionEditorCapability(
          identity,
          capabilityScope,
          "save-correction",
          lease
        ),
      },
      documentKey: capabilityScope.documentKey,
      lease: editorLeaseBridge(lease),
      prefill: lockedPrefillForSnapshot(response.prefillSnapshot),
      publicId: form.publicId,
      responseId: response.id,
      tagAliases: form.publishedTemplate?.objectKey
        ? onlyOfficeFieldTagAliases(
            await readObject(form.publishedTemplate.objectKey)
          )
        : {},
    },
    identity
  );
}

interface CorrectionRevisionSummary {
  actorEmail: string | null;
  actorName: string | null;
  createdAt: Date;
  data: JsonRecord;
  documentAvailable: boolean;
  documentKey: string;
  id: string | null;
  reason: string | null;
  revision: number;
  pictures: Record<string, boolean> | null;
}
type RevisionSelector = number | "latest";

function revisionSelector(value: unknown): RevisionSelector {
  if (value === undefined || value === "latest") {
    return "latest";
  }
  if (value === "original") {
    return 0;
  }
  if (typeof value === "string" && /^(0|[1-9]\d*)$/u.test(value)) {
    const revision = Number(value);
    if (Number.isSafeInteger(revision)) {
      return revision;
    }
  }
  fail(
    400,
    "invalid_request",
    "revision must be latest, original, or a revision number"
  );
}

function selectedSubmissionRevision(
  submission: Submission & { corrections: Correction[] },
  selector: RevisionSelector
): {
  correction: Correction | null;
  data: JsonRecord;
  documentKey: string;
  objectKey: string;
  revision: number;
} {
  const revision =
    selector === "latest"
      ? (submission.corrections[0]?.revision ?? 0)
      : selector;
  if (
    revision > 0 &&
    !submission.corrections.some((item) => item.revision === revision)
  ) {
    fail(404, "not_found", "Submission revision was not found");
  }
  const correction =
    revision === 0
      ? null
      : (submission.corrections.find((item) => item.revision === revision) ??
        null);
  return {
    correction,
    data: jsonRecord(correction?.data ?? submission.data),
    documentKey: correction?.documentKey ?? submission.documentKey,
    objectKey: correction?.objectKey ?? submission.objectKey,
    revision,
  };
}

function correctionRevisionSummary(
  correction: Correction,
  actor: Pick<Actor, "email" | "name"> | undefined,
  documentAvailable: boolean,
  pictures: Record<string, boolean> | null
): CorrectionRevisionSummary {
  return {
    actorEmail: actor?.email ?? null,
    actorName: actor?.name ?? null,
    createdAt: correction.createdAt,
    data: jsonRecord(correction.data),
    documentAvailable,
    documentKey: correction.documentKey,
    id: correction.id,
    pictures,
    reason: correction.reason,
    revision: correction.revision,
  };
}

function findSubmissionWithRevisions(
  id: string,
  includeFieldManifest: true
): Promise<SubmissionWithManifest | null>;
function findSubmissionWithRevisions(
  id: string,
  includeFieldManifest?: false
): Promise<SubmissionWithRevisions | null>;
function findSubmissionWithRevisions(
  id: string,
  includeFieldManifest = false
): Promise<SubmissionWithManifest | SubmissionWithRevisions | null> {
  return prisma.submission.findUnique({
    include: {
      corrections: { orderBy: { revision: "desc" } },
      form: true,
      owner: true,
      ...(includeFieldManifest
        ? {
            response: {
              include: {
                publishedTemplate: {
                  include: {
                    manifest: {
                      include: {
                        fields: { orderBy: { position: "asc" } },
                      },
                    },
                  },
                },
              },
            },
          }
        : {}),
    },
    where: { id },
  });
}
async function receiptManifestFields(submission: SubmissionWithManifest) {
  const publishedTemplate = submission.response?.publishedTemplate;
  const manifest = publishedTemplate?.manifest;
  if (!manifest || !publishedTemplate) {
    return [];
  }
  return manifestFieldsWithDocumentMetadata(
    publishedTemplate.objectKey,
    manifest.displayMetadataVersion,
    manifest.fields
  );
}
function receiptField(
  field: ManifestDisplayField & { options: unknown; type: FieldType }
) {
  return {
    label: field.label,
    options: Array.isArray(field.options)
      ? field.options.flatMap((option) => {
          if (!option || typeof option !== "object" || Array.isArray(option)) {
            return [];
          }
          const optionRecord = option as Record<string, unknown>;
          const displayText = optionRecord.displayText;
          const value = optionRecord.value;
          return typeof displayText === "string" && typeof value === "string"
            ? [{ displayText, value }]
            : [];
        })
      : [],
    placeholder: field.placeholder,
    position: field.position,
    tag: field.tag,
    type: field.type,
  };
}

async function userEditorConfig(
  form: FormWithDocuments,
  identity: Identity,
  responseId: string | undefined,
  requestedAction?: string
): Promise<Record<string, unknown>> {
  const publishedTemplate = form.publishedTemplate;
  if (!publishedTemplate?.objectKey || !publishedTemplate.documentKey) {
    fail(409, "not_published", "This form has no published document");
  }
  const response = responseId
    ? await prisma.response.findFirst({
        include: { prefillSnapshot: true },
        where: {
          formId: form.id,
          id: responseId,
          userId: identity.id,
        },
      })
    : await prisma.response.findFirst({
        include: { prefillSnapshot: true },
        where: { formId: form.id, userId: identity.id },
      });
  if (!response) {
    fail(404, "not_found", "Start a response before opening the editor");
  }
  if (response.status === ResponseStatus.submitted) {
    fail(409, "already_submitted", "This response has already been submitted");
  }
  if (!response.draftDocumentKey || !response.draftObjectKey) {
    fail(409, "document_unavailable", "Response document is unavailable");
  }
  if (!(await objectExists(response.draftObjectKey))) {
    fail(
      409,
      "document_unavailable",
      "Response document artifact is unavailable"
    );
  }
  const snapshot = response.prefillSnapshot;
  const capabilityScope = {
    documentKey: response.draftDocumentKey,
    formId: form.id,
    targetId: response.id,
    targetType: "response",
  } as const;
  const lease = await claimEditorLease(
    identity,
    capabilityScope.targetType,
    capabilityScope.targetId,
    form.id
  );
  if (form.fillMethod === FillMethod.native) {
    const nativeFieldConfig = await nativeFields(
      publishedTemplate.id,
      await readObject(response.draftObjectKey)
    );
    return {
      capabilities: {
        "save-draft": actionEditorCapability(
          identity,
          capabilityScope,
          "save-draft",
          lease
        ),
        submit: actionEditorCapability(
          identity,
          capabilityScope,
          "submit",
          lease
        ),
      },
      data:
        response.draftData === null
          ? jsonRecord(snapshot?.values ?? {})
          : jsonRecord(response.draftData),
      documentKey: response.draftDocumentKey,
      fields: nativeFieldConfig.fields,
      fillMethod: FillMethod.native,
      lockedFields: snapshot ? jsonRecord(snapshot.lockedFields) : {},
      pictures: nativeFieldConfig.pictures,
      responseId: response.id,
    };
  }
  return {
    ...editorConfig(
      {
        action:
          requestedAction === "submit"
            ? "submit"
            : requestedAction === "fill"
              ? "fill"
              : "draft",
        capabilities: {
          "save-draft": actionEditorCapability(
            identity,
            capabilityScope,
            "save-draft",
            lease
          ),
          submit: actionEditorCapability(
            identity,
            capabilityScope,
            "submit",
            lease
          ),
        },
        documentKey: response.draftDocumentKey,
        lease: editorLeaseBridge(lease),
        prefill: snapshot
          ? requestedAction === "fill"
            ? {
                data: jsonRecord(snapshot.values),
                editableFields: editableFieldsForSnapshot(snapshot),
              }
            : lockedPrefillForSnapshot(snapshot)
          : undefined,
        publicId: form.publicId,
        responseId: response.id,
        tagAliases: onlyOfficeFieldTagAliases(
          await readObject(publishedTemplate.objectKey)
        ),
      },
      identity
    ),
    fillMethod: FillMethod.onlyoffice,
  };
}

export function createApp(options: AppOptions = {}) {
  const onlyOffice = options.onlyOffice ?? createOnlyOfficeClient();
  const removeObject = options.deleteObject ?? deleteObject;
  const storeObject = options.putObject ?? putObject;
  const allowedCallbackOrigins = options.onlyOfficeCallbackOrigins
    ? new Set(options.onlyOfficeCallbackOrigins)
    : callbackOrigins;
  const callbackMaximumBytes =
    options.onlyOfficeCallbackMaxBytes ?? maxCallbackDocumentBytes;
  const prefillHandoffSecret =
    options.prefillHandoffSecret ?? env.PREFILL_HANDOFF_SECRET;
  const prefillReturnUrl = configuredPrefillReturnUrl(
    options.prefillReturnUrl ?? env.PREFILL_RETURN_URL
  );
  const handoffClock = options.clock ?? (() => new Date());
  const legacySso =
    options.legacySso === undefined
      ? legacySsoConfigurationFromEnv()
      : options.legacySso
        ? validatedLegacySsoConfiguration(options.legacySso)
        : null;
  let omniRoute = options.omniRoute;
  if (omniRoute === undefined) {
    const apiKey = env.OMNIROUTE_API_KEY;
    const baseUrl = env.OMNIROUTE_BASE_URL;
    const model = env.OMNIROUTE_MODEL;
    omniRoute = apiKey && baseUrl && model ? { apiKey, baseUrl, model } : null;
  }
  const aiAuthoring = new AiAuthoringSessions(omniRoute, options.clock);
  return new Elysia()
    .onError(({ error, set }) => {
      if (error instanceof AiAuthoringError) {
        set.status = error.status;
        return Response.json({
          error: error.code,
          message: error.message,
        });
      }
      if (error instanceof HttpError) {
        set.status = error.httpStatus;
        return Response.json({
          error: error.code,
          message: error.message,
        });
      }
      if (databaseErrorCode(error) === "P2002") {
        set.status = 409;
        return Response.json({
          error: "operation_conflict",
          message: "Another operation is already in progress",
        });
      }
      console.error(error);
      set.status = 500;
      return Response.json({
        error: "internal_error",
        message: "An unexpected server error occurred",
      });
    })
    .use(
      cors({
        allowedHeaders: [
          "Authorization",
          "Content-Type",
          "X-Editor-Capability",
          "X-Prefill-Handoff-Secret",
        ],
        credentials: true,
        methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
        origin: env.CORS_ORIGIN,
      })
    )
    .post("/api/integrations/prefill/handoffs", async ({ body, request }) => {
      if (
        !prefillHandoffSecretMatches(
          prefillHandoffSecret,
          request.headers.get("x-prefill-handoff-secret")
        )
      ) {
        try {
          await createFormFailureAudit({
            action: "create_handoff",
            actorId: null,
            error: new HttpError(
              404,
              "handoff_unavailable",
              "The prefill handoff is unavailable"
            ),
            targetId: null,
          });
        } catch {
          // Preserve the non-enumerating response if the audit cannot be written.
        }
        fail(404, "handoff_unavailable", "The prefill handoff is unavailable");
      }
      try {
        return await createPrefillHandoff(handoffCreateInput(asRecord(body)));
      } catch (error) {
        if (error instanceof HttpError && error.code === "not_found") {
          fail(
            404,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          );
        }
        throw error;
      }
    })
    .post("/api/integrations/prefill/status", async ({ body, request }) => {
      if (
        !prefillHandoffSecretMatches(
          prefillHandoffSecret,
          request.headers.get("x-prefill-handoff-secret")
        )
      ) {
        fail(404, "handoff_unavailable", "The prefill handoff is unavailable");
      }
      try {
        return await pollPrefillHandoffStatus(
          handoffStatusInput(asRecord(body)).externalReference,
          handoffClock
        );
      } catch (error) {
        if (
          error instanceof HttpError &&
          error.code === "handoff_unavailable"
        ) {
          fail(
            404,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          );
        }
        throw error;
      }
    })
    .post(
      "/prefill/handoff",
      async ({ request }) => {
        try {
          requireTopLevelNavigation(request);
          const launch = await launchPrefillHandoff(
            await readPrefillHandoffCode(request),
            handoffClock
          );
          return new Response(null, {
            headers: {
              Location: `/forms/${launch.publicId}/fill`,
              "Set-Cookie": pendingClaimCookie(
                launch.claimToken,
                pendingClaimLifetimeSeconds
              ),
            },
            status: 303,
          });
        } catch (error) {
          try {
            await createFormFailureAudit({
              action: "launch_handoff",
              actorId: null,
              error,
              targetId: null,
            });
          } catch {
            // Preserve the retryable redirect if the audit cannot be written.
          }
          return new Response(null, {
            headers: {
              Location: "/handoff?error=handoff_unavailable",
              "Set-Cookie": pendingClaimCookie("", 0),
            },
            status: 303,
          });
        }
      },
      { parse: "none" }
    )
    .get("/api/legacy-sso/status", () =>
      Response.json(
        { enabled: Boolean(legacySso) },
        { headers: { "Cache-Control": "no-store" } }
      )
    )
    .post(
      "/api/legacy-sso/start",
      async ({ request }) => {
        if (!legacySso) {
          fail(404, "legacy_sso_unavailable", "Legacy SSO is unavailable");
        }
        const identity = await identityFor(request);
        if (request.headers.has("authorization") && !identity) {
          fail(401, "unauthorized", "Authentication is required");
        }
        return startLegacySso(request, legacySso, handoffClock, identity);
      },
      { parse: "none" }
    )
    .get(
      "/api/legacy-sso/callback",
      ({ request }) =>
        legacySso
          ? completeLegacySso(request, legacySso, handoffClock)
          : legacySsoFailureResponse(),
      { parse: "none" }
    )
    .get(
      "/api/legacy-sso/switch",
      ({ request }) => readLegacySsoSwitch(request, handoffClock, legacySso),
      { parse: "none" }
    )
    .post(
      "/api/legacy-sso/switch/confirm",
      ({ request }) =>
        consumeLegacySsoSwitch(request, handoffClock, legacySso, true),
      { parse: "none" }
    )
    .post(
      "/api/legacy-sso/switch/cancel",
      ({ request }) =>
        consumeLegacySsoSwitch(request, handoffClock, legacySso, false),
      { parse: "none" }
    )
    .post(
      "/api/legacy-sso/session",
      ({ request }) => claimLegacySsoSession(request, handoffClock, legacySso),
      { parse: "none" }
    )
    .get("/prefill/handoff", () => {
      fail(
        405,
        "handoff_unavailable",
        "The prefill handoff must be launched with a top-level POST"
      );
    })
    .post("/api/auth/sign-in/email", ({ body, request, server }) => {
      const sourceIp =
        options.requestIp?.(request) ??
        server?.requestIP(request)?.address ??
        "unknown";
      return handleEmailSignIn(request, body, sourceIp, aiAuthoring);
    })
    .get("/api/session", async ({ request }) => {
      const identity = await identityFor(request);
      if (!identity) {
        fail(401, "unauthorized", "Authentication is required");
      }
      return {
        session: { expiresAt: identity.expiresAt },
        user: {
          email: identity.email,
          id: identity.id,
          mustChangePassword: identity.mustChangePassword && !identity.isSso,
          name: identity.name,
          role: identity.role,
        },
      };
    })
    .post("/api/editor-leases/:id/renew", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Editor lease");
      const lease = await renewEditorLease(identity, params.id);
      return {
        lease: { expiresAt: lease.expiresAt.toISOString(), id: lease.id },
      };
    })
    .delete("/api/editor-leases/:id", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Editor lease");
      await releaseEditorLease(identity, params.id);
      return { ok: true };
    })
    .post("/api/account/password", async ({ body, request }) => {
      const identity = await identityFor(request);
      if (!identity) {
        fail(401, "unauthorized", "Authentication is required");
      }
      const input = asRecord(body);
      const { currentPassword, newPassword } = input;
      if (
        typeof currentPassword !== "string" ||
        typeof newPassword !== "string"
      ) {
        fail(
          400,
          "invalid_request",
          "currentPassword and newPassword are required"
        );
      }
      if (newPassword.length < passwordMinimumLength) {
        fail(
          400,
          "password_too_short",
          `Password must contain at least ${passwordMinimumLength} characters`
        );
      }
      if (newPassword.length > passwordMaximumLength) {
        fail(
          400,
          "password_too_long",
          `Password must contain at most ${passwordMaximumLength} characters`
        );
      }
      if (
        currentPassword.length < passwordMinimumLength ||
        currentPassword.length > passwordMaximumLength
      ) {
        fail(400, "invalid_current_password", "Current password is invalid");
      }
      const authContext = await auth.$context;

      const account = await prisma.account.findFirst({
        where: { providerId: "credential", userId: identity.id },
      });
      const currentHash = account?.password;
      if (
        !account ||
        !currentHash ||
        !(await authContext.password.verify({
          hash: currentHash,
          password: currentPassword,
        }))
      ) {
        fail(400, "invalid_current_password", "Current password is invalid");
      }

      const passwordHash = await authContext.password.hash(newPassword);
      const revokedOwnerSessionIds = await prisma.$transaction(async (tx) => {
        const update = await tx.account.updateMany({
          data: { password: passwordHash },
          where: { id: account.id, password: currentHash },
        });
        if (update.count !== 1) {
          fail(409, "credential_changed", "Credential changed concurrently");
        }
        await tx.user.update({
          data: { mustChangePassword: false },
          where: { id: identity.id },
        });
        const sessions = await tx.session.findMany({
          select: { id: true },
          where: { userId: identity.id },
        });
        await tx.session.deleteMany({ where: { userId: identity.id } });
        return sessions.map(({ id }) => id);
      });
      await endAiAuthoringSessions(aiAuthoring, revokedOwnerSessionIds);
      return { ok: true };
    })
    .get("/api/admin/account-links", async ({ request, query }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const queryRecord = query as unknown as JsonRecord;
      const cursor = queryString(queryRecord, "cursor");
      if (cursor !== undefined && !idPattern.test(cursor)) {
        fail(400, "invalid_request", "cursor is invalid");
      }
      const where: Prisma.LegacyAccountLinkRequestWhereInput = {
        status: LegacyAccountLinkStatus.pending,
      };
      if (cursor) {
        where.id = { gt: cursor };
      }
      const rows = await prisma.legacyAccountLinkRequest.findMany({
        orderBy: { id: "asc" },
        select: {
          createdAt: true,
          email: true,
          id: true,
          providerId: true,
          status: true,
          subject: true,
          user: { select: adminUserSelect },
        },
        take: accountUserPageSize + 1,
        where,
      });
      const page = rows.slice(0, accountUserPageSize);
      return {
        nextCursor:
          rows.length > accountUserPageSize ? (page.at(-1)?.id ?? null) : null,
        requests: page.map(({ user, ...linkRequest }) => ({
          ...linkRequest,
          user: accountUserSummary(user),
        })),
      };
    })
    .post(
      "/api/admin/account-links/:id/approve",
      ({ request, params }) => {
        const requestId = accountAuditTargetId(params.id);
        if (!requestId) {
          fail(400, "invalid_request", "Account link request id is invalid");
        }
        return withAdminMutation(
          request,
          "approve_legacy_account_link",
          async () => {
            const linkRequest =
              await prisma.legacyAccountLinkRequest.findUnique({
                select: { userId: true },
                where: { id: requestId },
              });
            if (!linkRequest) {
              fail(404, "not_found", "Account link request was not found");
            }
            return linkRequest.userId;
          },
          (identity) => reviewLegacyAccountLink(identity, requestId, "approved")
        );
      },
      { parse: "none" }
    )
    .post(
      "/api/admin/account-links/:id/reject",
      ({ request, params }) => {
        const requestId = accountAuditTargetId(params.id);
        if (!requestId) {
          fail(400, "invalid_request", "Account link request id is invalid");
        }
        return withAdminMutation(
          request,
          "reject_legacy_account_link",
          async () => {
            const linkRequest =
              await prisma.legacyAccountLinkRequest.findUnique({
                select: { userId: true },
                where: { id: requestId },
              });
            if (!linkRequest) {
              fail(404, "not_found", "Account link request was not found");
            }
            return linkRequest.userId;
          },
          (identity) => reviewLegacyAccountLink(identity, requestId, "rejected")
        );
      },
      { parse: "none" }
    )
    .get("/api/admin/users", async ({ request, query }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const queryRecord = query as unknown as JsonRecord;
      const cursor = queryString(queryRecord, "cursor");
      const emailQuery = queryString(queryRecord, "email");
      const roleQuery = queryString(queryRecord, "role");
      const enabledQuery = queryString(queryRecord, "enabled");
      if (cursor !== undefined && !idPattern.test(cursor)) {
        fail(400, "invalid_request", "cursor is invalid");
      }
      const email = emailQuery ? normalizeEmail(emailQuery) : undefined;
      if (email && email.length > accountEmailMaximumLength) {
        fail(400, "invalid_request", "email filter is too long");
      }
      let role: UserRole | undefined;
      if (roleQuery !== undefined) {
        if (roleQuery !== "admin" && roleQuery !== "user") {
          fail(400, "invalid_request", "role must be admin or user");
        }
        role = roleQuery;
      }
      let enabled: boolean | undefined;
      if (enabledQuery !== undefined) {
        if (enabledQuery !== "true" && enabledQuery !== "false") {
          fail(400, "invalid_request", "enabled must be true or false");
        }
        enabled = enabledQuery === "true";
      }
      const where: Prisma.UserWhereInput = {};
      if (cursor) {
        where.id = { gt: cursor };
      }
      if (email) {
        where.email = { contains: email };
      }
      if (role) {
        where.role = role;
      }
      if (enabled !== undefined) {
        where.enabled = enabled;
      }
      const users = await prisma.user.findMany({
        orderBy: { id: "asc" },
        select: adminUserSelect,
        take: accountUserPageSize + 1,
        where,
      });
      const page = users.slice(0, accountUserPageSize);
      return {
        nextCursor:
          users.length > accountUserPageSize ? (page.at(-1)?.id ?? null) : null,
        users: page.map(accountUserSummary),
      };
    })
    .post(
      "/api/admin/users",
      ({ request }) =>
        withAdminMutation(request, "create_user", null, async (identity) => {
          const input = await readJsonRecord(request, accountBodyMaximumBytes);
          const keys = Object.keys(input);
          if (
            keys.length !== 3 ||
            keys.some(
              (key) => key !== "name" && key !== "email" && key !== "role"
            )
          ) {
            fail(400, "invalid_request", "name, email, and role are required");
          }
          const name = requiredString(input, "name");
          if (name.length > accountNameMaximumLength) {
            fail(400, "invalid_request", "name is too long");
          }
          const email = normalizedAccountEmail(requiredString(input, "email"));
          const roleValue = requiredString(input, "role");
          if (roleValue !== "admin" && roleValue !== "user") {
            fail(400, "invalid_request", "role must be admin or user");
          }
          const temporaryPassword = generateTemporaryPassword();
          const user = await accountTransaction(identity, async (tx) => {
            const existing = await tx.user.findUnique({
              select: { id: true },
              where: { email },
            });
            if (existing) {
              fail(409, "email_in_use", "Email is already in use");
            }
            const authContext = await auth.$context;
            const passwordHash =
              await authContext.password.hash(temporaryPassword);
            const userId = crypto.randomUUID();
            const created = await tx.user.create({
              data: {
                accounts: {
                  create: {
                    accountId: userId,
                    id: crypto.randomUUID(),
                    issuer: "local:credential",
                    password: passwordHash,
                    providerId: "credential",
                  },
                },
                email,
                emailVerified: true,
                enabled: true,
                id: userId,
                mustChangePassword: true,
                name,
                role: roleValue,
              },
              select: adminUserSelect,
            });
            await createAccountAudit(tx, {
              action: "create_user",
              actorId: identity.id,
              outcome: AuditOutcome.success,
              safeMetadata: { change: "created" },
              targetId: created.id,
            });
            return created;
          });
          return {
            temporaryPassword,
            user: accountUserSummary(user),
          };
        }),
      { parse: "none" }
    )
    .patch(
      "/api/admin/users/:id",
      ({ request, params }) =>
        withAdminMutation(
          request,
          "update_user",
          accountAuditTargetId(params.id),
          async (identity, setAction) => {
            const userId = validateId(params.id, "User");
            const input = await readJsonRecord(
              request,
              accountBodyMaximumBytes
            );
            const keys = Object.keys(input);
            const key = keys[0];
            if (
              keys.length !== 1 ||
              (key !== "enabled" && key !== "email" && key !== "role")
            ) {
              fail(
                400,
                "invalid_request",
                "Exactly one of enabled, email, or role is required"
              );
            }
            let action: AccountAuditAction = "update_user";
            let change = "updated";
            let updateData: Prisma.UserUpdateInput;
            let newEmail: string | undefined;
            let revokeSessions = false;
            if (key === "enabled") {
              if (typeof input.enabled !== "boolean") {
                fail(400, "invalid_request", "enabled must be a boolean");
              }
              action = input.enabled ? "enable_user" : "disable_user";
              change = input.enabled ? "enabled" : "disabled";
              revokeSessions = !input.enabled;
              updateData = { enabled: input.enabled };
            } else if (key === "email") {
              if (typeof input.email !== "string") {
                fail(400, "invalid_request", "email is required");
              }
              newEmail = normalizedAccountEmail(input.email);
              action = "change_user_email";
              change = "email_changed";
              revokeSessions = true;
              updateData = { email: newEmail };
            } else {
              if (input.role !== "admin" && input.role !== "user") {
                fail(400, "invalid_request", "role must be admin or user");
              }
              action = input.role === "admin" ? "promote_user" : "demote_user";
              change = input.role === "admin" ? "promoted" : "demoted";
              revokeSessions = true;
              updateData = { role: input.role };
            }
            setAction(action);
            const result = await accountTransaction(identity, async (tx) => {
              const target = await lockAccountUser(tx, userId);
              const removesFinalAdmin =
                target.role === "admin" &&
                target.enabled &&
                (updateData.enabled === false || updateData.role === "user");
              if (removesFinalAdmin) {
                const enabledAdminCount = await tx.user.count({
                  where: { enabled: true, role: "admin" },
                });
                if (enabledAdminCount <= 1) {
                  fail(
                    409,
                    "final_admin_required",
                    "At least one enabled Admin is required"
                  );
                }
              }
              if (newEmail !== undefined) {
                const existing = await tx.user.findUnique({
                  select: { id: true },
                  where: { email: newEmail },
                });
                if (existing && existing.id !== target.id) {
                  fail(409, "email_in_use", "Email is already in use");
                }
              }
              if (newEmail !== undefined && newEmail !== target.email) {
                const staleHandoffs = await tx.handoff.findMany({
                  select: { id: true },
                  where: {
                    normalizedEmail: target.email,
                    responseId: null,
                  },
                });
                if (staleHandoffs.length > 0) {
                  const staleHandoffIds = staleHandoffs.map(
                    (handoff) => handoff.id
                  );
                  await tx.pendingClaim.deleteMany({
                    where: { handoffId: { in: staleHandoffIds } },
                  });
                  await tx.handoff.deleteMany({
                    where: { id: { in: staleHandoffIds } },
                  });
                }
              }
              const updated = await tx.user.update({
                data: updateData,
                select: adminUserSelect,
                where: { id: target.id },
              });
              const ownerSessionIds = revokeSessions
                ? (
                    await tx.session.findMany({
                      select: { id: true },
                      where: { userId: target.id },
                    })
                  ).map(({ id }) => id)
                : [];
              if (revokeSessions) {
                await tx.session.deleteMany({ where: { userId: target.id } });
              }
              await createAccountAudit(tx, {
                action,
                actorId: identity.id,
                outcome: AuditOutcome.success,
                safeMetadata: { change },
                targetId: target.id,
              });
              return { ownerSessionIds, user: updated };
            });
            await endAiAuthoringSessions(aiAuthoring, result.ownerSessionIds);
            return { user: accountUserSummary(result.user) };
          }
        ),
      { parse: "none" }
    )
    .delete(
      "/api/admin/users/:id",
      ({ request, params }) =>
        withAdminMutation(
          request,
          "delete_user",
          accountAuditTargetId(params.id),
          async (identity) => {
            const userId = validateId(params.id, "User");
            const input = await readJsonRecord(
              request,
              accountBodyMaximumBytes
            );
            if (Object.keys(input).length !== 1 || input.confirm !== true) {
              fail(400, "invalid_request", "confirm must be true");
            }
            const revokedOwnerSessionIds = await accountTransaction(
              identity,
              async (tx) => {
                const target = await lockAccountUser(tx, userId);
                if (target.role === "admin" && target.enabled) {
                  const enabledAdminCount = await tx.user.count({
                    where: { enabled: true, role: "admin" },
                  });
                  if (enabledAdminCount <= 1) {
                    fail(
                      409,
                      "final_admin_required",
                      "At least one enabled Admin is required"
                    );
                  }
                }
                const responseCount = await tx.response.count({
                  where: { userId: target.id },
                });
                if (responseCount > 0) {
                  fail(
                    409,
                    "personal_data_remains",
                    "Personal Responses must be deleted first"
                  );
                }
                const cleanupIntentCount = await tx.objectCleanupIntent.count({
                  where: { deletionOwnerUserId: target.id },
                });
                if (cleanupIntentCount > 0) {
                  fail(
                    409,
                    "personal_data_remains",
                    "Response object cleanup is still pending"
                  );
                }
                const handoffs = await tx.handoff.findMany({
                  select: { id: true },
                  where: { normalizedEmail: target.email },
                });
                if (handoffs.length > 0) {
                  await tx.pendingClaim.deleteMany({
                    where: {
                      handoffId: {
                        in: handoffs.map((handoff) => handoff.id),
                      },
                    },
                  });
                  await tx.handoff.updateMany({
                    data: {
                      codeDigest: null,
                      configurationHash: null,
                      consumedAt: null,
                      filteredValues: Prisma.JsonNull,
                      formId: null,
                      normalizedEmail: null,
                      reservedAt: null,
                      responseId: null,
                      status: HandoffStatus.deleted,
                    },
                    where: {
                      id: { in: handoffs.map((handoff) => handoff.id) },
                    },
                  });
                }
                await createAccountAudit(tx, {
                  action: "delete_user",
                  actorId: identity.id,
                  outcome: AuditOutcome.success,
                  safeMetadata: { change: "deleted" },
                  targetId: target.id,
                });
                const sessions = await tx.session.findMany({
                  select: { id: true },
                  where: { userId: target.id },
                });
                await tx.user.delete({ where: { id: target.id } });
                return sessions.map(({ id }) => id);
              }
            );
            await endAiAuthoringSessions(aiAuthoring, revokedOwnerSessionIds);
            return { deleted: true };
          }
        ),
      { parse: "none" }
    )
    .post(
      "/api/admin/users/:id/password-reset",
      ({ request, params }) =>
        withAdminMutation(
          request,
          "reset_user_password",
          accountAuditTargetId(params.id),
          async (identity) => {
            const userId = validateId(params.id, "User");
            const temporaryPassword = generateTemporaryPassword();
            const result = await accountTransaction(identity, async (tx) => {
              const target = await lockAccountUser(tx, userId);
              const account = await tx.account.findFirst({
                where: { providerId: "credential", userId: target.id },
              });
              if (!account?.password) {
                fail(500, "credential_missing", "Credential is unavailable");
              }
              const authContext = await auth.$context;
              const passwordHash =
                await authContext.password.hash(temporaryPassword);
              await tx.account.update({
                data: { password: passwordHash },
                where: { id: account.id },
              });
              const updated = await tx.user.update({
                data: { mustChangePassword: true },
                select: adminUserSelect,
                where: { id: target.id },
              });
              const ownerSessionIds = (
                await tx.session.findMany({
                  select: { id: true },
                  where: { userId: target.id },
                })
              ).map(({ id }) => id);
              await tx.session.deleteMany({ where: { userId: target.id } });
              await createAccountAudit(tx, {
                action: "reset_user_password",
                actorId: identity.id,
                outcome: AuditOutcome.success,
                safeMetadata: { change: "password_reset" },
                targetId: target.id,
              });
              return { ownerSessionIds, user: updated };
            });
            await endAiAuthoringSessions(aiAuthoring, result.ownerSessionIds);
            return {
              temporaryPassword,
              user: accountUserSummary(result.user),
            };
          }
        ),
      { parse: "none" }
    )

    .post("/api/auth/sign-out", async ({ request }) => {
      const identity = await identityFor(request);
      if (identity) {
        let sessionRevoked = false;
        try {
          await aiAuthoring.endForSession(identity.sessionId, async () => {
            await prisma.session.deleteMany({
              where: { id: identity.sessionId },
            });
            sessionRevoked = true;
          });
        } catch (error) {
          if (!sessionRevoked) {
            throw error;
          }
          console.error(error);
          return Response.json(
            {
              error: "sign_out_cleanup_failed",
              message: "Session was revoked, but AI authoring cleanup failed",
              sessionRevoked: true,
            },
            { status: 500 }
          );
        }
      }
      return { ok: true };
    })
    .all("/api/auth/*", () =>
      Response.json(
        { error: "not_found", message: "Authentication route was not found" },
        { status: 404 }
      )
    )
    .get("/api/admin/ai-authoring", async ({ request }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      return { disclosure: authoringDisclosure, enabled: aiAuthoring.enabled };
    })
    .post(
      "/api/admin/ai-authoring/sessions",
      async ({ request }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const reservation = aiAuthoring.reserveRequest(identity.sessionId);
        let sourcePdf: Uint8Array | undefined;
        try {
          const input = await readAuthoringCreationInput(request);
          sourcePdf = input.pdfBytes;
          const session = await aiAuthoring.create(
            { authSessionId: identity.sessionId, userId: identity.id },
            input.prompt,
            input.consent,
            (document, expected) => validateTemplateControls(document, expected),
            reservation,
            async () => {
              const currentIdentity = await identityFor(request);
              return (
                currentIdentity !== null &&
                currentIdentity.id === identity.id &&
                currentIdentity.sessionId === identity.sessionId &&
                currentIdentity.role === "admin" &&
                (!currentIdentity.mustChangePassword || currentIdentity.isSso)
              );
            },
            input.pdfBytes
          );
          return Response.json(
            { session },
            { headers: { "Cache-Control": "private, no-store" } }
          );
        } finally {
          sourcePdf?.fill(0);
          aiAuthoring.releaseRequest(reservation);
        }
      },
      { parse: "none" }
    )
    .get(
      "/api/admin/ai-authoring/sessions/current",
      async ({ request }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const session = await aiAuthoring.current({
          authSessionId: identity.sessionId,
          userId: identity.id,
        });
        return Response.json(
          { session },
          { headers: { "Cache-Control": "private, no-store" } }
        );
      }
    )
    .post(
      "/api/admin/ai-authoring/sessions/:sessionId/revisions",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const reservation = aiAuthoring.reserveRequest(identity.sessionId);
        try {
          const input = await readJsonRecord(request, accountBodyMaximumBytes);
          if (
            Object.keys(input).length !== 2 ||
            !Object.hasOwn(input, "consent") ||
            !Object.hasOwn(input, "prompt") ||
            typeof input.prompt !== "string"
          ) {
            fail(
              400,
              "invalid_request",
              "Only consent and prompt are accepted"
            );
          }
          const session = await aiAuthoring.revise(
            params.sessionId,
            { authSessionId: identity.sessionId, userId: identity.id },
            input.prompt,
            input.consent,
            (document, expected) => validateTemplateControls(document, expected),
            reservation,
            async () => {
              const currentIdentity = await identityFor(request);
              return (
                currentIdentity !== null &&
                currentIdentity.id === identity.id &&
                currentIdentity.sessionId === identity.sessionId &&
                currentIdentity.role === "admin" &&
                (!currentIdentity.mustChangePassword || currentIdentity.isSso)
              );
            }
          );
          return Response.json(
            { session },
            { headers: { "Cache-Control": "private, no-store" } }
          );
        } finally {
          aiAuthoring.releaseRequest(reservation);
        }
      },
      { parse: "none" }
    )
    .get(
      "/api/admin/ai-authoring/sessions/:sessionId/docx",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const { bytes } = await aiAuthoring.download(params.sessionId, {
          authSessionId: identity.sessionId,
          userId: identity.id,
        });
        return new Response(bytes, {
          headers: {
            "Cache-Control": "private, no-store",
            "Content-Disposition":
              'attachment; filename="ai-authored-template.docx"',
            "Content-Type": DOCX_CONTENT_TYPE,
            "X-Content-Type-Options": "nosniff",
          },
        });
      }
    )
    .delete(
      "/api/admin/ai-authoring/sessions/:sessionId",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        await aiAuthoring.end(params.sessionId, {
          authSessionId: identity.sessionId,
          userId: identity.id,
        });
        return { ok: true };
      }
    )
    .get("/health", () => ({ ok: true }))
    .get("/ready", async ({ set }) => {
      if (!(await readinessStatus())) {
        set.status = 503;
        return { ok: false };
      }
      return { ok: true };
    })
    .get("/api/admin/forms", async ({ request }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const items = await prisma.form.findMany({
        include: {
          _count: {
            select: {
              responses: { where: { status: ResponseStatus.draft } },
              submissions: true,
            },
          },
          publishedTemplate: true,
          templateDraft: true,
        },
        orderBy: { updatedAt: "desc" },
      });
      const forms = items.map((item) =>
        formDto(item, {
          activeDraftCount: item._count.responses,
          submissionCount: item._count.submissions,
        })
      );
      return { forms };
    })
    .patch("/api/admin/forms/:publicId", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const auditTarget = publicIdPattern.test(params.publicId)
        ? params.publicId
        : null;
      let auditAction: FormAuditAction = "update_form_metadata";
      try {
        const input = await readFormMetadataInput(request);
        auditAction =
          input.status === FormStatus.archived
            ? "archive_form"
            : input.status === FormStatus.published
              ? "unarchive_form"
              : "update_form_metadata";
        const updated = await prisma.$transaction(async (tx) => {
          if (!publicIdPattern.test(params.publicId)) {
            fail(404, "not_found", "Form was not found");
          }
          const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
            Prisma.sql`
              SELECT "id"
              FROM "forms"
              WHERE "public_id" = ${params.publicId}
              FOR UPDATE
            `
          );
          if (!lockedForm) {
            fail(404, "not_found", "Form was not found");
          }
          const current = await tx.form.findUnique({
            include: { publishedTemplate: true, templateDraft: true },
            where: { id: lockedForm.id },
          });
          if (!current) {
            fail(404, "not_found", "Form was not found");
          }
          if (
            input.status !== undefined &&
            (!current.publishedTemplate ||
              (current.status !== FormStatus.published &&
                current.status !== FormStatus.archived))
          ) {
            fail(
              409,
              "form_not_published",
              "Only a published Form can change archive state"
            );
          }
          if (input.fillMethod !== undefined) {
            if (input.fillMethod !== current.fillMethod) {
              const activeOperation = await tx.operation.findFirst({
                select: { id: true },
                where: {
                  formId: current.id,
                  status: {
                    in: [OperationStatus.pending, OperationStatus.processing],
                  },
                },
              });
              if (activeOperation) {
                fail(
                  409,
                  "operation_in_progress",
                  "A response operation is in progress"
                );
              }
            }
            if (input.fillMethod === FillMethod.native) {
              if (
                !current.publishedTemplate ||
                (current.status !== FormStatus.published &&
                  current.status !== FormStatus.archived)
              ) {
                fail(
                  409,
                  "native_fill_unsupported",
                  "Native filling requires supported fields"
                );
              }
              const manifest = await tx.fieldManifest.findUnique({
                include: { fields: { select: { tag: true, type: true } } },
                where: {
                  publishedTemplateId: current.publishedTemplate.id,
                },
              });
              if (
                !manifest ||
                !(await supportsNativeTemplate(
                  current.publishedTemplate.objectKey,
                  manifest.fields
                ))
              ) {
                fail(
                  409,
                  "native_fill_unsupported",
                  "Native filling requires supported fields"
                );
              }
            }
          }
          const form = await tx.form.update({
            data: {
              ...(input.title === undefined ? {} : { title: input.title }),
              ...(Object.hasOwn(input, "description")
                ? { description: input.description }
                : {}),
              ...(input.status === undefined ? {} : { status: input.status }),
              ...(input.fillMethod === undefined
                ? {}
                : { fillMethod: input.fillMethod }),
            },
            include: { publishedTemplate: true, templateDraft: true },
            where: { id: current.id },
          });
          await createFormAudit(tx, {
            action: auditAction,
            actorId: identity.id,
            outcome: AuditOutcome.success,
            safeMetadata: {
              ...(input.status === undefined ? {} : { status: input.status }),
              ...(input.fillMethod === undefined
                ? {}
                : { fillMethod: input.fillMethod }),
            },
            targetId: form.publicId,
          });
          return form;
        });
        const [activeDraftCount, submissionCount] = await Promise.all([
          prisma.response.count({
            where: { formId: updated.id, status: ResponseStatus.draft },
          }),
          prisma.submission.count({ where: { formId: updated.id } }),
        ]);
        return {
          form: formDto(updated, { activeDraftCount, submissionCount }),
        };
      } catch (error) {
        try {
          await createFormFailureAudit({
            action: auditAction,
            actorId: identity.id,
            error,
            targetId: auditTarget,
          });
        } catch {
          // Preserve the route error if the failure audit cannot be persisted.
        }
        throw error;
      }
    })
    .post(
      "/api/admin/forms/:publicId/duplicate",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const auditTarget = publicIdPattern.test(params.publicId)
          ? params.publicId
          : null;
        let duplicateObjectKey: string | undefined;
        try {
          const input = await readJsonRecord(
            request,
            documentActionBodyMaximumBytes
          );
          if (Object.keys(input).length !== 0) {
            fail(400, "invalid_request", "Duplicate requests must be empty");
          }
          if (!publicIdPattern.test(params.publicId)) {
            fail(404, "not_found", "Form was not found");
          }
          const source = await prisma.form.findUnique({
            include: {
              publishedTemplate: {
                include: {
                  manifest: { include: { fields: true } },
                  prefillConfiguration: { include: { fields: true } },
                },
              },
              templateDraft: { include: { fieldRules: true } },
            },
            where: { publicId: params.publicId },
          });
          if (!source) {
            fail(404, "not_found", "Form was not found");
          }
          const sourceDocument =
            source.publishedTemplate ?? source.templateDraft;
          if (!sourceDocument) {
            fail(409, "document_unavailable", "The Form DOCX is unavailable");
          }
          if (!(await objectExists(sourceDocument.objectKey))) {
            fail(409, "document_unavailable", "The Form DOCX is unavailable");
          }
          const sourceBytes = await readObject(sourceDocument.objectKey);
          const duplicateId = crypto.randomUUID();
          const duplicatePublicId = crypto.randomUUID().replaceAll("-", "");
          duplicateObjectKey = objectKey(
            "forms",
            duplicateId,
            "template-draft",
            crypto.randomUUID(),
            "docx"
          );
          const duplicateDocumentKey = `template-${crypto.randomUUID()}`;
          const sourcePrefillFields =
            source.publishedTemplate?.prefillConfiguration?.fields ?? [];
          const sourceRules =
            source.publishedTemplate?.manifest?.fields.map((field) => {
              const prefill = sourcePrefillFields.find(
                (candidate) => candidate.tag === field.tag
              );
              return {
                prefillPointer: prefill?.pointer ?? null,
                prefillPolicy: prefill?.policy ?? field.prefillPolicy,
                required: field.required,
                tag: field.tag,
              };
            }) ??
            source.templateDraft?.fieldRules.map((field) => ({
              prefillPointer: field.prefillPointer,
              prefillPolicy: field.prefillPolicy,
              required: field.required,
              tag: field.tag,
            })) ??
            [];
          await prisma.objectCleanupIntent.create({
            data: {
              cleanupAfter: new Date(Date.now() + objectCleanupIntentGraceMs),
              objectKey: duplicateObjectKey,
            },
          });
          await putObject(duplicateObjectKey, sourceBytes, DOCX_CONTENT_TYPE);
          const duplicate = await prisma.$transaction(async (tx) => {
            const created = await tx.form.create({
              data: {
                creator: { connect: { id: identity.id } },
                description: source.description,
                id: duplicateId,
                publicId: duplicatePublicId,
                templateDraft: {
                  create: {
                    contentHash:
                      sourceDocument.contentHash ?? contentHash(sourceBytes),
                    documentKey: duplicateDocumentKey,
                    fieldRules: { create: sourceRules },
                    objectKey: duplicateObjectKey as string,
                  },
                },
                title: source.title,
              },
              include: { publishedTemplate: true, templateDraft: true },
            });
            await createFormAudit(tx, {
              action: "duplicate_form",
              actorId: identity.id,
              outcome: AuditOutcome.success,
              safeMetadata: { sourcePublicId: source.publicId },
              targetId: duplicatePublicId,
            });
            await tx.objectCleanupIntent.delete({
              where: { objectKey: duplicateObjectKey },
            });
            return created;
          });
          return {
            form: formDto(duplicate, {
              activeDraftCount: 0,
              submissionCount: 0,
            }),
          };
        } catch (error) {
          if (duplicateObjectKey) {
            await prisma.objectCleanupIntent.updateMany({
              data: { cleanupAfter: new Date() },
              where: { objectKey: duplicateObjectKey },
            });
            await drainObjectCleanupIntents([duplicateObjectKey]);
          }
          try {
            await createFormFailureAudit({
              action: "duplicate_form",
              actorId: identity.id,
              error,
              targetId: auditTarget,
            });
          } catch {
            // Preserve the route error if the failure audit cannot be persisted.
          }
          throw error;
        }
      }
    )
    .delete("/api/admin/forms/:publicId", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const publicId = params.publicId;
      const auditTarget = publicIdPattern.test(publicId) ? publicId : null;
      try {
        const { objectKeys: objectKeysToDelete } = await prisma.$transaction(
          async (tx) => {
            if (!publicIdPattern.test(publicId)) {
              fail(404, "not_found", "Form was not found");
            }
            const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
              Prisma.sql`
                  SELECT "id"
                  FROM "forms"
                  WHERE "public_id" = ${publicId}
                  FOR UPDATE
                `
            );
            if (!lockedForm) {
              fail(404, "not_found", "Form was not found");
            }
            const form = await tx.form.findUnique({
              include: {
                publishedTemplate: true,
                templateDraft: true,
              },
              where: { id: lockedForm.id },
            });
            if (!form) {
              fail(404, "not_found", "Form was not found");
            }
            if (form.status !== FormStatus.draft || form.version > 0) {
              fail(
                409,
                "form_not_draft",
                "Only unpublished draft forms can be removed"
              );
            }
            if (form.templateDraft) {
              const now = new Date();
              const activeLease = await tx.editorLease.findFirst({
                select: {
                  holderSessionId: true,
                  holderUserId: true,
                },
                where: {
                  expiresAt: { gt: now },
                  holderSession: { expiresAt: { gt: now } },
                  targetId: form.templateDraft.id,
                  targetType: OperationTargetType.template_draft,
                },
              });
              const competingLease =
                activeLease &&
                (activeLease.holderUserId !== identity.id ||
                  activeLease.holderSessionId !== identity.sessionId);
              if (competingLease) {
                fail(
                  409,
                  "editor_in_use",
                  "Another Admin is editing this Template Draft"
                );
              }
            }

            const activeOperations = await tx.operation.findMany({
              select: {
                actorId: true,
                id: true,
                metadata: true,
                stagingObjectKey: true,
                updatedAt: true,
              },
              where: {
                formId: form.id,
                status: {
                  in: [OperationStatus.pending, OperationStatus.processing],
                },
              },
            });
            for (const operation of activeOperations) {
              if (
                Date.now() - operation.updatedAt.getTime() <
                operationTimeoutMs
              ) {
                fail(
                  409,
                  "operation_in_progress",
                  "Wait for the draft operation to finish before removing this form"
                );
              }
              const failed = await tx.operation.updateMany({
                data: {
                  errorCode: "operation_timeout",
                  status: OperationStatus.failed,
                  updatedAt: new Date(),
                },
                where: {
                  id: operation.id,
                  status: {
                    in: [OperationStatus.pending, OperationStatus.processing],
                  },
                },
              });
              if (failed.count === 1) {
                const metadata = asRecord(operation.metadata);
                if (
                  metadata.action === "save-template" ||
                  metadata.action === "publish"
                ) {
                  await createFormAudit(tx, {
                    action:
                      metadata.action === "publish"
                        ? "publish_form"
                        : "save_template_draft",
                    actorId: operation.actorId,
                    outcome: AuditOutcome.failure,
                    safeMetadata: { errorCode: "operation_timeout" },
                    targetId: form.publicId,
                  });
                }
              }
            }

            const [responseCount, snapshotCount, submissionCount] =
              await Promise.all([
                tx.response.count({ where: { formId: form.id } }),
                tx.prefillSnapshot.count({ where: { formId: form.id } }),
                tx.submission.count({ where: { formId: form.id } }),
              ]);
            if (responseCount > 0 || snapshotCount > 0 || submissionCount > 0) {
              fail(
                409,
                "form_has_responses",
                "A form with responses cannot be removed"
              );
            }

            const formOperations = await tx.operation.findMany({
              select: { metadata: true, stagingObjectKey: true },
              where: { formId: form.id },
            });
            const objectKeys = uniqueObjectKeys([
              form.templateDraft?.objectKey,
              form.publishedTemplate?.objectKey,
              ...formOperations.flatMap((operation) => {
                const metadata = asRecord(operation.metadata);
                return [
                  operation.stagingObjectKey,
                  typeof metadata.finalObjectKey === "string"
                    ? metadata.finalObjectKey
                    : undefined,
                  ...(Array.isArray(metadata.cleanupObjectKeys)
                    ? metadata.cleanupObjectKeys.filter(
                        (key): key is string => typeof key === "string"
                      )
                    : []),
                ];
              }),
            ]);
            if (objectKeys.length > 0) {
              await tx.objectCleanupIntent.createMany({
                data: objectKeys.map((objectKeyValue) => ({
                  objectKey: objectKeyValue,
                })),
                skipDuplicates: true,
              });
            }
            if (form.templateDraft) {
              await tx.editorLease.deleteMany({
                where: {
                  targetId: form.templateDraft.id,
                  targetType: OperationTargetType.template_draft,
                },
              });
            }
            await tx.operation.deleteMany({ where: { formId: form.id } });
            await createFormAudit(tx, {
              action: "delete_form",
              actorId: identity.id,
              outcome: AuditOutcome.success,
              safeMetadata: {},
              targetId: form.publicId,
            });
            const deleted = await tx.form.deleteMany({
              where: {
                id: form.id,
                status: FormStatus.draft,
                version: 0,
              },
            });
            if (deleted.count !== 1) {
              fail(409, "form_not_draft", "Only draft forms can be removed");
            }
            return { objectKeys };
          },
          { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
        );
        await drainObjectCleanupIntents(objectKeysToDelete);
        return { deleted: true };
      } catch (error) {
        try {
          await createFormFailureAudit({
            action: "delete_form",
            actorId: identity.id,
            error,
            targetId: auditTarget,
          });
        } catch {
          // Preserve the route error if the failure audit cannot be persisted.
        }
        throw error;
      }
    })
    .post(
      "/api/admin/forms",
      async ({ request }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const id = crypto.randomUUID();
        const publicId = crypto.randomUUID().replaceAll("-", "");
        let source: FormSource | undefined;
        let templateObjectKey: string | undefined;
        try {
          const input = await readTemplateCreationInput(request);
          source = input.source;
          let templateBytes: Uint8Array | undefined = input.templateBytes;
          if (input.source === "blank") {
            const sourcePath = await findTemplateSource();
            if (sourcePath) {
              templateBytes = await readTemplateSourceBytes(sourcePath);
            }
          }
          if (!templateBytes) {
            fail(
              409,
              "blank_template_unavailable",
              "The configured blank template is unavailable"
            );
          }
          if (input.source === "upload") {
            validateTemplatePackage(templateBytes);
          }
          templateObjectKey = objectKey(
            "forms",
            id,
            "template-draft",
            crypto.randomUUID(),
            "docx"
          );
          const templateDocumentKey = `template-${crypto.randomUUID()}`;
          await prisma.objectCleanupIntent.create({
            data: {
              cleanupAfter: new Date(Date.now() + objectCleanupIntentGraceMs),
              objectKey: templateObjectKey,
            },
          });
          await putObject(templateObjectKey, templateBytes, DOCX_CONTENT_TYPE);
          const form = await prisma.$transaction(async (tx) => {
            const created = await tx.form.create({
              data: {
                creator: { connect: { id: identity.id } },
                description: input.description,
                id,
                publicId,
                templateDraft: {
                  create: {
                    contentHash: contentHash(templateBytes),
                    documentKey: templateDocumentKey,
                    objectKey: templateObjectKey as string,
                  },
                },
                title: input.title,
              },
              include: { publishedTemplate: true, templateDraft: true },
            });
            await createFormAudit(tx, {
              action: "create_form",
              actorId: identity.id,
              outcome: AuditOutcome.success,
              safeMetadata: { source: input.source },
              targetId: publicId,
            });
            await tx.objectCleanupIntent.delete({
              where: { objectKey: templateObjectKey },
            });
            return created;
          });
          return {
            form: formDto(form, {
              activeDraftCount: 0,
              submissionCount: 0,
            }),
          };
        } catch (error) {
          if (templateObjectKey) {
            await prisma.objectCleanupIntent.updateMany({
              data: { cleanupAfter: new Date() },
              where: { objectKey: templateObjectKey },
            });
            await drainObjectCleanupIntents([templateObjectKey]);
          }
          try {
            await createFormFailureAudit({
              action: "create_form",
              actorId: identity.id,
              error,
              source,
              targetId: null,
            });
          } catch {
            // Preserve the route error if the failure audit cannot be persisted.
          }
          throw error;
        }
      },
      { parse: "none" }
    )
    .get("/api/admin/forms/:publicId", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const form = await findFormByPublicId(params.publicId);
      const [activeDraftCount, submissionCount] = await Promise.all([
        prisma.response.count({
          where: { formId: form.id, status: ResponseStatus.draft },
        }),
        prisma.submission.count({ where: { formId: form.id } }),
      ]);
      const manifest = form.publishedTemplate
        ? await prisma.fieldManifest.findUnique({
            include: { fields: { select: { tag: true, type: true } } },
            where: {
              publishedTemplateId: form.publishedTemplate.id,
            },
          })
        : null;
      const nativeFillAvailable =
        manifest && form.publishedTemplate
          ? await supportsNativeTemplate(
              form.publishedTemplate.objectKey,
              manifest.fields
            )
          : false;
      return {
        editorConfigUrl: `/api/admin/forms/${form.publicId}/editor-config`,
        form: {
          ...formDto(form, { activeDraftCount, submissionCount }),
          nativeFillAvailable,
        },
      };
    })
    .get(
      "/api/admin/forms/:publicId/editor-config",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const form = await findFormByPublicId(params.publicId);
        if (form.status === FormStatus.published || form.publishedTemplate) {
          fail(
            409,
            "published_immutable",
            "Published forms cannot be structurally edited"
          );
        }
        const templateDraft = form.templateDraft;
        if (!templateDraft) {
          fail(
            409,
            "document_unavailable",
            "No template DOCX is configured; provide TEMPLATE_PATH or upload a template"
          );
        }
        if (!(await objectExists(templateDraft.objectKey))) {
          fail(
            409,
            "document_unavailable",
            "The template DOCX artifact is unavailable"
          );
        }
        const capabilityScope = {
          documentKey: templateDraft.documentKey,
          formId: form.id,
          targetId: templateDraft.id,
          targetType: "template-draft",
        } as const;
        const lease = await claimEditorLease(
          identity,
          capabilityScope.targetType,
          capabilityScope.targetId,
          form.id
        );
        return editorConfig(
          {
            action: "template-edit",
            capabilities: {
              "configure-fields": actionEditorCapability(
                identity,
                capabilityScope,
                "configure-fields",
                lease
              ),
              publish: actionEditorCapability(
                identity,
                capabilityScope,
                "publish",
                lease
              ),
              "save-template": actionEditorCapability(
                identity,
                capabilityScope,
                "save-template",
                lease
              ),
            },
            documentKey: templateDraft.documentKey,
            lease: editorLeaseBridge(lease),
            publicId: form.publicId,
          },
          identity
        );
      }
    )
    .get(
      "/api/admin/forms/:publicId/schema",
      async ({ request, params, query }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        requireAdmin(identity);
        const form = await findFormByPublicId(params.publicId);
        if (form.status === FormStatus.published || form.publishedTemplate) {
          fail(
            409,
            "published_immutable",
            "Published forms cannot change Field rules"
          );
        }
        const templateDraft = requireTemplateDraft(form);
        const capabilityScope = fieldRuleCapabilityScope(form, templateDraft);
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "configure-fields",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const queryRecord = query as unknown as JsonRecord;
        return schemaPage(queryRecord.q, queryRecord.cursor);
      }
    )
    .get(
      "/api/admin/forms/:publicId/field-rules",
      async ({ request, params }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        requireAdmin(identity);
        const form = await findFormByPublicId(params.publicId);
        if (form.status === FormStatus.published || form.publishedTemplate) {
          fail(
            409,
            "published_immutable",
            "Published forms cannot change Field rules"
          );
        }
        const templateDraft = requireTemplateDraft(form);
        const capabilityScope = fieldRuleCapabilityScope(form, templateDraft);
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "configure-fields",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const rules = await prisma.draftFieldRule.findMany({
          orderBy: { tag: "asc" },
          select: {
            prefillPointer: true,
            prefillPolicy: true,
            required: true,
            tag: true,
          },
          where: { templateDraftId: templateDraft.id },
        });
        return { rules: rules.map(fieldRuleDto) };
      }
    )
    .patch(
      "/api/admin/forms/:publicId/field-rules",
      async ({ request, params }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        requireAdmin(identity);
        const form = await findFormByPublicId(params.publicId);
        if (form.status === FormStatus.published || form.publishedTemplate) {
          fail(
            409,
            "published_immutable",
            "Published forms cannot change Field rules"
          );
        }
        const templateDraft = requireTemplateDraft(form);
        const capabilityScope = fieldRuleCapabilityScope(form, templateDraft);
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "configure-fields",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const input = fieldRuleInput(
          await readJsonRecord(request, fieldRuleBodyMaximumBytes)
        );
        if (input.documentKey !== templateDraft.documentKey) {
          fail(
            409,
            "stale_document",
            "The editor document is no longer current"
          );
        }
        validateFieldRulePointer(input.prefillPointer);
        let rule: DraftFieldRule;
        try {
          rule = await prisma.$transaction(
            async (tx) => {
              await lockActiveEditorLease(tx, authorization, capabilityScope);
              const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
                Prisma.sql`
                  SELECT "id"
                  FROM "forms"
                  WHERE "id" = ${form.id}::uuid
                  FOR UPDATE
                `
              );
              if (!lockedForm) {
                fail(404, "not_found", "Form was not found");
              }
              const currentForm = await tx.form.findUnique({
                select: {
                  publishedTemplate: { select: { id: true } },
                  status: true,
                  templateDraft: {
                    select: { documentKey: true, id: true },
                  },
                  version: true,
                },
                where: { id: form.id },
              });
              if (
                !currentForm ||
                currentForm.status === FormStatus.published ||
                currentForm.publishedTemplate
              ) {
                fail(
                  409,
                  "published_immutable",
                  "Published forms cannot change Field rules"
                );
              }
              if (
                currentForm.version !== form.version ||
                currentForm.templateDraft?.id !== templateDraft.id ||
                currentForm.templateDraft?.documentKey !==
                  templateDraft.documentKey
              ) {
                fail(
                  409,
                  "stale_document",
                  "The editor document is no longer current"
                );
              }
              const previousRule = input.previousTag
                ? await tx.draftFieldRule.findUnique({
                    where: {
                      templateDraftId_tag: {
                        tag: input.previousTag,
                        templateDraftId: templateDraft.id,
                      },
                    },
                  })
                : null;
              if (input.previousTag !== null && !previousRule) {
                fail(
                  400,
                  "invalid_field_selection",
                  "previousTag does not identify a configured field"
                );
              }
              const targetRule = await tx.draftFieldRule.findUnique({
                where: {
                  templateDraftId_tag: {
                    tag: input.tag,
                    templateDraftId: templateDraft.id,
                  },
                },
              });
              if (targetRule && targetRule.id !== previousRule?.id) {
                fail(
                  409,
                  "field_rule_conflict",
                  "The field tag is already configured"
                );
              }
              const pointerRule =
                input.prefillPointer === null
                  ? null
                  : await tx.draftFieldRule.findFirst({
                      where: {
                        prefillPointer: input.prefillPointer,
                        templateDraftId: templateDraft.id,
                      },
                    });
              if (pointerRule && pointerRule.id !== previousRule?.id) {
                fail(
                  409,
                  "field_rule_conflict",
                  "The schema pointer is already configured"
                );
              }
              const data = {
                prefillPointer: input.prefillPointer,
                prefillPolicy: fieldRulePrefillPolicy(input.prefillPolicy),
                required: input.required,
                tag: input.tag,
              };
              let persistedRule: DraftFieldRule;
              if (previousRule && previousRule.tag !== input.tag) {
                await tx.draftFieldRule.delete({
                  where: { id: previousRule.id },
                });
                persistedRule = await tx.draftFieldRule.create({
                  data: { ...data, templateDraftId: templateDraft.id },
                });
              } else if (previousRule) {
                persistedRule = await tx.draftFieldRule.update({
                  data,
                  where: { id: previousRule.id },
                });
              } else {
                persistedRule = await tx.draftFieldRule.create({
                  data: { ...data, templateDraftId: templateDraft.id },
                });
              }
              await createFormAudit(tx, {
                action: "configure_field_rule",
                actorId: identity.id,
                outcome: AuditOutcome.success,
                safeMetadata: {},
                targetId: form.publicId,
              });
              return persistedRule;
            },
            { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
          );
        } catch (error) {
          try {
            await createFormFailureAudit({
              action: "configure_field_rule",
              actorId: identity.id,
              error,
              targetId: form.publicId,
            });
          } catch {
            // Preserve the route error if the failure audit cannot be persisted.
          }
          if (databaseErrorCode(error) === "P2002") {
            fail(
              409,
              "field_rule_conflict",
              "The field tag or schema pointer is already configured"
            );
          }
          throw error;
        }
        return { rule: fieldRuleDto(rule) };
      },
      { parse: "none" }
    )
    .post(
      "/api/admin/forms/:publicId/save",
      async ({ request, params, set }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        requireAdmin(identity);
        const publicId = params.publicId;
        try {
          const form = await findFormByPublicId(publicId);
          if (form.status === FormStatus.published || form.publishedTemplate) {
            fail(
              409,
              "published_immutable",
              "Published forms cannot be structurally edited"
            );
          }
          const templateDraft = form.templateDraft;
          if (!templateDraft) {
            fail(409, "document_unavailable", "No template DOCX is configured");
          }
          const capabilityScope = {
            documentKey: templateDraft.documentKey,
            formId: form.id,
            targetId: templateDraft.id,
            targetType: "template-draft",
          } as const;
          requireEditorScope(authorization, {
            ...capabilityScope,
            action: "save-template",
          });
          await requireActiveEditorLease(authorization, capabilityScope);
          const documentKey = await readDocumentKeyInput(request);
          if (templateDraft.documentKey !== documentKey) {
            fail(
              409,
              "stale_document",
              "The editor document is no longer current"
            );
          }
          if (await activeOperationForForm(form.id)) {
            fail(
              409,
              "operation_in_progress",
              "A template save is already in progress"
            );
          }
          const operationId = crypto.randomUUID();
          const nextDocumentKey = `template-${crypto.randomUUID()}`;
          const stagedObjectKey = objectKey(
            "operations",
            operationId,
            "template",
            crypto.randomUUID(),
            "docx"
          );
          const metadata: OperationMetadata = {
            action: "save-template",
            finalObjectKey: objectKey(
              "forms",
              form.id,
              "template-draft",
              crypto.randomUUID(),
              "docx"
            ),
            formId: form.id,
            nextDocumentKey,
            publicId: form.publicId,
            result: { documentKey: nextDocumentKey },
            stagedObjectKey,
          };
          const operation = await createOperation({
            actorId: identity.id,
            authorization,
            capabilityScope,
            documentKey,
            formId: form.id,
            metadata,
            ownerUserId: identity.id,
            stagingObjectKey: stagedObjectKey,
            targetId: templateDraft.id,
            targetType: OperationTargetType.template_draft,
            type: operationTypeForAction["save-template"],
          });
          set.status = 202;
          launchForceSave(operation, onlyOffice, allowedCallbackOrigins);
          return {
            operationCapability: operationEditorCapability(
              identity,
              capabilityScope,
              operation.id
            ),
            operationId: operation.id,
            status: operation.status,
          };
        } catch (error) {
          try {
            await createFormFailureAudit({
              action: "save_template_draft",
              actorId: identity.id,
              error,
              targetId: publicIdPattern.test(publicId) ? publicId : null,
            });
          } catch {
            // Preserve the route error if the failure audit cannot be persisted.
          }
          throw error;
        }
      },
      { parse: "none" }
    )
    .post(
      "/api/admin/forms/:publicId/publish",
      async ({ request, params, set }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        requireAdmin(identity);
        try {
          const form = await findFormByPublicId(params.publicId);
          if (form.status === FormStatus.published || form.publishedTemplate) {
            fail(
              409,
              "published_immutable",
              "A Published Template already exists for this Form"
            );
          }
          const templateDraft = form.templateDraft;
          if (!templateDraft) {
            fail(409, "document_unavailable", "No template DOCX is configured");
          }
          const capabilityScope = {
            documentKey: templateDraft.documentKey,
            formId: form.id,
            targetId: templateDraft.id,
            targetType: "template-draft",
          } as const;
          requireEditorScope(authorization, {
            ...capabilityScope,
            action: "publish",
          });
          await requireActiveEditorLease(authorization, capabilityScope);
          const documentKey = await readDocumentKeyInput(request);
          if (templateDraft.documentKey !== documentKey) {
            fail(
              409,
              "stale_document",
              "The editor document is no longer current"
            );
          }
          if (await activeOperationForForm(form.id)) {
            fail(
              409,
              "operation_in_progress",
              "A publish operation is already in progress"
            );
          }
          const operationId = crypto.randomUUID();
          const version = form.version + 1;
          const publishedKey = `published-${version}-${crypto.randomUUID()}`;
          const stagedObjectKey = objectKey(
            "operations",
            operationId,
            "published",
            crypto.randomUUID(),
            "docx"
          );
          const metadata: OperationMetadata = {
            action: "publish",
            finalObjectKey: objectKey(
              "forms",
              form.id,
              "published",
              String(version),
              crypto.randomUUID(),
              "docx"
            ),
            formId: form.id,
            publicId: form.publicId,
            publishedKey,
            publishedVersion: version,
            stagedObjectKey,
          };
          const operation = await createOperation({
            actorId: identity.id,
            authorization,
            capabilityScope,
            documentKey,
            formId: form.id,
            metadata,
            ownerUserId: identity.id,
            stagingObjectKey: stagedObjectKey,
            targetId: templateDraft.id,
            targetType: OperationTargetType.template_draft,
            type: operationTypeForAction.publish,
          });
          set.status = 202;
          launchForceSave(operation, onlyOffice, allowedCallbackOrigins);
          return {
            operationCapability: operationEditorCapability(
              identity,
              capabilityScope,
              operation.id
            ),
            operationId: operation.id,
            status: operation.status,
          };
        } catch (error) {
          try {
            await createFormFailureAudit({
              action: "publish_form",
              actorId: identity.id,
              error,
              targetId: publicIdPattern.test(params.publicId)
                ? params.publicId
                : null,
            });
          } catch {
            // Preserve the route error if the failure audit cannot be persisted.
          }
          throw error;
        }
      },
      { parse: "none" }
    )
    .get(
      "/api/admin/forms/:publicId/submissions",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const form = await findFormByPublicId(params.publicId);
        const submissions = await prisma.submission.findMany({
          include: { form: true, owner: true },
          orderBy: { createdAt: "desc" },
          where: { formId: form.id },
        });
        return {
          submissions: submissions.map((submission) =>
            submissionSummary(submission, {
              formPublicId: submission.form.publicId,
              formTitle: submission.form.title,
              userEmail: submission.owner.email,
            })
          ),
        };
      }
    )
    .get("/api/admin/audit-events", async ({ request, query }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const queryRecord = query as unknown as JsonRecord;
      const cursor = auditEventCursor(queryString(queryRecord, "cursor"));
      const actorId = auditFilterValue(
        queryString(queryRecord, "actor"),
        "actor",
        64
      );
      const action = auditFilterValue(
        queryString(queryRecord, "action"),
        "action",
        80
      );
      const targetId = auditFilterValue(
        queryString(queryRecord, "target"),
        "target",
        200
      );
      const targetType = auditFilterValue(
        queryString(queryRecord, "targetType"),
        "targetType",
        80
      );
      const outcome = auditOutcomeValue(queryString(queryRecord, "outcome"));
      const from = adminResultDate(queryString(queryRecord, "from"), "from");
      const to = adminResultDate(queryString(queryRecord, "to"), "to");
      if (from && to && from > to) {
        fail(400, "invalid_request", "from must be before to");
      }
      if (actorId && !idPattern.test(actorId)) {
        fail(400, "invalid_request", "actor is invalid");
      }
      const and: Prisma.AuditEventWhereInput[] = [];
      if (cursor) {
        and.push({
          OR: [
            { createdAt: { lt: cursor.createdAt } },
            { createdAt: cursor.createdAt, id: { lt: cursor.id } },
          ],
        });
      }
      if (actorId) {
        and.push({ actorId });
      }
      if (action) {
        and.push({ action });
      }
      if (targetId) {
        and.push({ targetId });
      }
      if (targetType) {
        and.push({ targetType });
      }
      if (outcome) {
        and.push({ outcome });
      }
      if (from || to) {
        and.push({
          createdAt: {
            ...(from ? { gte: from } : {}),
            ...(to ? { lte: to } : {}),
          },
        });
      }
      const events = await prisma.auditEvent.findMany({
        orderBy: [{ createdAt: "desc" }, { id: "desc" }],
        select: {
          action: true,
          actorId: true,
          createdAt: true,
          id: true,
          outcome: true,
          safeMetadata: true,
          targetId: true,
          targetType: true,
        },
        take: auditEventPageSize + 1,
        where: and.length > 0 ? { AND: and } : {},
      });
      const page = events.slice(0, auditEventPageSize);
      return {
        events: page.map((event) => ({
          action: event.action,
          actorId: event.actorId,
          createdAt: event.createdAt,
          id: event.id,
          outcome: event.outcome,
          safeMetadata: safeAuditMetadata(event.safeMetadata),
          targetId: event.targetId,
          targetType: event.targetType,
        })),
        nextCursor:
          events.length > auditEventPageSize
            ? auditEventCursorValue(page.at(-1) as AuditEventCursor)
            : null,
      };
    })
    .get("/api/admin/results", async ({ request, query }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const queryRecord = query as unknown as JsonRecord;
      const cursor = adminResultCursor(queryString(queryRecord, "cursor"));
      const formPublicId = queryString(queryRecord, "form");
      const userQuery = queryString(queryRecord, "user");
      const state = queryString(queryRecord, "state");
      const from = adminResultDate(queryString(queryRecord, "from"), "from");
      const to = adminResultDate(queryString(queryRecord, "to"), "to");
      const correctionQuery = queryString(queryRecord, "correction");
      let correction: number | undefined;
      if (correctionQuery !== undefined) {
        correction = Number(correctionQuery);
        if (
          !Number.isInteger(correction) ||
          correction < 0 ||
          correction > 10_000
        ) {
          fail(400, "invalid_request", "correction is invalid");
        }
      }
      if (formPublicId !== undefined && !publicIdPattern.test(formPublicId)) {
        fail(400, "invalid_request", "form is invalid");
      }
      if (state !== undefined && state !== "draft" && state !== "submitted") {
        fail(400, "invalid_request", "state is invalid");
      }
      const and: Prisma.ResponseWhereInput[] = [];
      if (cursor) {
        and.push({
          OR: [
            { updatedAt: { lt: cursor.updatedAt } },
            { id: { lt: cursor.id }, updatedAt: cursor.updatedAt },
          ],
        });
      }
      if (formPublicId) {
        and.push({ form: { publicId: formPublicId } });
      }
      if (userQuery) {
        and.push({ owner: { email: { contains: normalizeEmail(userQuery) } } });
      }
      if (state === "draft") {
        and.push({
          status: { in: [ResponseStatus.draft, ResponseStatus.submitting] },
        });
      } else if (state === "submitted") {
        and.push({ status: ResponseStatus.submitted });
      }
      if (from || to) {
        and.push({
          updatedAt: {
            ...(from ? { gte: from } : {}),
            ...(to ? { lte: to } : {}),
          },
        });
      }
      if (correction !== undefined) {
        and.push(
          correction === 0
            ? { corrections: { none: {} } }
            : { corrections: { some: { revision: correction } } },
          { corrections: { none: { revision: { gt: correction } } } }
        );
      }
      const responses = await prisma.response.findMany({
        include: {
          corrections: {
            orderBy: { revision: "desc" },
            select: { revision: true },
            take: 1,
          },
          form: { select: { publicId: true, title: true } },
          owner: { select: { email: true } },
          submission: { select: { createdAt: true, id: true } },
        },
        orderBy: [{ updatedAt: "desc" }, { id: "desc" }],
        take: adminResultPageSize + 1,
        where: and.length > 0 ? { AND: and } : {},
      });
      const page = responses.slice(0, adminResultPageSize);
      return {
        nextCursor:
          responses.length > adminResultPageSize
            ? adminResultCursorValue(page.at(-1) as (typeof page)[number])
            : null,
        results: page.map(adminResultSummary),
      };
    })
    .get("/api/admin/results/:id", async ({ request, params, query }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      validateId(params.id, "Response");
      const selector = revisionSelector(query.revision);
      const response = await prisma.response.findUnique({
        include: {
          corrections: {
            orderBy: { revision: "desc" },
          },
          form: { select: { publicId: true, title: true } },
          owner: { select: { email: true } },
          publishedTemplate: {
            include: {
              manifest: {
                include: {
                  fields: { orderBy: { position: "asc" } },
                },
              },
            },
          },
          submission: true,
        },
        where: { id: params.id },
      });
      if (!response) {
        fail(404, "not_found", "Response was not found");
      }
      const submitted = response.status === ResponseStatus.submitted;
      if (!submitted && selector !== "latest") {
        fail(404, "not_found", "Submission revision was not found");
      }
      const latestCorrection = submitted ? response.corrections[0] : undefined;
      const selected = submitted
        ? response.submission
          ? selectedSubmissionRevision(
              { ...response.submission, corrections: response.corrections },
              selector
            )
          : fail(404, "not_found", "Response was not found")
        : undefined;
      await createResponseAudit({
        action: "view_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: {
          revision: selected?.revision ?? 0,
          state: submitted ? "submitted" : "draft",
        },
        targetId: response.id,
        targetType: "response",
      });
      if (selected?.correction) {
        await createResponseAudit({
          action: "view_correction",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {
            revision: selected.correction.revision,
            state: "submitted",
          },
          targetId: selected.correction.id,
          targetType: "correction",
        });
      }
      const documentObjectKey = submitted
        ? selected?.objectKey
        : response.draftObjectKey;
      const documentAvailable = submitted
        ? Boolean(documentObjectKey && (await objectExists(documentObjectKey)))
        : Boolean(
            response.draftDocumentKey &&
            documentObjectKey &&
            (await objectExists(documentObjectKey))
          );
      const publishedTemplate = response.publishedTemplate;
      const manifest = publishedTemplate?.manifest;
      const manifestFields =
        manifest && publishedTemplate
          ? await manifestFieldsWithDocumentMetadata(
              publishedTemplate.objectKey,
              manifest.displayMetadataVersion,
              manifest.fields
            )
          : [];
      return {
        result: {
          correction: selected?.correction
            ? {
                createdAt: selected.correction.createdAt,
                reason: selected.correction.reason,
                revision: selected.correction.revision,
              }
            : null,
          createdAt: response.createdAt,
          data: submitted
            ? (selected?.data ?? {})
            : jsonRecord(response.draftData ?? {}),
          document: {
            available: documentAvailable,
            state: selected?.correction
              ? "correction"
              : submitted
                ? "submission"
                : "draft",
          },
          fields: manifestFields.map(receiptField),
          formPublicId: response.form.publicId,
          formTitle: response.form.title,
          id: response.id,
          latestCorrectionNumber: latestCorrection?.revision ?? null,
          revision: submitted ? (selected?.revision ?? 0) : null,
          state: submitted ? "submitted" : "draft",
          submissionId: response.submission?.id ?? null,
          submittedAt: response.submission?.createdAt ?? null,
          updatedAt: response.updatedAt,
          userEmail: response.owner.email,
        },
      };
    })
    .get(
      "/api/admin/results/:id/viewer-config",
      async ({ request, params, query }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        validateId(params.id, "Response");
        const selector = revisionSelector(query.revision);
        const response = await prisma.response.findUnique({
          include: {
            corrections: { orderBy: { revision: "desc" } },
            form: { select: { publicId: true } },
            submission: true,
          },
          where: { id: params.id },
        });
        if (
          !response ||
          (response.status === ResponseStatus.submitted && !response.submission)
        ) {
          fail(404, "not_found", "Response was not found");
        }
        const submitted = response.status === ResponseStatus.submitted;
        if (!submitted && selector !== "latest") {
          fail(404, "not_found", "Submission revision was not found");
        }
        const selected =
          submitted && response.submission
            ? selectedSubmissionRevision(
                { ...response.submission, corrections: response.corrections },
                selector
              )
            : undefined;
        const selectedCorrection = selected?.correction;
        const documentKey = submitted
          ? selected?.documentKey
          : response.draftDocumentKey;
        const documentObjectKey = submitted
          ? selected?.objectKey
          : response.draftObjectKey;
        if (
          !documentKey ||
          !documentObjectKey ||
          !(await objectExists(documentObjectKey))
        ) {
          fail(409, "document_unavailable", "Response document is unavailable");
        }
        await createResponseAudit({
          action: selectedCorrection ? "view_correction" : "view_response",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {
            revision: selected?.revision ?? 0,
            state: submitted ? "submitted" : "draft",
          },
          targetId:
            selectedCorrection?.id ??
            (submitted
              ? (response.submission?.id ?? response.id)
              : response.id),
          targetType: selectedCorrection
            ? "correction"
            : submitted
              ? "submission"
              : "response",
        });
        return readOnlyViewerConfig(documentKey, response.form.publicId, {
          id: identity.id,
          name: identity.name,
        });
      }
    )
    .get(
      "/api/admin/results/:id/correction/editor-config",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        return correctionEditorConfig(params.id, identity);
      }
    )
    .post(
      "/api/admin/results/:id/correction",
      async ({ request, params, body, set }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        requireAdmin(identity);
        validateId(params.id, "Response");
        const input = correctionInput(body);
        const response = await prisma.response.findUnique({
          include: {
            corrections: {
              orderBy: { revision: "desc" },
              take: 1,
            },
            prefillSnapshot: true,
            submission: true,
          },
          where: { id: params.id },
        });
        if (
          !response ||
          response.status !== ResponseStatus.submitted ||
          !response.submission
        ) {
          fail(404, "not_found", "Submitted Response was not found");
        }
        const form = await prisma.form.findUnique({
          include: { publishedTemplate: true, templateDraft: true },
          where: { id: response.formId },
        });
        if (!form) {
          fail(404, "not_found", "Form was not found");
        }
        const latest = response.corrections[0];
        const currentRevision = latest?.revision ?? 0;
        const currentDocumentKey =
          latest?.documentKey ?? response.submission.documentKey;
        const capabilityScope = {
          documentKey: input.documentKey,
          formId: form.id,
          targetId: response.id,
          targetType: "correction",
        } as const;
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "save-correction",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const workspaceLease = await prisma.editorLease.findFirst({
          select: {
            workspaceBaseDocumentKey: true,
            workspaceBaseRevision: true,
            workspaceDocumentKey: true,
            workspaceObjectKey: true,
          },
          where: {
            id: authorization.capability.leaseId,
            targetId: response.id,
            targetType: OperationTargetType.correction,
            workspaceDocumentKey: input.documentKey,
          },
        });
        if (
          !workspaceLease?.workspaceBaseDocumentKey ||
          workspaceLease.workspaceBaseRevision === null ||
          workspaceLease.workspaceDocumentKey !== input.documentKey ||
          !workspaceLease.workspaceObjectKey ||
          workspaceLease.workspaceBaseRevision !== currentRevision ||
          workspaceLease.workspaceBaseDocumentKey !== currentDocumentKey
        ) {
          fail(409, "stale_document", "The response document is stale");
        }
        const baseRevision = workspaceLease.workspaceBaseRevision;
        const baseDocumentKey = workspaceLease.workspaceBaseDocumentKey;
        const previousData = jsonRecord(
          latest?.data ?? response.submission.data
        );
        const data = await normalizeResponseData(
          form,
          response,
          { ...previousData, ...input.data },
          true,
          true
        );
        if (await activeOperationForResponse(response.id)) {
          fail(
            409,
            "operation_in_progress",
            "Another response operation is already in progress"
          );
        }
        const operationId = crypto.randomUUID();
        const stagedObjectKey = objectKey(
          "operations",
          operationId,
          "correction",
          crypto.randomUUID(),
          "docx"
        );
        const metadata: OperationMetadata = {
          action: "save-correction",
          baseDocumentKey,
          baseRevision,
          data,
          finalObjectKey: objectKey(
            "responses",
            response.id,
            "corrections",
            crypto.randomUUID(),
            "docx"
          ),
          formId: form.id,
          nextDocumentKey: `correction-${response.id}-${crypto.randomUUID()}`,
          publicId: form.publicId,
          reason: input.reason,
          ...(form.fillMethod === FillMethod.onlyoffice
            ? callbackFieldMetadata(data, input.data, response.prefillSnapshot)
            : {}),
          responseId: response.id,
          stagedObjectKey,
          submissionId: response.submission.id,
          workspaceDocumentKey: workspaceLease.workspaceDocumentKey,
          workspaceObjectKey: workspaceLease.workspaceObjectKey,
        };

        const operation = await createOperation({
          actorId: identity.id,
          authorization,
          capabilityScope,
          documentKey: input.documentKey,
          formId: form.id,
          metadata,
          ownerUserId: identity.id,
          responseId: response.id,
          stagingObjectKey: stagedObjectKey,
          submissionId: response.submission.id,
          targetId: response.id,
          targetType: OperationTargetType.correction,
          type: operationTypeForAction["save-correction"],
        });
        set.status = 202;
        launchForceSave(operation, onlyOffice, allowedCallbackOrigins);
        return {
          correction: { baseRevision, status: operation.status },
          operationCapability: operationEditorCapability(
            identity,
            capabilityScope,
            operation.id
          ),
          operationId: operation.id,
          responseId: response.id,
          status: operation.status,
        };
      }
    )
    .get("/api/forms/:publicId", async ({ params, request }) => {
      const identity = await requireIdentity(request);
      const form = await findFormByPublicId(params.publicId);
      if (
        !form.publishedTemplate?.objectKey ||
        !form.publishedTemplate.documentKey
      ) {
        fail(404, "not_found", "Form was not found");
      }
      if (form.status === FormStatus.archived) {
        const existingResponse = await prisma.response.findUnique({
          select: { id: true },
          where: {
            formId_userId: { formId: form.id, userId: identity.id },
          },
        });
        if (!existingResponse) {
          fail(404, "not_found", "Form was not found");
        }
      } else if (form.status !== FormStatus.published) {
        fail(404, "not_found", "Form was not found");
      }
      return {
        form: {
          description: form.description ?? "",
          fillMethod: form.fillMethod,
          publicId: form.publicId,
          published: Boolean(
            form.publishedTemplate?.objectKey &&
            form.publishedTemplate.documentKey
          ),
          title: form.title,
          version: form.version,
        },
      };
    })
    .get(
      "/api/forms/:publicId/editor-config",
      async ({ request, params, query }) => {
        const identity = await requireIdentity(request);
        const form = await findFormByPublicId(params.publicId);
        const responseId =
          typeof query.responseId === "string" ? query.responseId : undefined;
        const requestedAction =
          typeof query.action === "string" ? query.action : undefined;
        return userEditorConfig(form, identity, responseId, requestedAction);
      }
    )
    .post("/api/forms/:publicId/start", async ({ request, params, set }) => {
      const identity = await requireIdentity(request);
      const form = await findFormByPublicId(params.publicId);
      const publishedTemplate = form.publishedTemplate;
      const pendingClaimToken = pendingClaimFor(request);
      const prefillConfiguration = publishedTemplate
        ? await prisma.prefillConfiguration.findUnique({
            include: { fields: true },
            where: { publishedTemplateId: publishedTemplate.id },
          })
        : null;
      const hasPrefillConfiguration =
        (prefillConfiguration?.fields.length ?? 0) > 0;
      const existing = await prisma.response.findUnique({
        include: { prefillSnapshot: true, submission: true },
        where: {
          formId_userId: { formId: form.id, userId: identity.id },
        },
      });
      if (
        !publishedTemplate?.objectKey ||
        !publishedTemplate.documentKey ||
        (form.status !== FormStatus.published &&
          !(form.status === FormStatus.archived && existing))
      ) {
        if (pendingClaimToken) {
          handoffUnavailable();
        }
        fail(
          409,
          "form_unavailable",
          "This form is not accepting new responses"
        );
      }
      if (existing?.status === ResponseStatus.submitted) {
        if (pendingClaimToken) {
          try {
            await consumeSubmittedPrefillHandoff(
              form,
              identity,
              pendingClaimToken,
              existing.id,
              handoffClock
            );
            set.headers["Set-Cookie"] = pendingClaimCookie("", 0);
          } catch (error) {
            set.headers["Set-Cookie"] = pendingClaimCookie("", 0);
            throw error;
          }
        }
        if (!existing.submission) {
          fail(500, "internal_error", "The submitted receipt is unavailable");
        }
        return {
          fillMethod: form.fillMethod,
          receiptUrl: `/receipt/${existing.submission.id}`,
          response: responseSummary(existing, {
            submissionId: existing.submission.id,
            submittedAt: existing.submission.createdAt,
          }),
          submissionId: existing.submission.id,
        };
      }
      if (existing?.status === ResponseStatus.submitting) {
        fail(
          409,
          "operation_in_progress",
          "Your submission is being processed"
        );
      }
      if (pendingClaimToken) {
        if (!hasPrefillConfiguration) {
          handoffUnavailable();
        }
        try {
          const redeemed = await redeemPrefillHandoff(
            form,
            identity,
            pendingClaimToken,
            handoffClock
          );
          await deleteObjects(redeemed.cleanupObjectKeys);
          set.headers["Set-Cookie"] = pendingClaimCookie("", 0);
          return {
            editorConfigUrl: `/api/forms/${form.publicId}/editor-config?responseId=${redeemed.response.id}&action=fill`,
            fillMethod: form.fillMethod,
            prefill: {
              data: jsonRecord(redeemed.response.prefillSnapshot?.values),
              editableFields: redeemed.response.prefillSnapshot
                ? editableFieldsForSnapshot(redeemed.response.prefillSnapshot)
                : {},
            },
            response: responseSummary(redeemed.response),
          };
        } catch (error) {
          set.headers["Set-Cookie"] = pendingClaimCookie("", 0);
          throw error;
        }
      }
      if (hasPrefillConfiguration && !existing) {
        prefillRequired();
      }
      if (
        existing?.status === ResponseStatus.draft &&
        existing.publishedVersion === form.version &&
        existing.draftObjectKey &&
        existing.draftDocumentKey
      ) {
        if (!(await objectExists(existing.draftObjectKey))) {
          fail(
            409,
            "document_unavailable",
            "Response document artifact is unavailable"
          );
        }
        return {
          editorConfigUrl: `/api/forms/${form.publicId}/editor-config?responseId=${existing.id}&action=${existing.draftData ? "draft" : "fill"}`,
          fillMethod: form.fillMethod,
          prefill: existing.draftData ? null : undefined,
          response: responseSummary(existing),
        };
      }

      if (!(await objectExists(publishedTemplate.objectKey))) {
        fail(
          409,
          "document_unavailable",
          "The published document artifact is unavailable"
        );
      }
      const document = await readObject(publishedTemplate.objectKey);
      const responseId = existing?.id ?? crypto.randomUUID();
      const draftObjectKey = objectKey(
        "responses",
        responseId,
        "draft",
        crypto.randomUUID(),
        "docx"
      );
      const draftDocumentKey = `response-${responseId}-${crypto.randomUUID()}`;
      const snapshotId = crypto.randomUUID();
      let unusedDraftObjectKey: string | undefined;
      try {
        await putObject(draftObjectKey, document, DOCX_CONTENT_TYPE);
        const response = await prisma.$transaction(
          async (tx) => {
            const [lockedFormRow] = await tx.$queryRaw<{ id: string }[]>(
              Prisma.sql`
                SELECT "id"
                FROM "forms"
                WHERE "id" = ${form.id}::uuid
                FOR UPDATE
              `
            );
            if (!lockedFormRow) {
              fail(404, "not_found", "The form was not found");
            }
            const lockedForm = await tx.form.findUnique({
              include: { publishedTemplate: true },
              where: { id: lockedFormRow.id },
            });
            const current = await tx.response.findUnique({
              where: {
                formId_userId: { formId: form.id, userId: identity.id },
              },
            });
            if (lockedForm?.status === FormStatus.archived && !current) {
              fail(
                409,
                "form_unavailable",
                "This form is not accepting new responses"
              );
            }
            if (
              !lockedForm ||
              (lockedForm.status !== FormStatus.published &&
                !(lockedForm.status === FormStatus.archived && current)) ||
              lockedForm.version !== form.version ||
              !lockedForm.publishedTemplate ||
              lockedForm.publishedTemplate.id !== publishedTemplate.id
            ) {
              fail(409, "stale_form", "The form was published while starting");
            }
            if (current?.status === ResponseStatus.submitted) {
              fail(
                409,
                "already_submitted",
                "You have already submitted this form"
              );
            }
            if (current?.status === ResponseStatus.submitting) {
              fail(
                409,
                "operation_in_progress",
                "Your submission is being processed"
              );
            }
            if (
              current?.status === ResponseStatus.draft &&
              current.publishedVersion === form.version &&
              current.draftDocumentKey &&
              current.draftObjectKey
            ) {
              unusedDraftObjectKey = draftObjectKey;
              return current;
            }
            const responseTargetId = current?.id ?? responseId;
            if (!current) {
              await tx.response.create({
                data: {
                  draftData: Prisma.DbNull,
                  draftDocumentKey,
                  draftObjectKey,
                  form: { connect: { id: form.id } },
                  id: responseTargetId,
                  owner: { connect: { id: identity.id } },
                  publishedTemplate: {
                    connect: { id: lockedForm.publishedTemplate.id },
                  },
                  publishedVersion: form.version,
                  status: ResponseStatus.draft,
                },
              });
            }
            await tx.prefillSnapshot.deleteMany({
              where: { responseId: responseTargetId },
            });
            await tx.prefillSnapshot.create({
              data: {
                form: { connect: { id: form.id } },
                id: snapshotId,
                lockedFields: jsonValue({}),
                owner: { connect: { id: identity.id } },
                response: { connect: { id: responseTargetId } },
                values: jsonValue({}),
              },
            });
            const updated = await tx.response.updateMany({
              data: {
                draftData: Prisma.DbNull,
                draftDocumentKey,
                draftObjectKey,
                publishedTemplateId: lockedForm.publishedTemplate.id,
                publishedVersion: form.version,
                status: ResponseStatus.draft,
                updatedAt: new Date(),
              },
              where: { id: responseTargetId },
            });
            if (updated.count !== 1) {
              fail(500, "start_failed", "Unable to start response");
            }
            return tx.response.findUnique({ where: { id: responseTargetId } });
          },
          { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
        );
        if (!response) {
          fail(500, "start_failed", "Unable to start response");
        }
        await deleteObjects([existing?.draftObjectKey, unusedDraftObjectKey]);
        return {
          editorConfigUrl: `/api/forms/${form.publicId}/editor-config?responseId=${response.id}&action=${response.draftData ? "draft" : "fill"}`,
          fillMethod: form.fillMethod,
          prefill: response.draftData ? null : { data: {}, editableFields: {} },
          response: responseSummary(response),
        };
      } catch (error) {
        await deleteObjectUnlessCanonical(draftObjectKey);
        if (databaseErrorCode(error) === "P2034") {
          const current = await prisma.response.findUnique({
            where: {
              formId_userId: { formId: form.id, userId: identity.id },
            },
          });
          if (
            current?.status === ResponseStatus.draft &&
            current.publishedVersion === form.version &&
            current.draftDocumentKey &&
            current.draftObjectKey &&
            (await objectExists(current.draftObjectKey))
          ) {
            return {
              editorConfigUrl: `/api/forms/${form.publicId}/editor-config?responseId=${current.id}&action=${current.draftData ? "draft" : "fill"}`,
              fillMethod: form.fillMethod,
              prefill: current.draftData
                ? null
                : { data: {}, editableFields: {} },
              response: responseSummary(current),
            };
          }
        }
        throw error;
      }
    })
    .get("/api/responses/me", async ({ request }) => {
      const identity = await requireIdentity(request);
      const responses = await prisma.response.findMany({
        include: {
          corrections: {
            orderBy: { revision: "desc" },
            select: { revision: true },
            take: 1,
          },
          form: true,
          submission: true,
        },
        orderBy: { updatedAt: "desc" },
        where: { userId: identity.id },
      });
      return {
        responses: responses.map((response) =>
          responseSummary(response, {
            formPublicId: response.form.publicId,
            formTitle: response.form.title,
            latestCorrectionNumber: response.corrections[0]?.revision ?? null,
            submissionId: response.submission?.id,
            submittedAt: response.submission?.createdAt,
          })
        ),
      };
    })
    .get("/api/responses/:id/corrections", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Response");
      const response = await prisma.response.findUnique({
        include: {
          corrections: { orderBy: { revision: "desc" } },
          publishedTemplate: {
            include: { manifest: { include: { fields: true } } },
          },
          submission: true,
        },
        where: { id: params.id },
      });
      if (
        !response ||
        response.status !== ResponseStatus.submitted ||
        !response.submission
      ) {
        fail(404, "not_found", "Submitted Response was not found");
      }
      if (identity.role !== "admin" && response.userId !== identity.id) {
        fail(403, "forbidden", "You may only access your own Response");
      }
      const actorIds = response.corrections.map(
        (correction) => correction.actorId
      );
      const actors = await prisma.user.findMany({
        select: { email: true, id: true, name: true },
        where: { id: { in: [...new Set(actorIds)] } },
      });
      const actorById = new Map(
        actors.map((actor) => [actor.id, actor] as const)
      );
      const originalAvailable = await objectExists(
        response.submission.objectKey
      );
      const pictureFields =
        response.publishedTemplate.manifest?.fields.filter(
          (field) => field.type === FieldType.picture
        ) ?? [];
      const originalPictures =
        originalAvailable && pictureFields.length > 0
          ? responsePicturePresence(
              await readObject(response.submission.objectKey),
              pictureFields
            )
          : null;
      const correctionSummaries: CorrectionRevisionSummary[] = [];
      for (const correction of response.corrections.toReversed()) {
        const documentAvailable = await objectExists(correction.objectKey);
        const pictures =
          documentAvailable && pictureFields.length > 0
            ? responsePicturePresence(
                await readObject(correction.objectKey),
                pictureFields
              )
            : null;
        correctionSummaries.push(
          correctionRevisionSummary(
            correction,
            actorById.get(correction.actorId),
            documentAvailable,
            pictures
          )
        );
      }
      await createResponseAudit({
        action: "view_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: { revision: 0, state: "submitted" },
        targetId: response.submission.id,
        targetType: "submission",
      });
      for (const correction of response.corrections) {
        await createResponseAudit({
          action: "view_correction",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {
            revision: correction.revision,
            state: "submitted",
          },
          targetId: correction.id,
          targetType: "correction",
        });
      }
      return {
        latestRevision: response.corrections[0]?.revision ?? 0,
        revisions: [
          {
            actorEmail: null,
            actorName: null,
            createdAt: response.submission.createdAt,
            data: jsonRecord(response.submission.data),
            document: {
              available: originalAvailable,
              state: "submission",
            },
            id: null,
            pictures: originalPictures,
            reason: null,
            revision: 0,
          },
          ...correctionSummaries.map((correction) => ({
            actorEmail: correction.actorEmail,
            actorName: correction.actorName,
            createdAt: correction.createdAt,
            data: correction.data,
            document: {
              available: correction.documentAvailable,
              state: "correction",
            },
            id: correction.id,
            pictures: correction.pictures,
            reason: correction.reason,
            revision: correction.revision,
          })),
        ],
      };
    })
    .get(
      "/api/responses/:id/draft/:format",
      async ({ request, params, set }) => {
        const identity = await requireIdentity(request);
        validateId(params.id, "Response");
        if (
          params.format !== "json" &&
          params.format !== "docx" &&
          params.format !== "pdf"
        ) {
          fail(404, "not_found", "Draft export was not found");
        }
        const response = await prisma.response.findUnique({
          where: { id: params.id },
        });
        if (!response) {
          fail(404, "not_found", "Response was not found");
        }
        if (response.userId !== identity.id) {
          fail(403, "forbidden", "You may only export your own Draft");
        }
        if (response.status !== ResponseStatus.draft) {
          fail(409, "draft_unavailable", "Only a Draft can be exported");
        }
        if (await activeOperationForResponse(response.id)) {
          fail(409, "operation_in_progress", "The Draft is still being saved");
        }
        if (params.format === "json") {
          return Response.json(jsonRecord(response.draftData ?? {}), {
            headers: {
              "Content-Disposition": `attachment; filename="response-${response.id}.json"`,
              "Content-Type": "application/json; charset=utf-8",
            },
          });
        }
        if (!response.draftDocumentKey || !response.draftObjectKey) {
          fail(409, "document_unavailable", "Draft document is unavailable");
        }
        if (!(await objectExists(response.draftObjectKey))) {
          fail(404, "document_unavailable", "Draft document is unavailable");
        }
        if (params.format === "docx") {
          return new Response(streamObject(response.draftObjectKey), {
            headers: {
              "Content-Disposition": `attachment; filename="response-${response.id}.docx"`,
              "Content-Type": DOCX_CONTENT_TYPE,
            },
          });
        }
        const pdf = await onlyOffice.convertDocxToPdf(
          response.draftDocumentKey,
          true
        );
        set.headers["Content-Disposition"] =
          `attachment; filename="response-${response.id}.pdf"`;
        set.headers["Content-Type"] = "application/pdf";
        return pdf;
      }
    )
    .delete(
      "/api/admin/responses/:id",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const responseId = validateId(params.id, "Response");
        try {
          const input = await readJsonRecord(request, accountBodyMaximumBytes);
          if (Object.keys(input).length !== 1 || input.confirm !== true) {
            fail(400, "invalid_request", "confirm must be true");
          }
          const result = await deleteResponseData({
            actor: identity,
            aiAuthoring,
            allowActiveLease: false,
            missingOk: false,
            removeObject,
            responseId,
            revokeOwnerSessions: true,
          });
          if (result.alreadyDeleted) {
            await createResponseDeletionAudit({
              actorId: identity.id,
              outcome: AuditOutcome.success,
              targetId: responseId,
            });
          }
          return { deleted: true };
        } catch (error) {
          try {
            await createResponseDeletionAudit({
              actorId: identity.id,
              error,
              outcome: AuditOutcome.failure,
              targetId: responseId,
            });
          } catch {
            // Preserve the deletion error if the failure audit cannot be persisted.
          }
          throw error;
        }
      },
      { parse: "none" }
    )
    .delete("/api/responses/:id", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      const responseId = validateId(params.id, "Response");
      try {
        await deleteResponseData({
          actor: identity,
          aiAuthoring,
          allowActiveLease: true,
          missingOk: true,
          removeObject,
          responseId,
          revokeOwnerSessions: false,
        });
        return { discarded: true };
      } catch (error) {
        try {
          await createResponseDeletionAudit({
            actorId: identity.id,
            error,
            outcome: AuditOutcome.failure,
            targetId: responseId,
          });
        } catch {
          // Preserve the deletion error if the failure audit cannot be persisted.
        }
        throw error;
      }
    })
    .post(
      "/api/forms/:publicId/draft",
      async ({ request, params, set }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        const form = await findFormByPublicId(params.publicId);
        const pictureFields =
          form.fillMethod === FillMethod.native && form.publishedTemplate
            ? await nativePictureManifestFields(form.publishedTemplate.id)
            : [];
        const nativeRequest = await nativeResponseRequest(
          request,
          pictureFields
        );
        const input = nativeRequest.input;
        requireCurrentFillMethod(form, input);
        const responseId = requiredString(input, "responseId");
        const documentKey = requiredString(input, "documentKey");
        const response = await findOwnedResponse(
          responseId,
          form.id,
          identity.id
        );
        if (
          response.status !== ResponseStatus.draft ||
          response.draftDocumentKey !== documentKey ||
          response.publishedVersion !== form.version
        ) {
          fail(409, "stale_response", "The response is no longer editable");
        }
        const capabilityScope = {
          documentKey,
          formId: form.id,
          targetId: response.id,
          targetType: "response",
        } as const;
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "save-draft",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const data = await normalizeResponseData(
          form,
          response,
          input.data,
          false,
          form.fillMethod === FillMethod.onlyoffice,
          form.fillMethod === FillMethod.onlyoffice
            ? input.canonicalDateFields
            : undefined
        );
        if (await activeOperationForResponse(response.id)) {
          fail(
            409,
            "operation_in_progress",
            "Another response operation is already in progress"
          );
        }
        const nativeDocumentBytes =
          pictureFields.length === 0
            ? null
            : response.draftObjectKey
              ? await nativeResponseDocument(
                  response.publishedTemplateId,
                  data,
                  await readObject(response.draftObjectKey),
                  nativeRequest.pictures
                )
              : fail(
                  409,
                  "document_unavailable",
                  "Response document is unavailable"
                );
        if (nativeDocumentBytes) {
          await validateResponseDocument(
            response.publishedTemplateId,
            nativeDocumentBytes,
            false
          );
        }
        const operationId = crypto.randomUUID();
        const stagedObjectKey = objectKey(
          "operations",
          operationId,
          "draft",
          crypto.randomUUID(),
          "docx"
        );
        const metadata: OperationMetadata = {
          action: "save-draft",
          data,
          finalObjectKey: objectKey(
            "responses",
            response.id,
            "draft",
            crypto.randomUUID(),
            "docx"
          ),
          formId: form.id,
          nextDocumentKey: `response-${response.id}-${crypto.randomUUID()}`,
          publicId: form.publicId,
          responseId: response.id,
          ...(form.fillMethod === FillMethod.onlyoffice
            ? callbackFieldMetadata(data, input.data, response.prefillSnapshot)
            : {}),
          ...(nativeDocumentBytes ? { nativeDocumentStaged: true } : {}),
          stagedObjectKey,
        };

        const operation = await createOperation({
          actorId: identity.id,
          authorization,
          capabilityScope,
          documentKey,
          expectedFillMethod: form.fillMethod,
          formId: form.id,
          metadata,
          ownerUserId: identity.id,
          responseId: response.id,
          stagingObjectKey: stagedObjectKey,
          targetId: response.id,
          targetType: OperationTargetType.response,
          type: operationTypeForAction["save-draft"],
        });
        if (nativeDocumentBytes) {
          try {
            await storeObject(
              stagedObjectKey,
              nativeDocumentBytes,
              DOCX_CONTENT_TYPE
            );
          } catch (error) {
            await updateOperationFailed(operation.id, "document_save_failed");
            throw error;
          }
        }
        set.status = 202;
        if (form.fillMethod === FillMethod.native) {
          void processNativeResponseOperation(operation, storeObject);
        } else {
          launchForceSave(operation, onlyOffice, allowedCallbackOrigins);
        }
        return {
          operationCapability: operationEditorCapability(
            identity,
            capabilityScope,
            operation.id
          ),
          operationId: operation.id,
          responseId: response.id,
          status: operation.status,
        };
      },
      { parse: "none" }
    )
    .post(
      "/api/forms/:publicId/submit",
      async ({ request, params, set }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        const form = await findFormByPublicId(params.publicId);
        const pictureFields =
          form.fillMethod === FillMethod.native && form.publishedTemplate
            ? await nativePictureManifestFields(form.publishedTemplate.id)
            : [];
        const nativeRequest = await nativeResponseRequest(
          request,
          pictureFields
        );
        const input = nativeRequest.input;
        requireCurrentFillMethod(form, input);
        const responseId = requiredString(input, "responseId");
        const documentKey = requiredString(input, "documentKey");
        const response = await findOwnedResponse(
          responseId,
          form.id,
          identity.id
        );
        if (
          response.status !== ResponseStatus.draft ||
          response.draftDocumentKey !== documentKey ||
          response.publishedVersion !== form.version
        ) {
          fail(409, "stale_response", "The response is no longer editable");
        }
        const capabilityScope = {
          documentKey,
          formId: form.id,
          targetId: response.id,
          targetType: "response",
        } as const;
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "submit",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const data = await normalizeResponseData(
          form,
          response,
          input.data,
          true,
          form.fillMethod === FillMethod.onlyoffice,
          form.fillMethod === FillMethod.onlyoffice
            ? input.canonicalDateFields
            : undefined
        );
        if (await activeOperationForResponse(response.id)) {
          fail(
            409,
            "operation_in_progress",
            "Another response operation is already in progress"
          );
        }
        const nativeDocumentBytes =
          pictureFields.length === 0
            ? null
            : response.draftObjectKey
              ? await nativeResponseDocument(
                  response.publishedTemplateId,
                  data,
                  await readObject(response.draftObjectKey),
                  nativeRequest.pictures
                )
              : fail(
                  409,
                  "document_unavailable",
                  "Response document is unavailable"
                );
        if (nativeDocumentBytes) {
          await validateResponseDocument(
            response.publishedTemplateId,
            nativeDocumentBytes,
            true
          );
        }
        const operationId = crypto.randomUUID();
        const submissionId = crypto.randomUUID();
        const submissionDocumentKey = `submission-${submissionId}-${crypto.randomUUID()}`;
        const stagedObjectKey = objectKey(
          "operations",
          operationId,
          "submission",
          crypto.randomUUID(),
          "docx"
        );
        const metadata: OperationMetadata = {
          action: "submit",
          data,
          finalObjectKey: objectKey(
            "submissions",
            submissionId,
            "filled",
            crypto.randomUUID(),
            "docx"
          ),
          formId: form.id,
          publicId: form.publicId,
          responseId: response.id,
          ...(form.fillMethod === FillMethod.onlyoffice
            ? callbackFieldMetadata(data, input.data, response.prefillSnapshot)
            : {}),
          stagedObjectKey,
          submissionDocumentKey,
          submissionId,
          ...(nativeDocumentBytes ? { nativeDocumentStaged: true } : {}),
        };
        let operation: Operation;
        try {
          operation = await prisma.$transaction(
            async (tx) => {
              const [lockedForm] = await tx.$queryRaw<
                { fillMethod: FillMethod }[]
              >(
                Prisma.sql`
                  SELECT "fill_method" AS "fillMethod"
                  FROM "forms"
                  WHERE "id" = ${form.id}::uuid
                  FOR UPDATE
                `
              );
              if (!lockedForm || lockedForm.fillMethod !== form.fillMethod) {
                fail(
                  409,
                  "fill_method_changed",
                  "The form Fill Method changed"
                );
              }
              await lockActiveEditorLease(tx, authorization, capabilityScope);
              const activeOperation = await tx.operation.findFirst({
                where: {
                  responseId: response.id,
                  status: {
                    in: [OperationStatus.pending, OperationStatus.processing],
                  },
                  targetId: response.id,
                  targetType: OperationTargetType.response,
                },
              });
              if (activeOperation) {
                fail(
                  409,
                  "operation_in_progress",
                  "Another response operation is already in progress"
                );
              }
              const claimed = await tx.response.updateMany({
                data: {
                  status: ResponseStatus.submitting,
                  updatedAt: new Date(),
                },
                where: { id: response.id, status: ResponseStatus.draft },
              });
              if (claimed.count !== 1) {
                fail(
                  409,
                  "operation_in_progress",
                  "Another response operation is already in progress"
                );
              }
              return tx.operation.create({
                data: {
                  actorId: identity.id,
                  documentKey,
                  errorCode: null,
                  formId: form.id,
                  metadata: jsonValue(metadata),
                  ownerUserId: identity.id,
                  responseId: response.id,
                  stagingObjectKey: stagedObjectKey,
                  status: OperationStatus.pending,
                  submissionId: null,
                  targetId: response.id,
                  targetType: OperationTargetType.response,
                  type: operationTypeForAction.submit,
                },
              });
            },
            { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
          );
        } catch (error) {
          if (databaseErrorCode(error) === "P2034") {
            fail(
              409,
              "operation_in_progress",
              "Another response operation is already in progress"
            );
          }
          throw error;
        }
        if (nativeDocumentBytes) {
          try {
            await storeObject(
              stagedObjectKey,
              nativeDocumentBytes,
              DOCX_CONTENT_TYPE
            );
          } catch (error) {
            await updateOperationFailed(operation.id, "document_save_failed");
            throw error;
          }
        }
        set.status = 202;
        if (form.fillMethod === FillMethod.native) {
          void processNativeResponseOperation(operation, storeObject);
        } else {
          launchForceSave(operation, onlyOffice, allowedCallbackOrigins);
        }
        return {
          operationCapability: operationEditorCapability(
            identity,
            capabilityScope,
            operation.id
          ),
          operationId: operation.id,
          responseId: response.id,
          status: operation.status,
          submissionId,
        };
      },
      { parse: "none" }
    )
    .get("/api/operations/:id", async ({ request, params }) => {
      const authorization = await editorAuthorization(request);
      const { actor: identity } = authorization;
      validateId(params.id, "Operation");
      let operation = await prisma.operation.findUnique({
        where: { id: params.id },
      });
      if (!operation) {
        fail(404, "not_found", "Operation was not found");
      }
      if (authorization.capability) {
        const targetType =
          operation.targetType === OperationTargetType.template_draft
            ? "template-draft"
            : operation.targetType === OperationTargetType.response
              ? "response"
              : operation.targetType === OperationTargetType.correction
                ? "correction"
                : null;
        if (
          !targetType ||
          !operation.documentKey ||
          operation.actorId !== identity.id
        ) {
          fail(
            403,
            "editor_capability_scope_mismatch",
            "Editor capability does not permit this operation"
          );
        }
        requireEditorScope(authorization, {
          action: "poll-operation",
          documentKey: operation.documentKey,
          formId: operation.formId,
          operationId: operation.id,
          targetId: operation.targetId,
          targetType,
        });
      } else if (identity.role !== "admin") {
        if (!operation.responseId) {
          fail(403, "forbidden", "You may not access this operation");
        }
        const response = await prisma.response.findUnique({
          select: { userId: true },
          where: { id: operation.responseId },
        });
        if (response?.userId !== identity.id) {
          fail(403, "forbidden", "You may not access this operation");
        }
      }
      operation = await expireOperationIfNeeded(operation);
      await cleanupTerminalOperationObjects(operation);
      return {
        operation: {
          createdAt: operation.createdAt,
          error: operation.errorCode,
          id: operation.id,
          responseId: operation.responseId,
          result:
            operation.status === OperationStatus.completed
              ? operation.result
              : undefined,
          status: operation.status,
          submissionId: operation.submissionId,
          type: operation.type,
          updatedAt: operation.updatedAt,
        },
      };
    })
    .get("/api/submissions/:id/data", async ({ request, params, query }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Submission");
      const submission = await findSubmissionWithRevisions(params.id, true);
      if (!submission) {
        fail(404, "not_found", "Submission was not found");
      }
      canReadSubmission(identity, submission);
      const selected = selectedSubmissionRevision(
        submission,
        revisionSelector(query.revision)
      );
      await createResponseAudit({
        action: selected.correction ? "view_correction" : "view_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: {
          revision: selected.revision,
          state: "submitted",
        },
        targetId: selected.correction?.id ?? submission.responseId,
        targetType: selected.correction ? "correction" : "submission",
      });
      const manifestFields = await receiptManifestFields(submission);
      return {
        correction: selected.correction
          ? {
              createdAt: selected.correction.createdAt,
              reason: selected.correction.reason,
              revision: selected.correction.revision,
            }
          : null,
        data: selected.data,
        fields: manifestFields.map(receiptField),
        returnUrl: prefillReturnUrl,
        revision: selected.revision,
        submission: submissionSummary(submission, {
          formPublicId: submission.form.publicId,
          formTitle: submission.form.title,
          userEmail:
            identity.role === "admin" ? submission.owner.email : undefined,
        }),
      };
    })
    .get("/api/submissions/:id/json", async ({ request, params, query }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Submission");
      const submission = await findSubmissionWithRevisions(params.id);
      if (!submission) {
        fail(404, "not_found", "Submission was not found");
      }
      canReadSubmission(identity, submission);
      const selected = selectedSubmissionRevision(
        submission,
        revisionSelector(query.revision)
      );
      await createResponseAudit({
        action: selected.correction ? "export_correction" : "export_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: {
          format: "json",
          revision: selected.revision,
          state: "submitted",
        },
        targetId: selected.correction?.id ?? submission.id,
        targetType: selected.correction ? "correction" : "submission",
      });
      const suffix =
        selected.revision === 0 ? "" : `-revision-${selected.revision}`;
      return Response.json(selected.data, {
        headers: {
          "Content-Disposition": `attachment; filename="submission-${submission.id}${suffix}.json"`,
          "Content-Type": "application/json; charset=utf-8",
        },
      });
    })
    .get("/api/submissions/:id/docx", async ({ request, params, query }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Submission");
      const submission = await findSubmissionWithRevisions(params.id);
      if (!submission) {
        fail(404, "not_found", "Submission was not found");
      }
      canReadSubmission(identity, submission);
      const selected = selectedSubmissionRevision(
        submission,
        revisionSelector(query.revision)
      );
      if (!(await objectExists(selected.objectKey))) {
        fail(404, "not_found", "Submission document was not found");
      }
      await createResponseAudit({
        action: selected.correction ? "export_correction" : "export_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: {
          format: "docx",
          revision: selected.revision,
          state: "submitted",
        },
        targetId: selected.correction?.id ?? submission.id,
        targetType: selected.correction ? "correction" : "submission",
      });
      const suffix =
        selected.revision === 0 ? "" : `-revision-${selected.revision}`;
      return new Response(streamObject(selected.objectKey), {
        headers: {
          "Content-Disposition": `attachment; filename="submission-${submission.id}${suffix}.docx"`,
          "Content-Type": DOCX_CONTENT_TYPE,
        },
      });
    })
    .get(
      "/api/submissions/:id/pdf",
      async ({ request, params, query, set }) => {
        const identity = await requireIdentity(request);
        validateId(params.id, "Submission");
        const submission = await findSubmissionWithRevisions(params.id);
        if (!submission) {
          fail(404, "not_found", "Submission was not found");
        }
        canReadSubmission(identity, submission);
        const selected = selectedSubmissionRevision(
          submission,
          revisionSelector(query.revision)
        );
        let pdf: Uint8Array;
        try {
          pdf = await onlyOffice.convertDocxToPdf(selected.documentKey, true);
        } catch (error) {
          await createResponseAudit({
            action: selected.correction
              ? "export_correction"
              : "export_response",
            actorId: identity.id,
            outcome: AuditOutcome.failure,
            safeMetadata: {
              errorCode: "pdf_conversion_failed",
              format: "pdf",
              revision: selected.revision,
              state: "submitted",
            },
            targetId: selected.correction?.id ?? submission.id,
            targetType: selected.correction ? "correction" : "submission",
          });
          throw error;
        }
        await createResponseAudit({
          action: selected.correction ? "export_correction" : "export_response",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {
            format: "pdf",
            revision: selected.revision,
            state: "submitted",
          },
          targetId: selected.correction?.id ?? submission.id,
          targetType: selected.correction ? "correction" : "submission",
        });
        const suffix =
          selected.revision === 0 ? "" : `-revision-${selected.revision}`;
        set.headers["Content-Type"] = "application/pdf";
        set.headers["Content-Disposition"] =
          `attachment; filename="submission-${submission.id}${suffix}.pdf"`;
        return pdf;
      }
    )
    .get("/onlyoffice/document/:key", async ({ request, params, query }) => {
      const { key } = params;
      const token = typeof query.token === "string" ? query.token : "";
      if (
        !verifyDocumentAccessToken(token, key) ||
        !verifyOnlyOfficeAuthorization(request.headers.get("authorization"), {
          url: request.url,
        })
      ) {
        fail(
          401,
          "unauthorized",
          "Document access token is invalid or expired"
        );
      }
      const objectKey = await operationDocumentKey(key);
      if (!objectKey || !(await objectExists(objectKey))) {
        fail(404, "not_found", "Document was not found");
      }
      if (query.pdf === "1") {
        return new Response(await nativePdfDocument(key), {
          headers: { "Content-Type": DOCX_CONTENT_TYPE },
        });
      }
      return new Response(streamObject(objectKey), {
        headers: { "Content-Type": DOCX_CONTENT_TYPE },
      });
    })
    .get("/onlyoffice-plugin/config.json", ({ request, set }) => {
      const origin = request.headers.get("origin");
      if (origin && !pluginOrigins.has(origin)) {
        fail(403, "forbidden_origin", "Origin is not allowed");
      }
      if (origin && origin !== corsOrigin) {
        set.headers["Access-Control-Allow-Origin"] = origin;
        set.headers.Vary = "Origin";
      }
      return {
        guid: pluginGuid,
        name: "ตั้งค่าฟิลด์",
        variations: [
          {
            EditorsSupport: ["word"],
            buttons: [],
            description: "แผงตั้งค่าฟิลด์สำหรับแบบฟอร์ม",
            events: [
              "onToolbarMenuClick",
              "onDocumentContentReady",
              "onChangeContentControl",
              "onTargetPositionChanged",
            ],
            initData: "",
            initDataType: "none",
            initOnSelectionChanged: true,
            isActivated: true,
            isInsideMode: false,
            isModal: false,
            isViewer: false,
            isVisual: true,
            type: "panelRight",
            url: "index.html",
          },
        ],
        version: "2.1.0",
      };
    })
    .get("/onlyoffice-plugin", pluginIndexResponse)
    .get("/onlyoffice-plugin/", pluginIndexResponse)
    .get("/onlyoffice-plugin/index.html", pluginIndexResponse)
    .get("/onlyoffice-plugin/plugin.js", () =>
      Bun.file(path.resolve(pluginDir, "plugin.js"))
    )
    .post(
      "/onlyoffice/callback",
      async ({ request }) => {
        const payload = await readJsonRecord(request, maxCallbackBodyBytes);
        if (
          !verifyOnlyOfficeAuthorization(
            request.headers.get("authorization"),
            payload
          )
        ) {
          fail(401, "invalid_onlyoffice_token", "OnlyOffice token is invalid");
        }
        const status =
          typeof payload.status === "number"
            ? payload.status
            : Number(payload.status);
        if (status !== 6 && status !== 7) {
          return { error: 0 };
        }
        if (typeof payload.userdata !== "string") {
          return { error: 1 };
        }
        const claim = callbackClaim(payload.userdata);
        if (!claim) {
          return { error: 1 };
        }
        const operation = await prisma.operation.findUnique({
          where: { id: claim.operationId },
        });
        if (
          !operation ||
          claim.documentKey !== operation.documentKey ||
          claim.operationType !== operation.type ||
          typeof payload.key !== "string" ||
          payload.key !== operation.documentKey
        ) {
          return { error: 1 };
        }
        const consumption = await consumeCallbackClaim(
          operation.id,
          payload.userdata
        );
        if (consumption === "invalid") {
          return { error: 1 };
        }
        if (
          consumption === "replayed" ||
          operation.status === OperationStatus.completed ||
          operation.status === OperationStatus.failed
        ) {
          if (
            operation.status === OperationStatus.completed ||
            operation.status === OperationStatus.failed
          ) {
            await cleanupTerminalOperationObjects(operation);
          }
          return { error: 0 };
        }
        if (status === 7) {
          await updateOperationFailed(
            operation.id,
            "onlyoffice_document_error"
          );
          return { error: 0 };
        }
        try {
          await finalizeCallback(
            operation.id,
            payload as unknown as CallbackPayload,
            undefined,
            allowedCallbackOrigins,
            callbackMaximumBytes
          );
          return { error: 0 };
        } catch (error) {
          await updateOperationFailed(
            operation.id,
            error instanceof HttpError && error.code === "invalid_template"
              ? "invalid_template"
              : "callback_processing_failed"
          );
          return { error: 1 };
        }
      },
      { parse: "none" }
    )
    .onStop(() => aiAuthoring.close());
}

async function readTemplateSourceBytes(
  absolutePath: string
): Promise<Uint8Array> {
  const file = Bun.file(absolutePath);
  if (!(await file.exists())) {
    fail(404, "not_found", "Template source was not found");
  }
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes.byteLength > maxTemplateUploadBytes) {
    fail(413, "payload_too_large", "Template source is too large");
  }
  validateTemplatePackage(bytes);
  return bytes;
}
