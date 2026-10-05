import { FillMethod, FormStatus } from "@onlyoffice/db";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { JsonRecord } from "../model-types";
import { DOCX_CONTENT_TYPE } from "../storage";
import { fail } from "./errors";

export const idPattern = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/iu;
export const publicIdPattern = /^[0-9a-f]{32}$/u;
export const accountEmailPattern = /^[^\s@]+@[^\s@]+$/iu;
export const maxTemplateUploadBytes = 25 * 1024 * 1024;
export const maxTemplateMultipartOverheadBytes = 64 * 1024;
const maxTemplateMultipartBodyBytes =
  maxTemplateUploadBytes + maxTemplateMultipartOverheadBytes;
const maxAuthoringPdfBytes = 10 * 1024 * 1024;
const maxAuthoringMultipartBodyBytes = 11 * 1024 * 1024;
const correctionReasonMaximumLength = 2000;
export const accountBodyMaximumBytes = 64 * 1024;
export const documentActionBodyMaximumBytes = 8 * 1024;
export const accountEmailMaximumLength = 254;
export type FormSource = "blank" | "upload";

export function asRecord(
  value: unknown,
  message = "Request body must be a JSON object"
): JsonRecord {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail(400, "invalid_request", message);
  }
  return value as JsonRecord;
}
export async function readRequestBytes(
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

export async function readJsonRecord(
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

export async function readAuthoringPromptInput(
  request: Request
): Promise<{ consent: unknown; prompt: string }> {
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

export async function readAuthoringCreationInput(
  request: Request
): Promise<{ consent: unknown; pdfBytes?: Uint8Array; prompt: string }> {
  const contentType = request.headers.get("content-type");
  if (!contentType || !/^multipart\/form-data(?:\s*;|$)/iu.test(contentType)) {
    return await readAuthoringPromptInput(request);
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
      fail(
        400,
        "invalid_request",
        "Only prompt, consent, and pdf are accepted"
      );
    }
    if (entries.has(key)) {
      fail(400, "invalid_request", `${key} must be provided once`);
    }
    entries.set(key, value);
  }
  const prompt = entries.get("prompt");
  const consent = entries.get("consent");
  const pdf = entries.get("pdf");
  if (
    typeof prompt !== "string" ||
    (consent !== undefined && typeof consent !== "string")
  ) {
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

export async function readTemplateCreationInput(
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

export function requiredString(record: JsonRecord, key: string): string {
  const value = record[key];
  if (typeof value !== "string" || value.trim().length === 0) {
    fail(400, "invalid_request", `${key} is required`);
  }
  return value.trim();
}

export async function readDocumentKeyInput(request: Request): Promise<string> {
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

export function correctionInput(body: unknown): CorrectionInput {
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

export async function readFormMetadataInput(
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

export function jsonRecord(
  value: unknown,
  message = "data must be a JSON object"
): JsonRecord {
  return asRecord(value, message);
}

export function validateId(value: string, label: string): string {
  if (!idPattern.test(value)) {
    fail(404, "not_found", `${label} was not found`);
  }
  return value;
}
export function queryString(
  record: JsonRecord,
  key: string
): string | undefined {
  const value = record[key];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    fail(400, "invalid_request", `${key} must be a string`);
  }
  return value;
}
