import { prisma, OperationStatus } from "@onlyoffice/db";
import { env } from "@onlyoffice/env/server";

import {
  responseTextControlValues,
  validateTemplateControls,
} from "../documents/fields";
import { overlayNativeResponseDocument } from "../documents/native-document";
import {
  onlyOfficeFieldDisplayMetadata,
  normalizeOnlyOfficeDisplayValue,
} from "../documents/office-values";
import {
  validateTemplatePackage,
  validateOfficeRelationshipsBytes,
} from "../documents/package";
import { supportsNativeTemplate } from "../forms/query";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { maxTemplateUploadBytes } from "../http/input";
import { originOf } from "../http/origins";
import { createOnlyOfficeAuthorization } from "../onlyoffice";
import { manifestFieldValueMatches } from "../responses/data";
import { readObject, putObject, DOCX_CONTENT_TYPE } from "../storage";
import { updateOperationFailed } from "./lifecycle";
import type { CallbackPayload, OperationCompletion } from "./model";
import { operationMetadata } from "./model";
import {
  cleanupTerminalOperationObjects,
  deleteObjects,
  deleteObjectUnlessCanonical,
} from "./object-cleanup";
import {
  completeDraftOperation,
  completeCorrectionOperation,
  completeSubmitOperation,
} from "./response-completion";
import {
  completeTemplateOperation,
  completePublishOperation,
} from "./template-completion";

export const maxCallbackDocumentBytes = maxTemplateUploadBytes;
export const maxCallbackBodyBytes = 64 * 1024;
const callbackInternalOrigin = originOf(env.ONLYOFFICE_INTERNAL_URL);
const callbackPublicOrigin = originOf(env.ONLYOFFICE_URL);
const callbackDocumentOrigin = originOf(env.ONLYOFFICE_DOCUMENT_BASE_URL);
export const callbackOrigins = new Set(
  [callbackInternalOrigin, callbackPublicOrigin, callbackDocumentOrigin].filter(
    (origin): origin is string => Boolean(origin)
  )
);
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

export async function finalizeCallback(
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
