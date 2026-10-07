// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { OperationType } from "@onlyoffice/db";

import { fail } from "../http/errors";
import { asRecord } from "../http/input";
import type { JsonRecord } from "../model-types";

type OperationAction =
  | "save-template"
  | "publish"
  | "save-draft"
  | "submit"
  | "save-correction";
export const operationTypeForAction: Record<OperationAction, OperationType> = {
  publish: "publish_form",
  "save-correction": "save_correction",
  "save-draft": "save_draft",
  "save-template": "save_template_draft",
  submit: "submit_response",
};
export type OperationErrorCode =
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
export interface OperationMetadata extends JsonRecord {
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

export interface OperationCompletion {
  cleanupObjectKeys: string[];
}
export interface CallbackPayload {
  key?: unknown;
  status?: unknown;
  url?: unknown;
  userdata?: unknown;
}

export function operationMetadata(value: unknown): OperationMetadata {
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
