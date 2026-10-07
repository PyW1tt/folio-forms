import type { FormStatus, FillMethod, Prisma } from "@onlyoffice/db";
import { AuditOutcome, prisma } from "@onlyoffice/db";

import { databaseErrorCode } from "../db-errors";
import { HttpError } from "../http/errors";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { FormSource } from "../http/input";
import { idPattern } from "../http/input";
import type { JsonRecord } from "../model-types";
import type { OperationErrorCode } from "../operations/model";
import { jsonValue } from "../responses/data";

export type AccountAuditAction =
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
export type FormAuditAction =
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

export function safeAuditMetadata(value: unknown): JsonRecord {
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
export function accountAuditTargetId(value: string): string | null {
  return idPattern.test(value) ? value : null;
}

interface AccountAuditMetadata {
  change?: string;
  errorCode?: string;
}

export async function createAccountAudit(
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

export async function createAccountFailureAudit({
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

export function formAuditErrorCodeFromCode(code: string): FormAuditErrorCode {
  return formAuditErrorCodes[code]
    ? (code as FormAuditErrorCode)
    : "internal_error";
}

function formAuditErrorCode(error: unknown): FormAuditErrorCode {
  const code = error instanceof HttpError ? error.code : "internal_error";
  return formAuditErrorCodeFromCode(code);
}

export async function createFormAudit(
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

export async function createFormFailureAudit({
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

export async function createResponseAudit({
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

export async function createResponseDeletionAudit({
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
