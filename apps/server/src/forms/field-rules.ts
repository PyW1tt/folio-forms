// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { PrefillPolicy, FormStatus } from "@onlyoffice/db";

import { requireAdmin } from "../auth/identity";
import type {
  ActionEditorAuthorization,
  EditorCapabilityScope,
} from "../editor/authorization";
import {
  requireActionEditorAuthorization,
  requireEditorScope,
} from "../editor/authorization";
import { requireActiveEditorLease } from "../editor/leases";
import { fail } from "../http/errors";
import type {
  JsonRecord,
  FormWithDocuments,
  TemplateDraft,
} from "../model-types";
import { externalSchemaItems } from "./field-schema";
import { findFormByPublicId } from "./query";

export const fieldRuleBodyMaximumBytes = 8 * 1024;
const fieldTagMaximumLength = 512;
const fieldPointerMaximumLength = 2048;

type FieldRulePolicy = "editable" | "lock-when-available";
interface FieldRuleInput {
  documentKey: string;
  previousTag: string | null;
  prefillPointer: string | null;
  prefillPolicy: FieldRulePolicy;
  required: boolean;
  tag: string;
}

export function fieldRuleDto(rule: {
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

export function fieldRuleInput(input: JsonRecord): FieldRuleInput {
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

export function fieldRulePrefillPolicy(policy: FieldRulePolicy): PrefillPolicy {
  return policy === "lock-when-available"
    ? PrefillPolicy.lock_when_available
    : PrefillPolicy.editable;
}

export async function requireFieldRuleContext(
  request: Request,
  publicId: string
): Promise<{
  authorization: ActionEditorAuthorization;
  capabilityScope: Omit<EditorCapabilityScope, "action" | "operationId">;
  form: FormWithDocuments;
  templateDraft: TemplateDraft;
}> {
  const authorization = await requireActionEditorAuthorization(request);
  requireAdmin(authorization.actor);
  const form = await findFormByPublicId(publicId);
  if (form.status === FormStatus.published || form.publishedTemplate) {
    fail(
      409,
      "published_immutable",
      "Published forms cannot change Field rules"
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
    action: "configure-fields",
  });
  await requireActiveEditorLease(authorization, capabilityScope);
  return { authorization, capabilityScope, form, templateDraft };
}

export function validateFieldRulePointer(prefillPointer: string | null): void {
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
