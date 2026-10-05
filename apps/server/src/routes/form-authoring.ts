import {
  prisma,
  Prisma,
  AuditOutcome,
  FormStatus,
  OperationTargetType,
} from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { createFormFailureAudit, createFormAudit } from "../audit/events";
import { requireIdentity, requireAdmin } from "../auth/identity";
import { databaseErrorCode } from "../db-errors";
import {
  actionEditorCapability,
  requireActionEditorAuthorization,
  requireEditorScope,
  operationEditorCapability,
} from "../editor/authorization";
import {
  claimEditorLease,
  editorLeaseBridge,
  lockActiveEditorLease,
  requireActiveEditorLease,
} from "../editor/leases";
import {
  requireFieldRuleContext,
  fieldRuleDto,
  fieldRuleInput,
  fieldRuleBodyMaximumBytes,
  validateFieldRulePointer,
  fieldRulePrefillPolicy,
} from "../forms/field-rules";
import { schemaPage } from "../forms/field-schema";
import { findFormByPublicId } from "../forms/query";
import { fail } from "../http/errors";
import {
  readJsonRecord,
  publicIdPattern,
  readDocumentKeyInput,
} from "../http/input";
import type { JsonRecord, DraftFieldRule } from "../model-types";
import { editorConfig } from "../onlyoffice";
import { createOperation } from "../operations/create";
import { launchForceSave } from "../operations/force-save";
import { activeOperationForForm } from "../operations/lifecycle";
import type { OperationMetadata } from "../operations/model";
import { operationTypeForAction } from "../operations/model";
import { submissionSummary } from "../responses/query";
import { objectExists, objectKey } from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerFormAuthoringRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "onlyOffice" | "allowedCallbackOrigins">
): void {
  const { onlyOffice, allowedCallbackOrigins } = dependencies;
  app
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
        await requireFieldRuleContext(request, params.publicId);
        const queryRecord = query as unknown as JsonRecord;
        return schemaPage(queryRecord.q, queryRecord.cursor);
      }
    )
    .get(
      "/api/admin/forms/:publicId/field-rules",
      async ({ request, params }) => {
        const { templateDraft } = await requireFieldRuleContext(
          request,
          params.publicId
        );
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
        const { authorization, capabilityScope, form, templateDraft } =
          await requireFieldRuleContext(request, params.publicId);
        const { actor: identity } = authorization;
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
    );
}
