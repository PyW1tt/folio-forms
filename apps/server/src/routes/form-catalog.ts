import {
  prisma,
  Prisma,
  AuditOutcome,
  ResponseStatus,
  FormStatus,
  OperationStatus,
  FillMethod,
  OperationTargetType,
} from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { createFormFailureAudit, createFormAudit } from "../audit/events";
import type { FormAuditAction } from "../audit/events";
import { requireIdentity, requireAdmin } from "../auth/identity";
import { contentHash } from "../digests";
import { validateTemplatePackage } from "../documents/package";
import {
  formDto,
  supportsNativeTemplate,
  findFormByPublicId,
} from "../forms/query";
import { fail } from "../http/errors";
import {
  asRecord,
  readJsonRecord,
  publicIdPattern,
  readFormMetadataInput,
  documentActionBodyMaximumBytes,
  readTemplateCreationInput,
} from "../http/input";
import type { FormSource } from "../http/input";
import { operationTimeoutMs } from "../operations/lifecycle";
import {
  objectCleanupIntentGraceMs,
  drainObjectCleanupIntents,
  uniqueObjectKeys,
} from "../operations/object-cleanup";
import {
  putObject,
  DOCX_CONTENT_TYPE,
  objectExists,
  readObject,
  objectKey,
} from "../storage";
import {
  findTemplateSource,
  readTemplateSourceBytes,
} from "../template-source";
import type { RouteDependencies } from "./dependencies";

export function registerFormCatalogRoutes(
  app: Elysia,
  _dependencies: Pick<RouteDependencies, never>
): void {
  app
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
    });
}
