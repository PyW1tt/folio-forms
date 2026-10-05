import { prisma, Prisma, ResponseStatus, FormStatus } from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { requireIdentity } from "../auth/identity";
import { databaseErrorCode } from "../db-errors";
import { findFormByPublicId } from "../forms/query";
import { fail } from "../http/errors";
import { jsonRecord } from "../http/input";
import {
  deleteObjects,
  deleteObjectUnlessCanonical,
} from "../operations/object-cleanup";
import {
  pendingClaimCookie,
  pendingClaimFor,
  handoffUnavailable,
  prefillRequired,
} from "../prefill/input";
import {
  consumeSubmittedPrefillHandoff,
  redeemPrefillHandoff,
} from "../prefill/redemption";
import { editableFieldsForSnapshot } from "../prefill/values";
import { jsonValue } from "../responses/data";
import { userEditorConfig } from "../responses/editor";
import { responseSummary } from "../responses/query";
import {
  putObject,
  DOCX_CONTENT_TYPE,
  objectExists,
  readObject,
  objectKey,
} from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerPublicFormRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "handoffClock">
): void {
  const { handoffClock } = dependencies;
  app
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
    });
}
