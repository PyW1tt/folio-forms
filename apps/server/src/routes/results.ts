import type { Prisma } from "@onlyoffice/db";
import {
  prisma,
  AuditOutcome,
  ResponseStatus,
  FillMethod,
  OperationTargetType,
} from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { createResponseAudit } from "../audit/events";
import {
  requireIdentity,
  requireAdmin,
  normalizeEmail,
} from "../auth/identity";
import { callbackFieldMetadata } from "../documents/office-values";
import {
  requireActionEditorAuthorization,
  requireEditorScope,
  operationEditorCapability,
} from "../editor/authorization";
import { requireActiveEditorLease } from "../editor/leases";
import {
  manifestFieldsWithDocumentMetadata,
  receiptField,
} from "../forms/query";
import { fail } from "../http/errors";
import {
  validateId,
  queryString,
  publicIdPattern,
  jsonRecord,
  correctionInput,
} from "../http/input";
import type { JsonRecord } from "../model-types";
import { readOnlyViewerConfig } from "../onlyoffice";
import { createOperation } from "../operations/create";
import { launchForceSave } from "../operations/force-save";
import { activeOperationForResponse } from "../operations/lifecycle";
import type { OperationMetadata } from "../operations/model";
import { operationTypeForAction } from "../operations/model";
import { normalizeResponseData } from "../responses/data";
import { correctionEditorConfig } from "../responses/editor";
import {
  adminResultDate,
  adminResultCursor,
  adminResultPageSize,
  adminResultCursorValue,
  adminResultSummary,
  revisionSelector,
  selectedSubmissionRevision,
} from "../responses/query";
import { objectExists, objectKey } from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerResultRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "onlyOffice" | "allowedCallbackOrigins">
): void {
  const { onlyOffice, allowedCallbackOrigins } = dependencies;
  app
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
    );
}
