import {
  prisma,
  AuditOutcome,
  ResponseStatus,
  FieldType,
} from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import {
  createResponseAudit,
  createResponseDeletionAudit,
} from "../audit/events";
import { requireIdentity, requireAdmin } from "../auth/identity";
import { responsePicturePresence } from "../documents/pictures";
import { fail } from "../http/errors";
import {
  validateId,
  readJsonRecord,
  accountBodyMaximumBytes,
  jsonRecord,
} from "../http/input";
import { activeOperationForResponse } from "../operations/lifecycle";
import { deleteResponseData } from "../responses/deletion";
import { responseSummary, correctionRevisionSummary } from "../responses/query";
import type { CorrectionRevisionSummary } from "../responses/query";
import {
  DOCX_CONTENT_TYPE,
  objectExists,
  readObject,
  streamObject,
} from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerResponseRoutes(
  app: Elysia,
  dependencies: Pick<
    RouteDependencies,
    "onlyOffice" | "removeObject" | "aiAuthoring"
  >
): void {
  const { onlyOffice, removeObject, aiAuthoring } = dependencies;
  app
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
    });
}
