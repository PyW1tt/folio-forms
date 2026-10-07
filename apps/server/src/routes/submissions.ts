import { AuditOutcome } from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { createResponseAudit } from "../audit/events";
import { requireIdentity } from "../auth/identity";
import { receiptField } from "../forms/query";
import { fail } from "../http/errors";
import { validateId } from "../http/input";
import {
  submissionSummary,
  revisionSelector,
  selectedSubmissionRevision,
  findSubmissionWithRevisions,
  canReadSubmission,
  receiptManifestFields,
} from "../responses/query";
import { DOCX_CONTENT_TYPE, objectExists, streamObject } from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerSubmissionRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "onlyOffice" | "prefillReturnUrl">
): void {
  const { onlyOffice, prefillReturnUrl } = dependencies;
  app
    .get("/api/submissions/:id/data", async ({ request, params, query }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Submission");
      const submission = await findSubmissionWithRevisions(params.id, true);
      if (!submission) {
        fail(404, "not_found", "Submission was not found");
      }
      canReadSubmission(identity, submission);
      const selected = selectedSubmissionRevision(
        submission,
        revisionSelector(query.revision)
      );
      await createResponseAudit({
        action: selected.correction ? "view_correction" : "view_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: {
          revision: selected.revision,
          state: "submitted",
        },
        targetId: selected.correction?.id ?? submission.responseId,
        targetType: selected.correction ? "correction" : "submission",
      });
      const manifestFields = await receiptManifestFields(submission);
      return {
        correction: selected.correction
          ? {
              createdAt: selected.correction.createdAt,
              reason: selected.correction.reason,
              revision: selected.correction.revision,
            }
          : null,
        data: selected.data,
        fields: manifestFields.map(receiptField),
        returnUrl: prefillReturnUrl,
        revision: selected.revision,
        submission: submissionSummary(submission, {
          formPublicId: submission.form.publicId,
          formTitle: submission.form.title,
          userEmail:
            identity.role === "admin" ? submission.owner.email : undefined,
        }),
      };
    })
    .get("/api/submissions/:id/json", async ({ request, params, query }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Submission");
      const submission = await findSubmissionWithRevisions(params.id);
      if (!submission) {
        fail(404, "not_found", "Submission was not found");
      }
      canReadSubmission(identity, submission);
      const selected = selectedSubmissionRevision(
        submission,
        revisionSelector(query.revision)
      );
      await createResponseAudit({
        action: selected.correction ? "export_correction" : "export_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: {
          format: "json",
          revision: selected.revision,
          state: "submitted",
        },
        targetId: selected.correction?.id ?? submission.id,
        targetType: selected.correction ? "correction" : "submission",
      });
      const suffix =
        selected.revision === 0 ? "" : `-revision-${selected.revision}`;
      return Response.json(selected.data, {
        headers: {
          "Content-Disposition": `attachment; filename="submission-${submission.id}${suffix}.json"`,
          "Content-Type": "application/json; charset=utf-8",
        },
      });
    })
    .get("/api/submissions/:id/docx", async ({ request, params, query }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Submission");
      const submission = await findSubmissionWithRevisions(params.id);
      if (!submission) {
        fail(404, "not_found", "Submission was not found");
      }
      canReadSubmission(identity, submission);
      const selected = selectedSubmissionRevision(
        submission,
        revisionSelector(query.revision)
      );
      if (!(await objectExists(selected.objectKey))) {
        fail(404, "not_found", "Submission document was not found");
      }
      await createResponseAudit({
        action: selected.correction ? "export_correction" : "export_response",
        actorId: identity.id,
        outcome: AuditOutcome.success,
        safeMetadata: {
          format: "docx",
          revision: selected.revision,
          state: "submitted",
        },
        targetId: selected.correction?.id ?? submission.id,
        targetType: selected.correction ? "correction" : "submission",
      });
      const suffix =
        selected.revision === 0 ? "" : `-revision-${selected.revision}`;
      return new Response(streamObject(selected.objectKey), {
        headers: {
          "Content-Disposition": `attachment; filename="submission-${submission.id}${suffix}.docx"`,
          "Content-Type": DOCX_CONTENT_TYPE,
        },
      });
    })
    .get(
      "/api/submissions/:id/pdf",
      async ({ request, params, query, set }) => {
        const identity = await requireIdentity(request);
        validateId(params.id, "Submission");
        const submission = await findSubmissionWithRevisions(params.id);
        if (!submission) {
          fail(404, "not_found", "Submission was not found");
        }
        canReadSubmission(identity, submission);
        const selected = selectedSubmissionRevision(
          submission,
          revisionSelector(query.revision)
        );
        let pdf: Uint8Array;
        try {
          pdf = await onlyOffice.convertDocxToPdf(selected.documentKey, true);
        } catch (error) {
          await createResponseAudit({
            action: selected.correction
              ? "export_correction"
              : "export_response",
            actorId: identity.id,
            outcome: AuditOutcome.failure,
            safeMetadata: {
              errorCode: "pdf_conversion_failed",
              format: "pdf",
              revision: selected.revision,
              state: "submitted",
            },
            targetId: selected.correction?.id ?? submission.id,
            targetType: selected.correction ? "correction" : "submission",
          });
          throw error;
        }
        await createResponseAudit({
          action: selected.correction ? "export_correction" : "export_response",
          actorId: identity.id,
          outcome: AuditOutcome.success,
          safeMetadata: {
            format: "pdf",
            revision: selected.revision,
            state: "submitted",
          },
          targetId: selected.correction?.id ?? submission.id,
          targetType: selected.correction ? "correction" : "submission",
        });
        const suffix =
          selected.revision === 0 ? "" : `-revision-${selected.revision}`;
        set.headers["Content-Type"] = "application/pdf";
        set.headers["Content-Disposition"] =
          `attachment; filename="submission-${submission.id}${suffix}.pdf"`;
        return pdf;
      }
    );
}
