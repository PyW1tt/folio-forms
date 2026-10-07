import { ResponseStatus, prisma } from "@onlyoffice/db";

import type { Identity, Actor } from "../auth/identity";
import { manifestFieldsWithDocumentMetadata } from "../forms/query";
import { fail } from "../http/errors";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { idPattern, validateId, jsonRecord } from "../http/input";
import type {
  Response,
  JsonRecord,
  Submission,
  ResponseWithSnapshot,
  Correction,
  SubmissionWithManifest,
  SubmissionWithRevisions,
} from "../model-types";

export const adminResultPageSize = 25;
interface AdminResultCursor {
  id: string;
  updatedAt: Date;
}

export function adminResultCursor(
  value: string | undefined
): AdminResultCursor | null {
  if (value === undefined) {
    return null;
  }
  try {
    const decoded = JSON.parse(
      Buffer.from(value, "base64url").toString("utf-8")
    ) as { id?: unknown; updatedAt?: unknown };
    if (
      typeof decoded.id !== "string" ||
      !idPattern.test(decoded.id) ||
      typeof decoded.updatedAt !== "string"
    ) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    const updatedAt = new Date(decoded.updatedAt);
    if (!Number.isFinite(updatedAt.getTime())) {
      fail(400, "invalid_request", "cursor is invalid");
    }
    return { id: decoded.id, updatedAt };
  } catch {
    fail(400, "invalid_request", "cursor is invalid");
  }
}

export function adminResultCursorValue(result: {
  id: string;
  updatedAt: Date;
}): string {
  return Buffer.from(
    JSON.stringify({ id: result.id, updatedAt: result.updatedAt.toISOString() })
  ).toString("base64url");
}

export function adminResultDate(
  value: string | undefined,
  key: string
): Date | null {
  if (value === undefined) {
    return null;
  }
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) {
    fail(400, "invalid_request", `${key} is invalid`);
  }
  return date;
}

export function responseSummary(
  response: Response,
  extra: {
    formPublicId?: string;
    formTitle?: string;
    latestCorrectionNumber?: number | null;
    submissionId?: string | null;
    submittedAt?: Date | null;
  } = {}
): JsonRecord {
  return {
    createdAt: response.createdAt,
    formPublicId: extra.formPublicId,
    formTitle: extra.formTitle,
    hasDraft: Boolean(response.draftObjectKey && response.draftData),
    id: response.id,
    latestCorrectionNumber: extra.latestCorrectionNumber,
    publishedVersion: response.publishedVersion,
    status: response.status,
    submissionId: extra.submissionId,
    submittedAt: extra.submittedAt,
    updatedAt: response.updatedAt,
  };
}

export function submissionSummary(
  submission: Submission,
  extra: {
    formPublicId?: string;
    formTitle?: string;
    userEmail?: string;
  } = {}
): JsonRecord {
  return {
    createdAt: submission.createdAt,
    formPublicId: extra.formPublicId,
    formTitle: extra.formTitle,
    id: submission.id,
    responseId: submission.responseId,
    status: "submitted",
    submittedAt: submission.createdAt,
    userEmail: extra.userEmail,
    userId: submission.userId,
  };
}

export function adminResultSummary(response: {
  corrections: { revision: number }[];
  createdAt: Date;
  form: { publicId: string; title: string };
  id: string;
  owner: { email: string };
  status: ResponseStatus;
  submission: { createdAt: Date; id: string } | null;
  updatedAt: Date;
}): JsonRecord {
  const submitted = response.status === ResponseStatus.submitted;
  return {
    createdAt: response.createdAt,
    formPublicId: response.form.publicId,
    formTitle: response.form.title,
    id: response.id,
    latestCorrectionNumber: response.corrections[0]?.revision ?? null,
    state: submitted ? "submitted" : "draft",
    submissionId: response.submission?.id ?? null,
    submittedAt: response.submission?.createdAt ?? null,
    updatedAt: response.updatedAt,
    userEmail: response.owner.email,
  };
}

export async function findOwnedResponse(
  responseId: string,
  formId: string,
  userId: string
): Promise<ResponseWithSnapshot> {
  validateId(responseId, "Response");
  const response = await prisma.response.findFirst({
    include: { prefillSnapshot: true },
    where: { formId, id: responseId, userId },
  });
  if (!response) {
    fail(404, "not_found", "Response was not found");
  }
  return response;
}

export function canReadSubmission(
  identity: Identity,
  submission: Submission
): void {
  if (identity.role === "admin") {
    return;
  }
  if (submission.userId !== identity.id) {
    fail(403, "forbidden", "You may only access your own submission");
  }
}

export interface CorrectionRevisionSummary {
  actorEmail: string | null;
  actorName: string | null;
  createdAt: Date;
  data: JsonRecord;
  documentAvailable: boolean;
  documentKey: string;
  id: string | null;
  reason: string | null;
  revision: number;
  pictures: Record<string, boolean> | null;
}
type RevisionSelector = number | "latest";

export function revisionSelector(value: unknown): RevisionSelector {
  if (value === undefined || value === "latest") {
    return "latest";
  }
  if (value === "original") {
    return 0;
  }
  if (typeof value === "string" && /^(?:0|[1-9]\d*)$/u.test(value)) {
    const revision = Number(value);
    if (Number.isSafeInteger(revision)) {
      return revision;
    }
  }
  fail(
    400,
    "invalid_request",
    "revision must be latest, original, or a revision number"
  );
}

export function selectedSubmissionRevision(
  submission: Submission & { corrections: Correction[] },
  selector: RevisionSelector
): {
  correction: Correction | null;
  data: JsonRecord;
  documentKey: string;
  objectKey: string;
  revision: number;
} {
  const revision =
    selector === "latest"
      ? (submission.corrections[0]?.revision ?? 0)
      : selector;
  if (
    revision > 0 &&
    !submission.corrections.some((item) => item.revision === revision)
  ) {
    fail(404, "not_found", "Submission revision was not found");
  }
  const correction =
    revision === 0
      ? null
      : (submission.corrections.find((item) => item.revision === revision) ??
        null);
  return {
    correction,
    data: jsonRecord(correction?.data ?? submission.data),
    documentKey: correction?.documentKey ?? submission.documentKey,
    objectKey: correction?.objectKey ?? submission.objectKey,
    revision,
  };
}

export function correctionRevisionSummary(
  correction: Correction,
  actor: Pick<Actor, "email" | "name"> | undefined,
  documentAvailable: boolean,
  pictures: Record<string, boolean> | null
): CorrectionRevisionSummary {
  return {
    actorEmail: actor?.email ?? null,
    actorName: actor?.name ?? null,
    createdAt: correction.createdAt,
    data: jsonRecord(correction.data),
    documentAvailable,
    documentKey: correction.documentKey,
    id: correction.id,
    pictures,
    reason: correction.reason,
    revision: correction.revision,
  };
}

export function findSubmissionWithRevisions(
  id: string,
  includeFieldManifest: true
): Promise<SubmissionWithManifest | null>;
export function findSubmissionWithRevisions(
  id: string,
  includeFieldManifest?: false
): Promise<SubmissionWithRevisions | null>;
export function findSubmissionWithRevisions(
  id: string,
  includeFieldManifest = false
): Promise<SubmissionWithManifest | SubmissionWithRevisions | null> {
  return prisma.submission.findUnique({
    include: {
      corrections: { orderBy: { revision: "desc" } },
      form: true,
      owner: true,
      ...(includeFieldManifest
        ? {
            response: {
              include: {
                publishedTemplate: {
                  include: {
                    manifest: {
                      include: {
                        fields: { orderBy: { position: "asc" } },
                      },
                    },
                  },
                },
              },
            },
          }
        : {}),
    },
    where: { id },
  });
}
export async function receiptManifestFields(
  submission: SubmissionWithManifest
) {
  const publishedTemplate = submission.response?.publishedTemplate;
  const manifest = publishedTemplate?.manifest;
  if (!manifest || !publishedTemplate) {
    return [];
  }
  return await manifestFieldsWithDocumentMetadata(
    publishedTemplate.objectKey,
    manifest.displayMetadataVersion,
    manifest.fields
  );
}
