// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Prisma } from "@onlyoffice/db";

type Form = Prisma.FormGetPayload<Prisma.FormDefaultArgs>;
export type Operation = Prisma.OperationGetPayload<Prisma.OperationDefaultArgs>;
export type PrefillSnapshot =
  Prisma.PrefillSnapshotGetPayload<Prisma.PrefillSnapshotDefaultArgs>;
type PublishedTemplate =
  Prisma.PublishedTemplateGetPayload<Prisma.PublishedTemplateDefaultArgs>;
export type Response = Prisma.ResponseGetPayload<Prisma.ResponseDefaultArgs>;
export type Correction =
  Prisma.CorrectionGetPayload<Prisma.CorrectionDefaultArgs>;
export type Submission =
  Prisma.SubmissionGetPayload<Prisma.SubmissionDefaultArgs>;
export type SubmissionWithManifest = Prisma.SubmissionGetPayload<{
  include: {
    corrections: true;
    form: true;
    owner: true;
    response: {
      include: {
        publishedTemplate: {
          include: {
            manifest: {
              include: { fields: true };
            };
          };
        };
      };
    };
  };
}>;
export type SubmissionWithRevisions = Prisma.SubmissionGetPayload<{
  include: { corrections: true; form: true; owner: true };
}>;
export type TemplateDraft =
  Prisma.TemplateDraftGetPayload<Prisma.TemplateDraftDefaultArgs>;
export type DraftFieldRule =
  Prisma.DraftFieldRuleGetPayload<Prisma.DraftFieldRuleDefaultArgs>;
export type JsonRecord = Record<string, unknown>;

export type FormWithDocuments = Form & {
  templateDraft: TemplateDraft | null;
  publishedTemplate: PublishedTemplate | null;
};

export type ResponseWithSnapshot = Response & {
  prefillSnapshot: PrefillSnapshot | null;
};
