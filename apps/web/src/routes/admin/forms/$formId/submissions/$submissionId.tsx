// oxlint-disable unicorn(filename-case) -- TanStack Router requires this dynamic route filename.
import { createFileRoute, Link, useParams } from "@tanstack/react-router";
import { ArrowLeft, Download, FileText } from "lucide-react";
import { useEffect, useState } from "react";

import { SubmissionResultViewer } from "@/components/submission-result-viewer";
import { Badge, Card, Notice, Spinner } from "@/components/ui";
import { apiGet, downloadArtifact, formatDate } from "@/lib/api";
import type {
  ReceiptField,
  ResponseRevisionsResponse,
  Submission,
} from "@/lib/api";

const SubmissionDetailRoute = () => {
  const { formId, submissionId } = useParams({
    from: "/admin/forms/$formId/submissions/$submissionId",
  });
  const [submission, setSubmission] = useState<Submission | null>(null);
  const [data, setData] = useState<Record<string, unknown> | null>(null);
  const [fields, setFields] = useState<ReceiptField[]>([]);
  const [pictures, setPictures] = useState<Record<string, boolean> | null>(
    null
  );
  const [documentAvailable, setDocumentAvailable] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    const loadSubmission = async () => {
      try {
        const payload = await apiGet<{
          data: Record<string, unknown>;
          fields: ReceiptField[];
          submission: Submission;
        }>(`/api/submissions/${submissionId}/data?revision=latest`);
        const history = payload.submission.responseId
          ? await apiGet<ResponseRevisionsResponse>(
              `/api/responses/${payload.submission.responseId}/corrections`
            )
          : null;
        if (!cancelled) {
          const latestRevision = history?.revisions.at(-1);
          setSubmission(payload.submission);
          setData(payload.data);
          setFields(payload.fields);
          setPictures(latestRevision?.pictures ?? null);
          setDocumentAvailable(
            Boolean(
              payload.submission.responseId &&
                latestRevision?.document.available
            )
          );
        }
      } catch (caughtError) {
        if (!cancelled) {
          setError(
            caughtError instanceof Error
              ? caughtError.message
              : "Could not load submission."
          );
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    void loadSubmission();
    return () => {
      cancelled = true;
    };
  }, [submissionId]);
  if (loading) {
    return (
      <div className="grid min-h-56 place-items-center">
        <Spinner />
      </div>
    );
  }
  const backLink = (
    <Link
      to="/admin/forms/$formId/submissions"
      params={{ formId }}
      className="mb-6 inline-flex items-center gap-2 text-sm font-semibold text-[var(--ink-soft)] hover:text-[var(--ink)]"
    >
      <ArrowLeft size={15} />
      Back to submissions
    </Link>
  );

  if (error) {
    return (
      <div>
        {backLink}
        <Notice tone="danger">{error}</Notice>
      </div>
    );
  }

  if (!submission) {
    return (
      <div>
        {backLink}
        <Notice tone="danger">Submission not found.</Notice>
      </div>
    );
  }

  const downloadDocx = async () => {
    await downloadArtifact(
      `/api/submissions/${submissionId}/docx?revision=latest`,
      `${submissionId}-latest.docx`
    );
  };

  const downloadPdf = async () => {
    await downloadArtifact(
      `/api/submissions/${submissionId}/pdf?revision=latest`,
      `${submissionId}-latest.pdf`
    );
  };

  return (
    <div>
      {backLink}
      <div className="mb-7 flex flex-wrap items-end justify-between gap-4">
        <div>
          <div className="mb-2 flex items-center gap-2">
            <Badge tone="success">{submission.status ?? "Submitted"}</Badge>
            <span className="text-sm text-[var(--ink-soft)]">
              {formatDate(submission.submittedAt ?? submission.createdAt)}
            </span>
          </div>
          <h1 className="text-3xl font-bold tracking-[-0.04em]">
            Submission detail
          </h1>
          <p className="mt-2 text-[var(--ink-soft)]">
            {submission.userEmail ?? "Respondent"}
          </p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={downloadDocx}
            className="inline-flex min-h-10 items-center gap-2 rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-sm font-semibold hover:border-[var(--ink)]"
          >
            <FileText size={15} />
            DOCX
          </button>
          <button
            onClick={downloadPdf}
            className="inline-flex min-h-10 items-center gap-2 rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-sm font-semibold hover:border-[var(--ink)]"
          >
            <Download size={15} />
            PDF
          </button>
        </div>
      </div>
      <Card className="overflow-hidden p-5">
        <SubmissionResultViewer
          configUrl={
            submission.responseId
              ? `/api/admin/results/${submission.responseId}/viewer-config?revision=latest`
              : undefined
          }
          data={data ?? {}}
          documentAvailable={documentAvailable}
          fields={fields}
          pictures={pictures}
        />
      </Card>
    </div>
  );
};

export const Route = createFileRoute(
  "/admin/forms/$formId/submissions/$submissionId"
)({ component: SubmissionDetailRoute });
