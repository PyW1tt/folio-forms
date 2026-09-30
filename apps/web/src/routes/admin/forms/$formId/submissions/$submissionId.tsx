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

// oxlint-disable-next-line complexity -- The detail route owns read, history, revision, and export state.
const SubmissionDetailRoute = () => {
  const { formId, submissionId } = useParams({
    from: "/admin/forms/$formId/submissions/$submissionId",
  });
  const [submission, setSubmission] = useState<Submission | null>(null);
  const [data, setData] = useState<Record<string, unknown> | null>(null);
  const [dataRevision, setDataRevision] = useState<number | null>(null);
  const [fields, setFields] = useState<ReceiptField[]>([]);
  const [revisionError, setRevisionError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [historyError, setHistoryError] = useState<string | null>(null);
  const [history, setHistory] = useState<
    ResponseRevisionsResponse["revisions"]
  >([]);
  const [viewRevision, setViewRevision] = useState<number | null>(null);

  useEffect(() => {
    let cancelled = false;
    const loadHistory = async (responseId: string) => {
      try {
        const payload = await apiGet<ResponseRevisionsResponse>(
          `/api/responses/${responseId}/corrections`
        );
        if (!cancelled) {
          setHistory(payload.revisions);
          setViewRevision(payload.latestRevision);
        }
      } catch {
        if (!cancelled) {
          setHistoryError("Could not load correction history.");
        }
      }
    };
    const loadSubmission = async () => {
      setHistoryError(null);
      try {
        const payload = await apiGet<{
          data: Record<string, unknown>;
          fields: ReceiptField[];
          submission: Submission;
        }>(`/api/submissions/${submissionId}/data?revision=latest`);
        if (cancelled) {
          return;
        }
        setSubmission(payload.submission);
        if (payload.submission.responseId) {
          void loadHistory(payload.submission.responseId);
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
  useEffect(() => {
    if (!submission || viewRevision === null) {
      return;
    }
    let cancelled = false;
    setData(null);
    setDataRevision(null);
    setFields([]);
    setRevisionError(null);
    const loadRevision = async () => {
      try {
        const payload = await apiGet<{
          data: Record<string, unknown>;
          fields: ReceiptField[];
        }>(`/api/submissions/${submissionId}/data?revision=${viewRevision}`);
        if (!cancelled) {
          setData(payload.data);
          setDataRevision(viewRevision);
          setFields(payload.fields);
        }
      } catch {
        if (!cancelled) {
          setData(null);
          setFields([]);
          setRevisionError("Could not load selected revision.");
        }
      }
    };
    void loadRevision();
    return () => {
      cancelled = true;
    };
  }, [submission, submissionId, viewRevision]);
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

  const selectedRevision = viewRevision ?? history.at(-1)?.revision ?? 0;
  const selectedRevisionRecord = history.find(
    (revision) => revision.revision === selectedRevision
  );
  const downloadDocx = async () => {
    await downloadArtifact(
      `/api/submissions/${submissionId}/docx?revision=${selectedRevision}`,
      `${submissionId}-${selectedRevision}.docx`
    );
  };

  const downloadPdf = async () => {
    await downloadArtifact(
      `/api/submissions/${submissionId}/pdf?revision=${selectedRevision}`,
      `${submissionId}-${selectedRevision}.pdf`
    );
  };

  return (
    <div>
      {backLink}
      {historyError ? (
        <div className="mb-4">
          <Notice tone="danger">{historyError}</Notice>
        </div>
      ) : null}
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
        <label className="flex items-center gap-2 text-sm font-semibold">
          <span>Revision</span>
          <select
            aria-label="Select revision"
            className="min-h-10 rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3"
            disabled={historyError !== null || history.length === 0}
            onChange={(event) =>
              setViewRevision(Number(event.currentTarget.value))
            }
            value={selectedRevision}
          >
            {history.map((revision) => (
              <option key={revision.revision} value={revision.revision}>
                {revision.revision === 0
                  ? "Original submission"
                  : `Correction ${revision.revision}`}
              </option>
            ))}
          </select>
        </label>
        <div className="flex gap-2">
          <button
            disabled={!selectedRevisionRecord}
            onClick={downloadDocx}
            className="inline-flex min-h-10 items-center gap-2 rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-sm font-semibold hover:border-[var(--ink)]"
          >
            <FileText size={15} />
            DOCX
          </button>
          <button
            disabled={!selectedRevisionRecord}
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
            submission.responseId && selectedRevisionRecord
              ? `/api/admin/results/${submission.responseId}/viewer-config?revision=${selectedRevision}`
              : undefined
          }
          data={dataRevision === selectedRevision ? (data ?? {}) : {}}
          documentAvailable={
            selectedRevisionRecord?.document.available ?? false
          }
          fields={dataRevision === selectedRevision ? fields : []}
          fieldsError={historyError ?? revisionError}
          pictures={selectedRevisionRecord?.pictures ?? null}
        />
      </Card>
    </div>
  );
};

export const Route = createFileRoute(
  "/admin/forms/$formId/submissions/$submissionId"
)({ component: SubmissionDetailRoute });
