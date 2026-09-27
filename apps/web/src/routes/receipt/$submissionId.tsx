// oxlint-disable unicorn(filename-case) -- TanStack Router requires this dynamic route filename.
import { createFileRoute, Link, useParams } from "@tanstack/react-router";
import { ArrowLeft, CheckCircle2, Download, FileText } from "lucide-react";
import { useEffect, useState } from "react";

import { Badge, Card, Notice, Spinner } from "@/components/ui";
import {
  ApiError,
  apiGet,
  downloadArtifact,
  formatDate,
  formatDateTime,
} from "@/lib/api";
import type {
  ReceiptField,
  ResponseRevisionsResponse,
  Submission,
} from "@/lib/api";

const receiptErrorMessage = (error: unknown): string => {
  if (error instanceof ApiError) {
    if (error.status === 401) {
      return "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่";
    }
    if (error.status === 403 || error.code === "forbidden") {
      return "คุณไม่มีสิทธิ์ดูใบรับคำตอบนี้";
    }
    if (error.status === 404 || error.code === "not_found") {
      return "ไม่พบใบรับคำตอบนี้";
    }
  }
  return "โหลดใบรับคำตอบไม่ได้ กรุณาลองใหม่";
};
const receiptFieldValue = (field: ReceiptField, value: unknown): string => {
  if (field.type === "picture") {
    return "เปิดเอกสาร DOCX เพื่อดูรูปภาพ";
  }
  if (value === null || value === undefined || value === "") {
    return "—";
  }
  if (field.type === "checkbox" && typeof value === "boolean") {
    return value ? "ใช่" : "ไม่ใช่";
  }
  if (typeof value === "string") {
    return (
      field.options.find((option) => option.value === value)?.displayText ??
      value
    );
  }
  if (typeof value === "number") {
    return String(value);
  }
  return "มีข้อมูลที่บันทึกแล้ว";
};
const ReceiptRoute = () => {
  const { submissionId } = useParams({ from: "/receipt/$submissionId" });
  const [submission, setSubmission] = useState<Submission | null>(null);
  const [fields, setFields] = useState<ReceiptField[]>([]);
  const [data, setData] = useState<Record<string, unknown> | null>(null);
  const [revisions, setRevisions] = useState<
    ResponseRevisionsResponse["revisions"]
  >([]);
  const [viewRevision, setViewRevision] = useState<"original" | "latest">(
    "latest"
  );
  const [returnUrl, setReturnUrl] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    const loadReceipt = async () => {
      try {
        const payload = await apiGet<{
          data: Record<string, unknown>;
          fields: ReceiptField[];
          returnUrl: string;
          submission: Submission;
        }>(`/api/submissions/${submissionId}/data`);
        if (!payload.submission.responseId) {
          throw new Error("Response is unavailable");
        }
        const history = await apiGet<ResponseRevisionsResponse>(
          `/api/responses/${payload.submission.responseId}/corrections`
        );
        if (!cancelled) {
          setViewRevision("latest");
          setData(payload.data);
          setFields(payload.fields);
          setRevisions(history.revisions);
          setReturnUrl(payload.returnUrl);
          setSubmission(payload.submission);
        }
      } catch (caughtError) {
        if (!cancelled) {
          setError(receiptErrorMessage(caughtError));
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    void loadReceipt();
    return () => {
      cancelled = true;
    };
  }, [submissionId]);
  if (loading) {
    return (
      <div className="grid min-h-screen place-items-center">
        <Spinner />
      </div>
    );
  }
  if (error) {
    return (
      <div className="min-h-screen bg-[var(--canvas)] px-5 py-12">
        <div className="mx-auto max-w-3xl">
          <Notice tone="danger">{error}</Notice>
        </div>
      </div>
    );
  }

  if (!submission) {
    return (
      <div className="min-h-screen bg-[var(--canvas)] px-5 py-12">
        <div className="mx-auto max-w-3xl">
          <Notice tone="danger">ไม่พบใบรับคำตอบนี้</Notice>
        </div>
      </div>
    );
  }

  const latestRevision = revisions.at(-1)?.revision ?? 0;
  const selectedRevision = viewRevision === "latest" ? latestRevision : 0;
  const displayedData =
    revisions.find((revision) => revision.revision === selectedRevision)
      ?.data ??
    data ??
    {};
  const revisionQuery = viewRevision === "latest" ? "?revision=latest" : "";
  const revisionSuffix =
    viewRevision === "latest" && latestRevision > 0
      ? `-revision-${latestRevision}`
      : "";
  const download = async (format: "docx" | "pdf") => {
    await downloadArtifact(
      `/api/submissions/${submissionId}/${format}${revisionQuery}`,
      `${submissionId}${revisionSuffix}.${format}`
    );
  };

  return (
    <div className="min-h-screen bg-[var(--canvas)] px-5 py-12">
      <div className="mx-auto max-w-3xl">
        <div className="mb-8 flex items-center justify-between gap-4">
          <Link
            to="/dashboard"
            className="inline-flex items-center gap-2 text-sm font-semibold text-[var(--ink-soft)] hover:text-[var(--ink)]"
          >
            <ArrowLeft size={15} />
            กลับไปยังคำตอบของฉัน
          </Link>
          <span className="text-sm font-bold tracking-[-0.03em]">
            Folio Forms
          </span>
        </div>
        <Card className="overflow-hidden">
          <div className="border-b border-[var(--line)] bg-[var(--success-soft)] p-7 sm:p-9">
            <CheckCircle2 className="mb-4 text-[var(--success)]" size={35} />
            <p className="text-sm font-semibold uppercase tracking-[0.12em] text-[var(--success)]">
              ส่งคำตอบเรียบร้อยแล้ว
            </p>
            <h1 className="mt-2 text-3xl font-bold tracking-[-0.04em]">
              ระบบบันทึกคำตอบของคุณแล้ว
            </h1>
            <p className="mt-2 text-[var(--ink-soft)]">
              {submission.formTitle ?? "แบบฟอร์ม"} · ส่งเมื่อ{" "}
              {formatDate(submission.submittedAt ?? submission.createdAt)}
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-3 border-b border-[var(--line)] p-5">
            <span className="text-sm font-semibold">ฉบับที่ดู</span>
            <button
              type="button"
              aria-pressed={viewRevision === "original"}
              onClick={() => setViewRevision("original")}
              className="inline-flex min-h-10 items-center rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-sm font-semibold hover:border-[var(--ink)]"
            >
              เดิม
            </button>
            <button
              type="button"
              aria-pressed={viewRevision === "latest"}
              onClick={() => setViewRevision("latest")}
              className="inline-flex min-h-10 items-center rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-sm font-semibold hover:border-[var(--ink)]"
            >
              ล่าสุด{latestRevision > 0 ? ` (Correction ${latestRevision})` : ""}
            </button>
            <button
              type="button"
              onClick={() => download("docx")}
              className="inline-flex min-h-10 items-center gap-2 rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-sm font-semibold hover:border-[var(--ink)]"
            >
              <FileText size={15} />
              DOCX
            </button>
            <button
              type="button"
              onClick={() => download("pdf")}
              className="inline-flex min-h-10 items-center gap-2 rounded-[10px] border border-[var(--ink)] bg-[var(--ink)] px-3 text-sm font-semibold text-[var(--paper)]"
            >
              <Download size={15} />
              PDF
            </button>
            {returnUrl ? (
              <a
                href={returnUrl}
                className="inline-flex min-h-10 items-center rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-sm font-semibold hover:border-[var(--ink)]"
              >
                กลับไปยังระบบต้นทาง
              </a>
            ) : null}
          </div>
          <div className="border-b border-[var(--line)] p-5 sm:p-7">
            <h2 className="font-semibold">ประวัติ Correction</h2>
            <ol className="mt-3 space-y-3">
              {revisions.map((revision) => (
                <li
                  className="rounded-[10px] border border-[var(--line)] p-3 text-sm"
                  key={revision.revision}
                >
                  <div className="flex flex-wrap items-center gap-2">
                    <Badge
                      tone={revision.revision === 0 ? "neutral" : "warning"}
                    >
                      {revision.revision === 0
                        ? "Submission เดิม"
                        : `Correction ${revision.revision}`}
                    </Badge>
                    <span className="text-[var(--ink-soft)]">
                      {formatDateTime(revision.createdAt)}
                    </span>
                  </div>
                  {revision.reason ? (
                    <p className="mt-2 text-[var(--ink-soft)]">
                      เหตุผล: {revision.reason}
                    </p>
                  ) : null}
                  {revision.actorName || revision.actorEmail ? (
                    <p className="mt-1 text-[var(--ink-soft)]">
                      ผู้แก้ไข: {revision.actorName ?? revision.actorEmail}
                    </p>
                  ) : null}
                </li>
              ))}
            </ol>
          </div>
          <div className="p-5 sm:p-7">
            <h2 className="mb-4 font-semibold">ข้อมูลคำตอบ</h2>
            <dl className="space-y-3">
              {fields.map((field) => (
                <div
                  className="rounded-[10px] border border-[var(--line)] p-4"
                  key={field.tag}
                >
                  <dt className="font-semibold">{field.label}</dt>
                  <dd className="mt-2 whitespace-pre-wrap break-words text-[var(--ink-soft)]">
                    {receiptFieldValue(field, displayedData[field.tag])}
                  </dd>
                </div>
              ))}
            </dl>
          </div>
        </Card>
      </div>
    </div>
  );
};

export const Route = createFileRoute("/receipt/$submissionId")({
  component: ReceiptRoute,
});
