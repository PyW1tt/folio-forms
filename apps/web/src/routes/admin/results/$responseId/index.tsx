// oxlint-disable unicorn/filename-case -- TanStack Router requires this dynamic route filename.
import { createFileRoute, Link, useParams } from "@tanstack/react-router";
import { FileText, LockKeyhole, Trash2 } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import { SubmissionResultViewer } from "@/components/submission-result-viewer";
import { Badge, Button, Card, Notice, Spinner } from "@/components/ui";
import {
  ApiError,
  apiDelete,
  apiGet,
  downloadArtifact,
  formatDate,
  formatDateTime,
} from "@/lib/api";
import type {
  AdminResultDetail,
  ReceiptField,
  ResponseRevisionsResponse,
} from "@/lib/api";

type ResponseRevision = ResponseRevisionsResponse["revisions"][number];

const CorrectionHistory = ({
  error,
  history,
  loading,
}: {
  error: string | null;
  history: ResponseRevision[];
  loading: boolean;
}) => {
  if (loading) {
    return (
      <div className="mt-4">
        <Spinner />
      </div>
    );
  }
  if (error) {
    return (
      <div className="mt-4">
        <Notice tone="danger">{error}</Notice>
      </div>
    );
  }
  return (
    <ol className="mt-4 space-y-3">
      {history.map((revision) => (
        <li
          className="rounded-[10px] border border-[var(--line)] p-3 text-sm"
          key={revision.revision}
        >
          <div className="flex flex-wrap items-center gap-2">
            <Badge tone={revision.revision === 0 ? "neutral" : "warning"}>
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
  );
};

const RevisionSelector = ({
  hasOriginalRevision,
  historyError,
  historyLoading,
  latestRevision,
  setViewRevision,
  viewRevision,
}: {
  hasOriginalRevision: boolean;
  historyError: string | null;
  historyLoading: boolean;
  latestRevision: number;
  setViewRevision: (revision: "original" | "latest") => void;
  viewRevision: "original" | "latest";
}) => (
  <div
    className="flex w-full flex-wrap items-center gap-2"
    aria-label="เลือกข้อมูลคำตอบ"
  >
    <span className="mr-1 text-sm font-semibold">ดูข้อมูล</span>
    <Button
      aria-pressed={viewRevision === "original"}
      disabled={historyLoading || historyError !== null || !hasOriginalRevision}
      onClick={() => setViewRevision("original")}
      variant={viewRevision === "original" ? "primary" : "secondary"}
    >
      ข้อมูลเดิม
    </Button>
    <Button
      aria-pressed={viewRevision === "latest"}
      onClick={() => setViewRevision("latest")}
      variant={viewRevision === "latest" ? "primary" : "secondary"}
    >
      ข้อมูลล่าสุด
      {latestRevision > 0 ? ` (Correction ${latestRevision})` : ""}
    </Button>
  </div>
);

const deletionErrorMessage = (caughtError: unknown) => {
  if (caughtError instanceof ApiError && caughtError.code === "editor_in_use") {
    return "คำตอบกำลังเปิดอยู่ในตัวแก้ไข กรุณาปิดหน้าต่างนั้นแล้วลองใหม่";
  }
  if (
    caughtError instanceof ApiError &&
    caughtError.code === "deletion_cleanup_failed"
  ) {
    return "ลบข้อมูลแล้ว แต่ไฟล์ยังอยู่ระหว่างการล้างข้อมูล กรุณาลองใหม่";
  }
  return "ลบคำตอบไม่สำเร็จ กรุณาลองใหม่อีกครั้ง";
};
// oxlint-disable-next-line complexity -- The detail route owns read, history, export, and deletion states.
const AdminResultDetailRoute = () => {
  const { responseId } = useParams({ from: "/admin/results/$responseId/" });
  const [result, setResult] = useState<AdminResultDetail | null>(null);
  const [viewRevision, setViewRevision] = useState<"original" | "latest">(
    "latest"
  );
  const [history, setHistory] = useState<
    ResponseRevisionsResponse["revisions"]
  >([]);
  const [fields, setFields] = useState<ReceiptField[]>([]);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [historyError, setHistoryError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const [downloading, setDownloading] = useState<string | null>(null);
  const [deleteConfirmation, setDeleteConfirmation] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [deleteSuccess, setDeleteSuccess] = useState(false);
  const deleteTriggerRef = useRef<HTMLButtonElement | null>(null);
  const deleteFeedbackRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    let cancelled = false;
    const loadResult = async () => {
      try {
        const payload = await apiGet<{ result: AdminResultDetail }>(
          `/api/admin/results/${responseId}`
        );
        if (!cancelled) {
          setResult(payload.result);
        }
      } catch (caughtError) {
        if (!cancelled) {
          setError(
            caughtError instanceof ApiError && caughtError.status === 403
              ? "คุณไม่มีสิทธิ์ดูผลลัพธ์นี้"
              : "ไม่สามารถโหลดรายละเอียดได้ กรุณาลองใหม่อีกครั้ง"
          );
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };
    void loadResult();
    return () => {
      cancelled = true;
    };
  }, [responseId]);

  useEffect(() => {
    if (
      !result ||
      result.state !== "submitted" ||
      !result.submissionId
    ) {
      return;
    }
    const submissionId = result.submissionId;
    let cancelled = false;
    setHistoryLoading(true);
    setHistoryError(null);
    const loadHistory = async () => {
      try {
        const [historyResult, submissionDataResult] = await Promise.allSettled([
          apiGet<ResponseRevisionsResponse>(
            `/api/responses/${responseId}/corrections`
          ),
          apiGet<{ fields: ReceiptField[] }>(
            `/api/submissions/${submissionId}/data?revision=latest`
          ),
        ]);
        if (!cancelled) {
          if (historyResult.status === "fulfilled") {
            setHistory(historyResult.value.revisions);
          } else {
            setHistory([]);
            setHistoryError("โหลดประวัติ Correction ไม่สำเร็จ กรุณาลองใหม่");
          }
          if (submissionDataResult.status === "fulfilled") {
            setFields(submissionDataResult.value.fields);
          }
        }
      } finally {
        if (!cancelled) {
          setHistoryLoading(false);
        }
      }
    };
    void loadHistory();
    return () => {
      cancelled = true;
    };
  }, [responseId, result]);
  useEffect(() => {
    if (deleteConfirmation) {
      document
        .querySelector<HTMLButtonElement>("#delete-response-cancel")
        ?.focus();
    }
  }, [deleteConfirmation]);

  useEffect(() => {
    if (deleteError || deleteSuccess) {
      deleteFeedbackRef.current?.focus();
    }
  }, [deleteError, deleteSuccess]);

  const requestDelete = (trigger: HTMLButtonElement) => {
    deleteTriggerRef.current = trigger;
    setDeleteError(null);
    setDeleteConfirmation(true);
  };

  const cancelDelete = () => {
    if (deleting) {
      return;
    }
    setDeleteConfirmation(false);
    setDeleteError(null);
    deleteTriggerRef.current?.focus();
    deleteTriggerRef.current = null;
  };

  const deleteResult = async () => {
    if (!result || deleting) {
      return;
    }
    setDeleting(true);
    setDeleteError(null);
    try {
      await apiDelete(`/api/admin/responses/${responseId}`, { confirm: true });
      setDeleteConfirmation(false);
      setResult(null);
      setDeleteSuccess(true);
    } catch (caughtError) {
      setDeleteError(deletionErrorMessage(caughtError));
    } finally {
      setDeleting(false);
    }
  };

  const download = async (
    format: "docx" | "pdf",
    revision: "original" | "latest"
  ) => {
    if (!result?.submissionId || downloading) {
      return;
    }
    setDownloading(`${format}-${revision}`);
    setDownloadError(null);
    try {
      const query = revision === "latest" ? "?revision=latest" : "";
      await downloadArtifact(
        `/api/submissions/${result.submissionId}/${format}${query}`,
        `submission-${result.submissionId}-${revision}.${format}`
      );
    } catch {
      setDownloadError("ส่งออกไฟล์ไม่สำเร็จ กรุณาลองใหม่อีกครั้ง");
    } finally {
      setDownloading(null);
    }
  };
  if (deleteSuccess) {
    return (
      <div
        ref={deleteFeedbackRef}
        className="space-y-4 focus:outline-none"
        role="status"
        tabIndex={-1}
      >
        <Notice tone="success">
          <strong>ลบข้อมูลคำตอบถาวรแล้ว</strong>
          <div>ระบบลบข้อมูลส่วนบุคคล เอกสาร และข้อมูล Prefill ที่เกี่ยวข้องแล้ว</div>
        </Notice>
        <Link search={{ form: undefined }} to="/admin/results">
          <Button variant="secondary">กลับไปรายการผลลัพธ์</Button>
        </Link>
      </div>
    );
  }

  if (loading) {
    return <Spinner />;
  }
  if (error || !result) {
    return (
      <Notice tone="danger">
        <strong>เปิดผลลัพธ์ไม่สำเร็จ</strong>
        <div>{error ?? "ไม่พบผลลัพธ์นี้"}</div>
      </Notice>
    );
  }
  const isSubmitted = result.state === "submitted";
  const isDraft = !isSubmitted;
  const latestRevision =
    history.at(-1)?.revision ?? result.latestCorrectionNumber ?? 0;
  const hasOriginalRevision = history.some(
    (revision) => revision.revision === 0
  );
  const selectedRevision = viewRevision === "latest" ? latestRevision : 0;
  const selectedRevisionRecord = history.find(
    (revision) => revision.revision === selectedRevision
  );
  const displayedData = selectedRevisionRecord?.data ?? result.data;
  const selectedPictures = selectedRevisionRecord?.pictures ?? null;
  const documentAvailable =
    selectedRevisionRecord?.document.available ??
    (viewRevision === "latest" && result.document.available);
  let submittedDataLabel = "ข้อมูลต้นฉบับของ Submission";
  if (viewRevision === "latest" && latestRevision > 0) {
    submittedDataLabel = "ข้อมูลล่าสุดของ Correction";
  }
  return (
    <>
      {deleteConfirmation ? (
        <div
          className="fixed inset-0 z-50 grid place-items-center bg-black/35 px-5"
          role="presentation"
        >
          <div
            className="w-full max-w-lg rounded-[var(--radius)] border border-[var(--line)] bg-[var(--paper)] p-6 shadow-xl"
            role="dialog"
            aria-modal="true"
            aria-labelledby="delete-response-title"
          >
            <h2 id="delete-response-title" className="text-lg font-semibold">
              ยืนยันการลบคำตอบถาวร
            </h2>
            <p className="mt-2 text-sm text-[var(--ink-soft)]">
              ระบบจะลบ Draft หรือ Submission นี้พร้อม Prefill, ประวัติ Correction,
              เอกสาร และเซสชันของเจ้าของบัญชี การลบไม่สามารถกู้คืนได้
            </p>
            {deleteError ? (
              <div
                ref={deleteFeedbackRef}
                className="mt-4 focus:outline-none"
                role="alert"
                tabIndex={-1}
              >
                <Notice tone="danger">{deleteError}</Notice>
              </div>
            ) : null}
            <div className="mt-6 flex justify-end gap-2">
              <Button
                id="delete-response-cancel"
                type="button"
                variant="ghost"
                onClick={cancelDelete}
                disabled={deleting}
              >
                ยกเลิก
              </Button>
              <Button
                type="button"
                variant="danger"
                onClick={deleteResult}
                disabled={deleting}
              >
                {deleting ? <Spinner /> : <Trash2 size={16} />}
                {deleting ? "กำลังลบ…" : "ยืนยันการลบถาวร"}
              </Button>
            </div>
          </div>
        </div>
      ) : null}
      <div className="space-y-6">
        <PageHeader
          title={result.formTitle}
          description={`${result.userEmail} · ${result.formPublicId}`}
          action={
            <Link to="/admin/results" search={{ form: undefined }}>
              <Button variant="secondary">กลับไปรายการ</Button>
            </Link>
          }
        />
        <Card>
          <div className="flex flex-wrap items-center gap-3">
            <Badge tone={isSubmitted ? "success" : "warning"}>
              {isSubmitted ? "ส่งแล้ว" : "ฉบับร่าง"}
            </Badge>
            <Badge tone="neutral">
              <LockKeyhole className="mr-1 inline" size={14} />
              อ่านอย่างเดียว
            </Badge>
            <span className="text-sm text-[var(--ink-soft)]">
              อัปเดต {formatDate(result.updatedAt)} · Correction{" "}
              {result.latestCorrectionNumber ?? "—"}
            </span>
          </div>
          {isDraft ? (
            <Notice>
              <strong>ฉบับร่างอ่านได้อย่างเดียว</strong>
              <div>
                Admin ไม่สามารถแก้ไข ยึด Lease หรือส่งออกฉบับร่างจาก workflow นี้ได้
              </div>
            </Notice>
          ) : (
            <div className="mt-5 flex flex-wrap gap-2" aria-label="จัดการคำตอบ">
              <RevisionSelector
                hasOriginalRevision={hasOriginalRevision}
                historyError={historyError}
                historyLoading={historyLoading}
                latestRevision={latestRevision}
                setViewRevision={setViewRevision}
                viewRevision={viewRevision}
              />
              <Link
                search={{ form: undefined }}
                to="/admin/results/$responseId/correction"
                params={{ responseId }}
              >
                <Button variant="primary">เปิด Correction</Button>
              </Link>
              <Button
                disabled={downloading !== null}
                onClick={() => download("docx", "original")}
                variant="secondary"
              >
                <FileText size={16} />
                DOCX เดิม
              </Button>
              <Button
                disabled={downloading !== null}
                onClick={() => download("pdf", "original")}
                variant="secondary"
              >
                <FileText size={16} />
                PDF เดิม
              </Button>
              <Button
                disabled={downloading !== null}
                onClick={() => download("docx", "latest")}
                variant="secondary"
              >
                <FileText size={16} />
                DOCX ล่าสุด
              </Button>
              <Button
                disabled={downloading !== null}
                onClick={() => download("pdf", "latest")}
                variant="secondary"
              >
                <FileText size={16} />
                PDF ล่าสุด
              </Button>
            </div>
          )}
          {downloadError ? (
            <Notice tone="danger">
              <strong>ส่งออกไม่สำเร็จ</strong>
              <div>{downloadError}</div>
            </Notice>
          ) : null}
          <div className="mt-5 border-t border-[var(--line)] pt-4">
            <Button
              type="button"
              variant="danger"
              onClick={(event) => requestDelete(event.currentTarget)}
              disabled={deleting || downloading !== null}
            >
              <Trash2 size={16} />
              ลบคำตอบถาวร
            </Button>
          </div>
        </Card>
        {isSubmitted ? (
          <Card>
            <h2 className="font-semibold">ประวัติ Correction</h2>
            <CorrectionHistory
              error={historyError}
              history={history}
              loading={historyLoading}
            />
          </Card>
        ) : null}
        {isSubmitted ? (
          <Card>
            <p className="mb-4 text-sm text-[var(--ink-soft)]">
              {submittedDataLabel}
            </p>
            <SubmissionResultViewer
              configUrl={`/api/admin/results/${responseId}/viewer-config?revision=${viewRevision}`}
              data={displayedData}
              documentAvailable={documentAvailable}
              fields={fields}
              pictures={selectedPictures}
            />
          </Card>
        ) : (
          <Card>
            <h2 className="font-semibold">ข้อมูลคำตอบ</h2>
            <div className="mt-4">
              <Notice>ข้อมูลฉบับร่างยังไม่แสดงในหน้าตรวจทานนี้</Notice>
            </div>
          </Card>
        )}
      </div>
    </>
  );
};

export const Route = createFileRoute("/admin/results/$responseId/")({
  component: AdminResultDetailRoute,
});
