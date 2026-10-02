// oxlint-disable unicorn/filename-case -- TanStack Router requires this dynamic route filename.
import { createFileRoute, Link, useParams } from "@tanstack/react-router";
import { FileText, LockKeyhole, Trash2 } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import {
  SubmissionResultViewer,
  SubmissionRevisionHistory,
} from "@/components/submission-result-viewer";
import { Badge, Button, Card, Notice, Spinner } from "@/components/ui";
import {
  ApiError,
  apiDelete,
  apiGet,
  downloadArtifact,
  formatDate,
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
    <SubmissionRevisionHistory className="mt-4 space-y-3" revisions={history} />
  );
};

const RevisionSelector = ({
  history,
  historyError,
  historyLoading,
  setViewRevision,
  viewRevision,
}: {
  history: ResponseRevision[];
  historyError: string | null;
  historyLoading: boolean;
  setViewRevision: (revision: number) => void;
  viewRevision: number;
}) => (
  <label className="flex items-center gap-2 text-sm font-semibold">
    <span>ดูข้อมูลรุ่น</span>
    <select
      aria-label="เลือก Revision"
      className="min-h-10 rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3"
      disabled={historyLoading || historyError !== null}
      onChange={(event) => setViewRevision(Number(event.currentTarget.value))}
      value={viewRevision}
    >
      {history.map((revision) => (
        <option key={revision.revision} value={revision.revision}>
          {revision.revision === 0
            ? "Submission เดิม"
            : `Correction ${revision.revision}`}
        </option>
      ))}
    </select>
  </label>
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
  const [viewRevision, setViewRevision] = useState<number | null>(null);
  const [history, setHistory] = useState<
    ResponseRevisionsResponse["revisions"]
  >([]);
  const [fields, setFields] = useState<ReceiptField[]>([]);
  const [fieldsRevision, setFieldsRevision] = useState<number | null>(null);
  const [fieldsError, setFieldsError] = useState<string | null>(null);
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
    if (!result || result.state !== "submitted" || !result.submissionId) {
      return;
    }
    let cancelled = false;
    setHistoryLoading(true);
    setHistoryError(null);
    setFieldsError(null);
    const loadHistory = async () => {
      try {
        const historyResult = await apiGet<ResponseRevisionsResponse>(
          `/api/responses/${responseId}/corrections`
        );
        if (!cancelled) {
          setHistory(historyResult.revisions);
          setViewRevision(historyResult.latestRevision);
        }
      } catch {
        if (!cancelled) {
          setHistory([]);
          setHistoryError("โหลดประวัติ Correction ไม่สำเร็จ กรุณาลองใหม่");
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
    if (!result || result.state !== "submitted" || !result.submissionId) {
      return;
    }
    const selectedRevision = viewRevision ?? result.latestCorrectionNumber ?? 0;
    let cancelled = false;
    setFields([]);
    setFieldsRevision(null);
    setFieldsError(null);
    const loadFields = async () => {
      try {
        const payload = await apiGet<{ fields: ReceiptField[] }>(
          `/api/submissions/${result.submissionId}/data?revision=${selectedRevision}`
        );
        if (!cancelled) {
          setFields(payload.fields);
          setFieldsRevision(selectedRevision);
        }
      } catch {
        if (!cancelled) {
          setFieldsError("ไม่สามารถโหลดข้อมูล Field ที่เผยแพร่ได้");
        }
      }
    };
    void loadFields();
    return () => {
      cancelled = true;
    };
  }, [result, viewRevision]);
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

  const download = async (format: "docx" | "pdf", revision: number) => {
    if (!result?.submissionId || downloading) {
      return;
    }
    setDownloading(`${format}-${revision}`);
    setDownloadError(null);
    try {
      const query = `?revision=${revision}`;
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
  const selectedRevision = viewRevision ?? latestRevision;
  const selectedRevisionRecord = history.find(
    (revision) => revision.revision === selectedRevision
  );
  let viewerConfigUrl: string | undefined;
  if (isSubmitted && selectedRevisionRecord) {
    viewerConfigUrl = `/api/admin/results/${responseId}/viewer-config?revision=${selectedRevision}`;
  } else if (isDraft && result.document.available) {
    viewerConfigUrl = `/api/admin/results/${responseId}/viewer-config`;
  }
  let submittedDataLabel = "ข้อมูลต้นฉบับของ Submission";
  if (selectedRevision > 0) {
    submittedDataLabel = `ข้อมูล Correction ${selectedRevision}`;
  }
  let displayedData: Record<string, unknown>;
  let displayedDocumentAvailable: boolean;
  let displayedFields: ReceiptField[];
  let displayedFieldsError: string | null = null;
  let displayedPictures: Record<string, boolean> | null = null;
  if (isDraft) {
    displayedData = result.data;
    displayedDocumentAvailable = result.document.available;
    displayedFields = result.fields;
  } else {
    displayedData = selectedRevisionRecord?.data ?? {};
    displayedDocumentAvailable =
      selectedRevisionRecord?.document.available ?? false;
    displayedFields = fieldsRevision === selectedRevision ? fields : [];
    displayedFieldsError = fieldsError;
    displayedPictures = selectedRevisionRecord?.pictures ?? null;
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
                history={history}
                historyError={historyError}
                historyLoading={historyLoading}
                setViewRevision={setViewRevision}
                viewRevision={selectedRevision}
              />
              <Link
                search={{ form: undefined }}
                to="/admin/results/$responseId/correction"
                params={{ responseId }}
              >
                <Button variant="primary">เปิด Correction</Button>
              </Link>
              <Button
                disabled={downloading !== null || !selectedRevisionRecord}
                onClick={() => download("docx", selectedRevision)}
                variant="secondary"
              >
                <FileText size={16} />
                DOCX
              </Button>
              <Button
                disabled={downloading !== null || !selectedRevisionRecord}
                onClick={() => download("pdf", selectedRevision)}
                variant="secondary"
              >
                <FileText size={16} />
                PDF
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
        <Card>
          <p className="mb-4 text-sm text-[var(--ink-soft)]">
            {isDraft ? "ข้อมูลฉบับร่าง" : submittedDataLabel}
          </p>
          <SubmissionResultViewer
            configUrl={viewerConfigUrl}
            data={displayedData}
            documentAvailable={displayedDocumentAvailable}
            fields={displayedFields}
            fieldsError={displayedFieldsError}
            pictures={displayedPictures}
          />
        </Card>
      </div>
    </>
  );
};

export const Route = createFileRoute("/admin/results/$responseId/")({
  component: AdminResultDetailRoute,
});
