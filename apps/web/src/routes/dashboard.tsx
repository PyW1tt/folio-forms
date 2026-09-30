import { createFileRoute, Link, Navigate } from "@tanstack/react-router";
import { ArrowUpRight, ClipboardList, FileCheck2, Inbox } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import { AppShell, PageHeader } from "@/components/app-shell";
import { Badge, Button, Card, Notice, Spinner } from "@/components/ui";
import { ApiError, apiDelete, apiGet, formatDate } from "@/lib/api";
import type { Submission } from "@/lib/api";
import { useAuth } from "@/lib/auth";

const Stat = ({
  label,
  value,
  icon,
}: {
  label: string;
  value: number;
  icon: React.ReactNode;
}) => (
  <Card className="flex items-center gap-4 p-5">
    <span className="grid size-10 place-items-center rounded-[10px] bg-[var(--accent-soft)] text-[var(--ink)]">
      {icon}
    </span>
    <div>
      <p className="text-2xl font-bold tracking-[-0.04em]">{value}</p>
      <p className="text-sm text-[var(--ink-soft)]">{label}</p>
    </div>
  </Card>
);

const Centered = ({ message }: { message: string }) => (
  <div className="grid min-h-64 place-items-center text-sm text-[var(--ink-soft)]">
    <Spinner />
    <span className="sr-only">{message}</span>
  </div>
);

const DashboardRoute = () => {
  const { user, loading: authLoading } = useAuth();
  const [rows, setRows] = useState<Submission[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [discardCandidateId, setDiscardCandidateId] = useState<string | null>(
    null
  );
  const [discardingId, setDiscardingId] = useState<string | null>(null);
  const [discardError, setDiscardError] = useState<string | null>(null);
  const [discardSuccess, setDiscardSuccess] = useState<string | null>(null);
  const discardTriggerRef = useRef<HTMLButtonElement | null>(null);
  const discardStatusRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (discardCandidateId) {
      document.querySelector<HTMLButtonElement>("#discard-cancel")?.focus();
    }
  }, [discardCandidateId]);
  useEffect(() => {
    if (discardSuccess) {
      discardStatusRef.current?.focus();
    }
  }, [discardSuccess]);
  useEffect(() => {
    if (!user) {
      setLoading(false);
      return;
    }

    let cancelled = false;
    const loadResponses = async () => {
      try {
        const payload = await apiGet<
          { responses: Submission[] } | Submission[]
        >("/api/responses/me");
        if (cancelled) {
          return;
        }
        setRows(Array.isArray(payload) ? payload : payload.responses);
      } catch (caughtError) {
        if (!cancelled) {
          setError(
            caughtError instanceof Error ? caughtError.message : "โหลดคำตอบไม่ได้"
          );
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    void loadResponses();
    return () => {
      cancelled = true;
    };
  }, [user]);
  const requestDiscard = (id: string, trigger: HTMLButtonElement) => {
    discardTriggerRef.current = trigger;
    setDiscardCandidateId(id);
    setDiscardError(null);
    setDiscardSuccess(null);
  };
  const restoreDiscardFocus = () => {
    const trigger = discardTriggerRef.current;
    discardTriggerRef.current = null;
    trigger?.focus();
  };
  const cancelDiscard = () => {
    if (discardingId) {
      return;
    }
    setDiscardCandidateId(null);
    setDiscardError(null);
    restoreDiscardFocus();
  };
  const discardDraft = async () => {
    if (!discardCandidateId || discardingId) {
      return;
    }
    const id = discardCandidateId;
    setDiscardingId(id);
    setDiscardError(null);
    try {
      await apiDelete(`/api/responses/${id}`);
      setRows((currentRows) => currentRows.filter((row) => row.id !== id));
      setDiscardCandidateId(null);
      setDiscardSuccess("ทิ้งฉบับร่างแล้ว");
    } catch (caughtError) {
      setDiscardError(
        caughtError instanceof ApiError && caughtError.status === 401
          ? "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่"
          : "ทิ้งฉบับร่างไม่สำเร็จ กรุณาลองใหม่"
      );
    } finally {
      setDiscardingId(null);
    }
  };

  if (authLoading) {
    return <Centered message="กำลังตรวจสอบเซสชัน…" />;
  }
  if (!user) {
    return <Navigate to="/login" search={{ returnTo: "/dashboard" }} />;
  }

  let responseContent: React.ReactNode;
  if (error) {
    responseContent = <Notice tone="danger">{error}</Notice>;
  } else if (loading) {
    responseContent = <Centered message="กำลังโหลดคำตอบ…" />;
  } else if (rows.length === 0) {
    responseContent = (
      <Card className="grid min-h-56 place-items-center p-8 text-center">
        <div>
          <Inbox className="mx-auto mb-3 text-[var(--ink-soft)]" size={30} />
          <h2 className="font-semibold">ยังไม่มีคำตอบ</h2>
          <p className="mt-1 text-sm text-[var(--ink-soft)]">
            เปิดลิงก์แบบฟอร์มที่แชร์เพื่อเริ่มกรอกคำตอบ
          </p>
        </div>
      </Card>
    );
  } else {
    responseContent = (
      <div className="overflow-hidden rounded-[var(--radius)] border border-[var(--line)] bg-[var(--paper)]">
        <div className="divide-y divide-[var(--line)]">
          {rows.map((row) => {
            const submitted =
              row.status === "submitted" || row.status === "completed";
            return (
              <div
                className="flex flex-col gap-4 p-5 sm:flex-row sm:items-center sm:justify-between"
                key={row.id}
              >
                <div>
                  <h2 className="font-semibold">
                    {row.formTitle ?? "แบบฟอร์มไม่มีชื่อ"}
                  </h2>
                  <p className="mt-1 text-sm text-[var(--ink-soft)]">
                    {submitted
                      ? `ส่งเมื่อ ${formatDate(row.submittedAt ?? row.createdAt)}`
                      : `บันทึกล่าสุด ${formatDate(row.updatedAt ?? row.createdAt)}`}
                    {submitted && row.latestCorrectionNumber
                      ? ` · แก้ไขครั้งที่ ${row.latestCorrectionNumber}`
                      : ""}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <Badge tone={submitted ? "success" : "warning"}>
                    {submitted ? "ส่งแล้ว" : "ฉบับร่าง"}
                  </Badge>
                  {submitted ? (
                    <Link
                      to="/receipt/$submissionId"
                      params={{ submissionId: row.submissionId ?? row.id }}
                      className="inline-flex items-center gap-1 text-sm font-semibold underline decoration-[var(--line-strong)] underline-offset-4 hover:decoration-[var(--ink)]"
                    >
                      ดูใบรับคำตอบ <ArrowUpRight size={15} />
                    </Link>
                  ) : (
                    <Link
                      to="/forms/$publicId/fill"
                      params={{
                        publicId: row.formPublicId ?? "",
                      }}
                      search={{ responseId: row.id }}
                      className="inline-flex items-center gap-1 text-sm font-semibold text-[var(--success)]"
                    >
                      กลับไปกรอกต่อ <ArrowUpRight size={15} />
                    </Link>
                  )}
                  {submitted ? null : (
                    <Button
                      type="button"
                      variant="ghost"
                      size="sm"
                      onClick={(event) =>
                        requestDiscard(row.id, event.currentTarget)
                      }
                      disabled={discardingId === row.id}
                    >
                      {discardingId === row.id ? <Spinner /> : null}
                      ทิ้งฉบับร่าง
                    </Button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  const discardCandidate = rows.find((row) => row.id === discardCandidateId);
  return (
    <AppShell>
      {discardCandidate ? (
        <div
          className="fixed inset-0 z-50 grid place-items-center bg-black/35 px-5"
          role="presentation"
        >
          <div
            className="w-full max-w-lg rounded-[var(--radius)] border border-[var(--line)] bg-[var(--paper)] p-6 shadow-xl"
            role="dialog"
            aria-modal="true"
            aria-labelledby="discard-title"
          >
            <h2 id="discard-title" className="text-lg font-semibold">
              ทิ้งฉบับร่างหรือไม่
            </h2>
            <p className="mt-2 text-sm text-[var(--ink-soft)]">
              ฉบับร่างของ {discardCandidate.formTitle ?? "แบบฟอร์มนี้"} จะถูกลบ
              พร้อมเอกสารที่บันทึกไว้และไม่สามารถกู้คืนได้
            </p>
            {discardError ? (
              <div className="mt-4">
                <Notice tone="danger">{discardError}</Notice>
              </div>
            ) : null}
            <div className="mt-6 flex justify-end gap-2">
              <Button
                type="button"
                variant="ghost"
                onClick={cancelDiscard}
                id="discard-cancel"
                disabled={discardingId !== null}
              >
                ยกเลิก
              </Button>
              <Button
                type="button"
                variant="danger"
                onClick={discardDraft}
                disabled={discardingId !== null}
              >
                {discardingId ? <Spinner /> : null}
                {discardingId ? "กำลังลบ…" : "ยืนยันการทิ้ง"}
              </Button>
            </div>
          </div>
        </div>
      ) : null}
      {discardSuccess ? (
        <div
          ref={discardStatusRef}
          className="mb-4"
          role="status"
          aria-live="polite"
          tabIndex={-1}
        >
          <Notice tone="success">{discardSuccess}</Notice>
        </div>
      ) : null}
      <PageHeader
        title="คำตอบของฉัน"
        description="กลับไปกรอกฉบับร่างที่บันทึกไว้ หรือดูคำตอบที่ส่งแล้ว"
      />
      <div className="mb-8 grid gap-4 sm:grid-cols-3">
        <Stat
          label="กำลังดำเนินการ"
          value={
            rows.filter(
              (row) => row.status === "draft" || row.status === "in_progress"
            ).length
          }
          icon={<ClipboardList size={18} />}
        />
        <Stat
          label="ส่งแล้ว"
          value={
            rows.filter(
              (row) => row.status === "submitted" || row.status === "completed"
            ).length
          }
          icon={<FileCheck2 size={18} />}
        />
        <Stat
          label="คำตอบทั้งหมด"
          value={rows.length}
          icon={<Inbox size={18} />}
        />
      </div>
      {responseContent}
    </AppShell>
  );
};

export const Route = createFileRoute("/dashboard")({
  component: DashboardRoute,
});
