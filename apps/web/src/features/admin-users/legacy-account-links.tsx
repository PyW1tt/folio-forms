import { ChevronRight, ShieldCheck, X } from "lucide-react";
import type { RefObject } from "react";

import { Badge, Button, Card, Notice, Spinner } from "@/components/ui";
import { formatDate } from "@/lib/api";
import type { AdminLegacyAccountLinkRequest } from "@/lib/api";

import { roleBadgeLabels, roleBadgeTones } from "./model";

export const LegacyAccountLinks = ({
  requests,
  nextCursor,
  loading,
  error,
  feedback,
  pendingAction,
  headingRef,
  reload,
  review,
  loadMore,
}: {
  requests: AdminLegacyAccountLinkRequest[];
  nextCursor: string | null;
  loading: boolean;
  error: string | null;
  feedback: string | null;
  pendingAction: string | null;
  headingRef: RefObject<HTMLHeadingElement | null>;
  reload: () => void;
  review: (
    request: AdminLegacyAccountLinkRequest,
    decision: "approve" | "reject"
  ) => Promise<void>;
  loadMore: () => Promise<void>;
}) => (
  <Card className="mt-6 overflow-hidden" aria-busy={loading}>
    <div className="border-b border-[var(--line)] px-5 py-4 sm:px-6">
      <div className="flex items-start gap-3">
        <span className="grid size-10 shrink-0 place-items-center rounded-[10px] bg-[var(--accent-soft)]">
          <ShieldCheck size={19} />
        </span>
        <div>
          <h2
            ref={headingRef}
            tabIndex={-1}
            className="text-xl font-bold tracking-[-0.03em]"
          >
            คำขอเชื่อมบัญชี Legacy
          </h2>
          <p className="mt-1 text-sm text-[var(--ink-soft)]">
            ตรวจข้อมูลตัวตนจาก PDMS และบัญชี Folio Forms ก่อนอนุมัติ PDMS
            ไม่ได้ยืนยันความเป็นเจ้าของกล่องอีเมล ผู้ใช้ต้องเริ่มเข้าสู่ระบบผ่านระบบเดิมใหม่หลังอนุมัติ
          </p>
        </div>
      </div>
    </div>
    {feedback ? (
      <div className="px-5 pt-4 sm:px-6">
        <Notice tone="success">{feedback}</Notice>
      </div>
    ) : null}
    {error ? (
      <div className="px-5 pt-4 sm:px-6">
        <Notice tone="danger">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <span>{error}</span>
            <Button
              type="button"
              variant="secondary"
              size="sm"
              onClick={reload}
              disabled={loading}
            >
              ลองใหม่
            </Button>
          </div>
        </Notice>
      </div>
    ) : null}
    {loading && requests.length === 0 ? (
      <div className="grid min-h-32 place-items-center gap-3 p-6 text-sm text-[var(--ink-soft)]">
        <Spinner />
        <span>กำลังโหลดคำขอ…</span>
      </div>
    ) : null}
    {!loading && requests.length === 0 && error === null ? (
      <p className="p-6 text-center text-sm text-[var(--ink-soft)]">
        ไม่มีคำขอที่รออนุมัติ
      </p>
    ) : null}
    {requests.length > 0 ? (
      <ul className="divide-y divide-[var(--line)]">
        {requests.map((request) => {
          let ineligibleReason: string | null = null;
          if (!request.user.enabled) {
            ineligibleReason = "บัญชี Folio Forms ปิดใช้งานอยู่";
          } else if (request.user.role !== "user") {
            ineligibleReason = "ไม่สามารถเชื่อมบัญชีกับบัญชีผู้ดูแลระบบ";
          } else if (request.user.email !== request.email) {
            ineligibleReason = "อีเมลบัญชี Folio Forms เปลี่ยนไปแล้ว";
          }
          const actionIsPending =
            pendingAction?.startsWith(`${request.id}:`) === true;
          const actionDisabled = loading || pendingAction !== null;
          return (
            <li key={request.id} className="space-y-4 p-5 sm:px-6">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div className="min-w-0">
                  <h3 className="font-semibold text-[var(--ink)]">
                    {request.user.name}
                  </h3>
                  <p className="break-all text-sm text-[var(--ink-soft)]">
                    {request.user.email}
                  </p>
                </div>
                <div className="flex flex-wrap gap-2">
                  <Badge tone={roleBadgeTones[request.user.role]}>
                    {roleBadgeLabels[request.user.role]}
                  </Badge>
                  <Badge tone={request.user.enabled ? "neutral" : "warning"}>
                    {request.user.enabled ? "เปิดใช้งาน" : "ปิดใช้งาน"}
                  </Badge>
                </div>
              </div>
              <dl className="grid gap-3 text-sm sm:grid-cols-2">
                <div>
                  <dt className="text-xs text-[var(--ink-soft)]">
                    อีเมลบัญชีจาก PDMS
                  </dt>
                  <dd className="break-all font-medium">{request.email}</dd>
                </div>
                <div>
                  <dt className="text-xs text-[var(--ink-soft)]">วันที่ส่งคำขอ</dt>
                  <dd>{formatDate(request.createdAt)}</dd>
                </div>
                <div>
                  <dt className="text-xs text-[var(--ink-soft)]">Provider</dt>
                  <dd className="break-all">{request.providerId}</dd>
                </div>
                <div>
                  <dt className="text-xs text-[var(--ink-soft)]">Subject</dt>
                  <dd className="break-all">{request.subject}</dd>
                </div>
              </dl>
              {ineligibleReason ? (
                <p className="text-sm text-[var(--danger)]">
                  {ineligibleReason}
                </p>
              ) : null}
              <div className="flex flex-wrap gap-2">
                <Button
                  type="button"
                  size="sm"
                  onClick={async () => {
                    await review(request, "approve");
                  }}
                  disabled={Boolean(ineligibleReason) || actionDisabled}
                  aria-label={`อนุมัติการเชื่อมบัญชี ${request.email} กับ ${request.user.email}`}
                >
                  {actionIsPending &&
                  pendingAction === `${request.id}:approve` ? (
                    <Spinner />
                  ) : (
                    <ShieldCheck size={15} />
                  )}
                  อนุมัติ
                </Button>
                <Button
                  type="button"
                  variant="danger"
                  size="sm"
                  onClick={async () => {
                    await review(request, "reject");
                  }}
                  disabled={actionDisabled}
                  aria-label={`ปฏิเสธคำขอเชื่อมบัญชี ${request.email}`}
                >
                  {actionIsPending &&
                  pendingAction === `${request.id}:reject` ? (
                    <Spinner />
                  ) : (
                    <X size={15} />
                  )}
                  ปฏิเสธ
                </Button>
              </div>
            </li>
          );
        })}
      </ul>
    ) : null}
    {nextCursor ? (
      <div className="flex justify-center border-t border-[var(--line)] p-4">
        <Button
          type="button"
          variant="secondary"
          size="sm"
          onClick={loadMore}
          disabled={loading || pendingAction !== null}
        >
          {loading ? <Spinner /> : <ChevronRight size={15} />}
          โหลดคำขอเพิ่มเติม
        </Button>
      </div>
    ) : null}
  </Card>
);
