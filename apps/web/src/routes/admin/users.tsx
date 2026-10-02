import {
  createFileRoute,
  useBlocker,
  useNavigate,
} from "@tanstack/react-router";
import {
  ChevronLeft,
  ChevronRight,
  Copy,
  KeyRound,
  Mail,
  Pencil,
  Plus,
  RefreshCw,
  Search,
  ShieldCheck,
  Trash2,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import type { FormEvent, RefObject } from "react";
import { useEffect, useRef, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import { Badge, Button, Card, Input, Notice, Spinner } from "@/components/ui";
import {
  ApiError,
  apiDelete,
  apiGet,
  apiPatch,
  apiPost,
  formatDate,
} from "@/lib/api";
import type {
  AdminLegacyAccountLinkListResponse,
  AdminLegacyAccountLinkMutationResponse,
  AdminLegacyAccountLinkRequest,
  AdminUser,
  AdminUserCredentialResponse,
  AdminUserListResponse,
  AdminUserMutationResponse,
  Role,
} from "@/lib/api";
import { useAuth } from "@/lib/auth";

type FilterRole = Role | "all";
type FilterEnabled = "all" | "true" | "false";
type ConfirmKind =
  | "delete"
  | "disable"
  | "email"
  | "promote"
  | "demote"
  | "reset";
type ActionKind = ConfirmKind | "enable";

interface Confirmation {
  user: AdminUser;
  kind: ConfirmKind;
  email?: string;
}

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+$/u;
const selectClass =
  "min-h-11 w-full rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-[var(--ink)] shadow-sm focus:border-[var(--ink)] focus:outline-none";

const normalizeEmail = (value: string) => value.trim().toLowerCase();

const actionsAreDisabled = (
  loading: boolean,
  creating: boolean,
  pendingAction: string | null,
  confirmation: Confirmation | null
): boolean =>
  loading || creating || pendingAction !== null || confirmation !== null;

const resetIsDisabled = (
  actionsDisabled: boolean,
  targetUserId: string,
  authenticatedUserId: string | undefined
): boolean => actionsDisabled || targetUserId === authenticatedUserId;
const credentialRequestIsInFlight = (
  creating: boolean,
  pendingAction: string | null
): boolean => creating || pendingAction?.endsWith(":reset") === true;
const actionKey = (userId: string, kind: ActionKind) => `${userId}:${kind}`;

const roleBadgeLabels: Record<Role, string> = {
  admin: "ผู้ดูแลระบบ",
  user: "ผู้ใช้",
};
const roleBadgeTones: Record<Role, "neutral" | "warning"> = {
  admin: "warning",
  user: "neutral",
};
const errorMessageFor = (caughtError: unknown, fallback: string) => {
  if (caughtError instanceof ApiError) {
    if (caughtError.code === "email_in_use") {
      return "อีเมลนี้ถูกใช้งานแล้ว กรุณาใช้อีเมลอื่น";
    }
    if (caughtError.code === "final_admin_required") {
      return "ต้องมีผู้ดูแลระบบที่เปิดใช้งานอยู่อย่างน้อยหนึ่งบัญชี";
    }
    if (caughtError.code === "personal_data_remains") {
      return "ยังมีข้อมูลส่วนบุคคลหรือไฟล์ของคำตอบค้างอยู่ ต้องลบคำตอบก่อน";
    }
    if (caughtError.code === "user_not_found" || caughtError.status === 404) {
      return "ไม่พบบัญชีผู้ใช้นี้ อาจถูกลบไปแล้ว กรุณาโหลดรายการใหม่";
    }
    if (caughtError.status === 400 || caughtError.code === "invalid_input") {
      return "ข้อมูลไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง";
    }
    if (caughtError.status === 401) {
      return "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่";
    }
    if (caughtError.status === 403) {
      return "คุณไม่มีสิทธิ์ดำเนินการนี้";
    }
  }
  return fallback;
};

const confirmationTitle = (kind: ConfirmKind) => {
  if (kind === "delete") {
    return "ยืนยันการลบบัญชีถาวร";
  }
  if (kind === "disable") {
    return "ยืนยันการปิดใช้งานบัญชี";
  }
  if (kind === "email") {
    return "ยืนยันการเปลี่ยนอีเมล";
  }
  if (kind === "promote") {
    return "ยืนยันการเลื่อนเป็นผู้ดูแลระบบ";
  }
  if (kind === "demote") {
    return "ยืนยันการลดสิทธิ์ผู้ดูแลระบบ";
  }
  return "ยืนยันการตั้งรหัสผ่านใหม่";
};

const confirmationDescription = (kind: ConfirmKind) => {
  if (kind === "delete") {
    return "ระบบจะลบข้อมูลบัญชีและข้อมูลส่วนบุคคลที่เหลืออยู่ถาวร ผู้ใช้ที่มี Response ต้องลบ Response ก่อน";
  }
  if (kind === "disable") {
    return "การปิดใช้งานจะออกจากระบบของบัญชีนี้ทุกอุปกรณ์";
  }
  if (kind === "email") {
    return "การเปลี่ยนอีเมลจะออกจากระบบของบัญชีนี้ทุกอุปกรณ์";
  }
  if (kind === "promote" || kind === "demote") {
    return "การเปลี่ยนบทบาทจะออกจากระบบของบัญชีนี้ทุกอุปกรณ์";
  }
  return "ระบบจะออกจากระบบของบัญชีนี้ทุกอุปกรณ์และสร้างรหัสผ่านชั่วคราวใหม่";
};

const actionSuccessMessage = (kind: ActionKind) => {
  if (kind === "delete") {
    return "ลบบัญชีและข้อมูลส่วนบุคคลถาวรแล้ว";
  }
  if (kind === "enable") {
    return "เปิดใช้งานบัญชีแล้ว";
  }
  if (kind === "disable") {
    return "ปิดใช้งานบัญชีแล้ว และยกเลิกเซสชันทั้งหมดแล้ว";
  }
  if (kind === "email") {
    return "เปลี่ยนอีเมลแล้ว และยกเลิกเซสชันทั้งหมดแล้ว";
  }
  if (kind === "promote") {
    return "เลื่อนบัญชีเป็นผู้ดูแลระบบแล้ว และยกเลิกเซสชันทั้งหมดแล้ว";
  }
  if (kind === "demote") {
    return "ลดสิทธิ์บัญชีแล้ว และยกเลิกเซสชันทั้งหมดแล้ว";
  }
  return "ตั้งรหัสผ่านใหม่แล้ว และยกเลิกเซสชันทั้งหมดแล้ว";
};

const actionFailureMessage = (kind: ActionKind) => {
  if (kind === "delete") {
    return "ไม่สามารถลบบัญชีและข้อมูลส่วนบุคคลได้ กรุณาตรวจสอบเงื่อนไขแล้วลองใหม่";
  }
  if (kind === "enable") {
    return "ไม่สามารถเปิดใช้งานบัญชีได้ กรุณาลองใหม่อีกครั้ง";
  }
  if (kind === "disable") {
    return "ไม่สามารถปิดใช้งานบัญชีได้ กรุณาลองใหม่อีกครั้ง";
  }
  if (kind === "email") {
    return "ไม่สามารถเปลี่ยนอีเมลได้ กรุณาลองใหม่อีกครั้ง";
  }
  if (kind === "promote") {
    return "ไม่สามารถเลื่อนบัญชีได้ กรุณาลองใหม่อีกครั้ง";
  }
  if (kind === "demote") {
    return "ไม่สามารถลดสิทธิ์บัญชีได้ กรุณาลองใหม่อีกครั้ง";
  }
  return "ไม่สามารถตั้งรหัสผ่านใหม่ได้ กรุณาลองใหม่อีกครั้ง";
};

const TemporaryPasswordNotice = ({
  copyTemporaryPassword,
  temporaryPassword,
  temporaryPasswordCopyFeedback,
  temporaryPasswordRef,
}: {
  copyTemporaryPassword: () => Promise<void>;
  temporaryPassword: string | null;
  temporaryPasswordCopyFeedback: {
    tone: "danger" | "success";
    message: string;
  } | null;
  temporaryPasswordRef: RefObject<HTMLDivElement | null>;
}) => {
  if (!temporaryPassword) {
    return null;
  }
  return (
    <>
      <div
        ref={temporaryPasswordRef}
        tabIndex={-1}
        role="status"
        aria-live="polite"
        aria-atomic="true"
        className="mt-4 rounded-[10px] border border-[var(--success)]/25 bg-[var(--success-soft)] px-4 py-3 text-sm text-[var(--success)] focus:outline-none"
      >
        <strong className="block">รหัสผ่านชั่วคราว (แสดงครั้งเดียว)</strong>
        <code className="mt-2 block break-all rounded-md bg-[var(--paper)] px-3 py-2 text-base font-semibold text-[var(--ink)]">
          {temporaryPassword}
        </code>
        <span className="mt-2 block text-xs">
          จดหรือส่งรหัสนี้ให้เจ้าของบัญชีอย่างปลอดภัย
          ระบบจะไม่แสดงรหัสนี้อีกหลังจากการดำเนินการครั้งถัดไป
        </span>
        <Button
          className="mt-3"
          onClick={copyTemporaryPassword}
          size="sm"
          type="button"
          variant="secondary"
        >
          <Copy aria-hidden="true" size={15} />
          คัดลอกรหัสผ่านชั่วคราว
        </Button>
      </div>
      <p
        aria-atomic="true"
        aria-live="polite"
        className={
          temporaryPasswordCopyFeedback
            ? `mt-2 text-sm font-medium ${
                temporaryPasswordCopyFeedback.tone === "danger"
                  ? "text-[var(--danger)]"
                  : "text-[var(--success)]"
              }`
            : "sr-only"
        }
        role="status"
      >
        {temporaryPasswordCopyFeedback?.message ?? ""}
      </p>
    </>
  );
};

const LegacyAccountLinks = ({
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

const AdminUsersRoute = () => {
  const navigate = useNavigate();
  const { clearSession, user: authenticatedUser } = useAuth();
  const [users, setUsers] = useState<AdminUser[]>([]);
  const [listLoading, setListLoading] = useState(true);
  const [listError, setListError] = useState<string | null>(null);
  const [nextCursor, setNextCursor] = useState<string | null>(null);
  const [currentCursor, setCurrentCursor] = useState<string | null>(null);
  const [cursorHistory, setCursorHistory] = useState<(string | null)[]>([]);
  const [reloadVersion, setReloadVersion] = useState(0);
  const [legacyAccountLinkRequests, setLegacyAccountLinkRequests] = useState<
    AdminLegacyAccountLinkRequest[]
  >([]);
  const [legacyAccountLinkNextCursor, setLegacyAccountLinkNextCursor] =
    useState<string | null>(null);
  const [legacyAccountLinkLoading, setLegacyAccountLinkLoading] =
    useState(true);
  const [legacyAccountLinkError, setLegacyAccountLinkError] = useState<
    string | null
  >(null);
  const [legacyAccountLinkFeedback, setLegacyAccountLinkFeedback] = useState<
    string | null
  >(null);
  const [legacyAccountLinkAction, setLegacyAccountLinkAction] = useState<
    string | null
  >(null);
  const legacyAccountLinksHeadingRef = useRef<HTMLHeadingElement | null>(null);

  const [draftEmailFilter, setDraftEmailFilter] = useState("");
  const [draftRoleFilter, setDraftRoleFilter] = useState<FilterRole>("all");
  const [draftEnabledFilter, setDraftEnabledFilter] =
    useState<FilterEnabled>("all");
  const [emailFilter, setEmailFilter] = useState("");
  const [roleFilter, setRoleFilter] = useState<FilterRole>("all");
  const [enabledFilter, setEnabledFilter] = useState<FilterEnabled>("all");

  const [createName, setCreateName] = useState("");
  const [createEmail, setCreateEmail] = useState("");
  const [createRole, setCreateRole] = useState<Role>("user");
  const [createError, setCreateError] = useState<string | null>(null);
  const [creating, setCreating] = useState(false);

  const [feedback, setFeedback] = useState<{
    tone: "danger" | "success";
    message: string;
  } | null>(null);
  const [temporaryPassword, setTemporaryPassword] = useState<string | null>(
    null
  );
  const [temporaryPasswordCopyFeedback, setTemporaryPasswordCopyFeedback] =
    useState<{
      tone: "danger" | "success";
      message: string;
    } | null>(null);
  const [pendingAction, setPendingAction] = useState<string | null>(null);
  const [confirmingAction, setConfirmingAction] = useState<Confirmation | null>(
    null
  );
  const [editingEmailId, setEditingEmailId] = useState<string | null>(null);
  const [editingEmail, setEditingEmail] = useState("");
  const [emailEditError, setEmailEditError] = useState<string | null>(null);
  const credentialRequestInFlight = credentialRequestIsInFlight(
    creating,
    pendingAction
  );
  useBlocker({
    disabled: credentialRequestInFlight === false,
    enableBeforeUnload: credentialRequestInFlight,
    shouldBlockFn: () => credentialRequestInFlight,
  });

  const rowActionsDisabled = actionsAreDisabled(
    listLoading,
    creating,
    pendingAction,
    confirmingAction
  );
  const temporaryPasswordRef = useRef<HTMLDivElement | null>(null);
  const temporaryPasswordCopyVersion = useRef(0);
  const usersHeadingRef = useRef<HTMLHeadingElement | null>(null);
  const focusListAfterLoadRef = useRef(false);
  const restoreFocusKeyRef = useRef<string | null>(null);

  const clearTransientState = () => {
    temporaryPasswordCopyVersion.current += 1;
    setFeedback(null);
    setTemporaryPassword(null);
    setTemporaryPasswordCopyFeedback(null);
    setEmailEditError(null);
  };
  const copyTemporaryPassword = async () => {
    if (!temporaryPassword) {
      return;
    }

    const copyVersion = temporaryPasswordCopyVersion.current;
    try {
      await navigator.clipboard.writeText(temporaryPassword);
      if (copyVersion !== temporaryPasswordCopyVersion.current) {
        return;
      }
      setTemporaryPasswordCopyFeedback({
        message: "คัดลอกรหัสผ่านชั่วคราวแล้ว",
        tone: "success",
      });
    } catch {
      if (copyVersion !== temporaryPasswordCopyVersion.current) {
        return;
      }
      setTemporaryPasswordCopyFeedback({
        message: "คัดลอกรหัสผ่านชั่วคราวไม่สำเร็จ กรุณาลองใหม่",
        tone: "danger",
      });
    }
  };
  const leaveAfterOwnMutation = async (targetId: string): Promise<boolean> => {
    if (authenticatedUser?.id !== targetId) {
      return false;
    }
    clearSession();
    await navigate({
      replace: true,
      search: { returnTo: "/admin/users" },
      to: "/login",
    });
    return true;
  };

  useEffect(() => {
    let cancelled = false;
    const loadUsers = async () => {
      setListLoading(true);
      setListError(null);
      const query = new URLSearchParams();
      if (currentCursor) {
        query.set("cursor", currentCursor);
      }
      if (emailFilter) {
        query.set("email", emailFilter);
      }
      if (roleFilter !== "all") {
        query.set("role", roleFilter);
      }
      if (enabledFilter !== "all") {
        query.set("enabled", enabledFilter);
      }

      try {
        const payload = await apiGet<AdminUserListResponse>(
          `/api/admin/users${query.toString() ? `?${query.toString()}` : ""}`
        );
        if (cancelled) {
          return;
        }
        setUsers(payload.users);
        setNextCursor(payload.nextCursor);
      } catch (caughtError) {
        if (cancelled) {
          return;
        }
        setUsers([]);
        setNextCursor(null);
        setListError(
          errorMessageFor(
            caughtError,
            "ไม่สามารถโหลดรายการผู้ใช้ได้ กรุณาลองใหม่อีกครั้ง"
          )
        );
      } finally {
        if (!cancelled) {
          setListLoading(false);
          if (focusListAfterLoadRef.current) {
            focusListAfterLoadRef.current = false;
            window.requestAnimationFrame(() => {
              usersHeadingRef.current?.focus();
            });
          }
        }
      }
    };

    void loadUsers();
    return () => {
      cancelled = true;
    };
  }, [currentCursor, emailFilter, enabledFilter, reloadVersion, roleFilter]);
  useEffect(() => {
    let cancelled = false;
    const loadLegacyAccountLinks = async () => {
      setLegacyAccountLinkLoading(true);
      setLegacyAccountLinkError(null);
      try {
        const payload = await apiGet<AdminLegacyAccountLinkListResponse>(
          "/api/admin/account-links"
        );
        if (cancelled) {
          return;
        }
        setLegacyAccountLinkRequests(payload.requests);
        setLegacyAccountLinkNextCursor(payload.nextCursor);
      } catch (caughtError) {
        if (cancelled) {
          return;
        }
        setLegacyAccountLinkRequests([]);
        setLegacyAccountLinkNextCursor(null);
        setLegacyAccountLinkError(
          errorMessageFor(
            caughtError,
            "ไม่สามารถโหลดคำขอเชื่อมบัญชีได้ กรุณาลองใหม่อีกครั้ง"
          )
        );
      } finally {
        if (!cancelled) {
          setLegacyAccountLinkLoading(false);
        }
      }
    };

    void loadLegacyAccountLinks();
    return () => {
      cancelled = true;
    };
  }, [reloadVersion]);

  useEffect(() => {
    if (confirmingAction) {
      const key = actionKey(confirmingAction.user.id, confirmingAction.kind);
      window.requestAnimationFrame(() => {
        document
          .querySelector<HTMLButtonElement>(
            `button[data-confirm-action-id="${CSS.escape(key)}"]`
          )
          ?.focus();
      });
      return;
    }

    if (editingEmailId) {
      window.requestAnimationFrame(() => {
        document
          .querySelector<HTMLInputElement>(
            `input[data-edit-email-id="${CSS.escape(editingEmailId)}"]`
          )
          ?.focus();
      });
      return;
    }

    if (pendingAction) {
      return;
    }
    const restoreFocusKey = restoreFocusKeyRef.current;
    if (restoreFocusKey) {
      window.requestAnimationFrame(() => {
        document
          .querySelector<HTMLButtonElement>(
            `button[data-user-action-id="${CSS.escape(restoreFocusKey)}"]`
          )
          ?.focus();
      });
      restoreFocusKeyRef.current = null;
    }
  }, [confirmingAction, editingEmailId, pendingAction]);

  useEffect(() => {
    if (temporaryPassword) {
      window.requestAnimationFrame(() => {
        temporaryPasswordRef.current?.focus();
      });
    }
  }, [temporaryPassword]);

  const applyFilters = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (rowActionsDisabled) {
      return;
    }
    clearTransientState();
    focusListAfterLoadRef.current = true;
    setEmailFilter(normalizeEmail(draftEmailFilter));
    setRoleFilter(draftRoleFilter);
    setEnabledFilter(draftEnabledFilter);
    setCurrentCursor(null);
    setCursorHistory([]);
    setReloadVersion((value) => value + 1);
  };

  const resetFilters = () => {
    clearTransientState();
    focusListAfterLoadRef.current = true;
    setDraftEmailFilter("");
    setDraftRoleFilter("all");
    setDraftEnabledFilter("all");
    setEmailFilter("");
    setRoleFilter("all");
    setEnabledFilter("all");
    setCurrentCursor(null);
    setCursorHistory([]);
    setReloadVersion((value) => value + 1);
  };

  const goToNextPage = () => {
    if (!nextCursor || rowActionsDisabled) {
      return;
    }
    clearTransientState();
    focusListAfterLoadRef.current = true;
    setCursorHistory((history) => [...history, currentCursor]);
    setCurrentCursor(nextCursor);
  };

  const goToPreviousPage = () => {
    if (cursorHistory.length === 0 || rowActionsDisabled) {
      return;
    }
    clearTransientState();
    focusListAfterLoadRef.current = true;
    const previousCursor = cursorHistory.at(-1) ?? null;
    setCursorHistory((history) => history.slice(0, -1));
    setCurrentCursor(previousCursor);
  };

  const createUser = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (rowActionsDisabled) {
      return;
    }
    clearTransientState();
    const name = createName.trim();
    const email = normalizeEmail(createEmail);
    if (!name || !EMAIL_PATTERN.test(email)) {
      setCreateError("กรุณากรอกชื่อและอีเมลให้ถูกต้อง");
      return;
    }

    setCreating(true);
    setCreateError(null);
    try {
      const payload = await apiPost<AdminUserCredentialResponse>(
        "/api/admin/users",
        { email, name, role: createRole }
      );
      setCreateName("");
      setCreateEmail("");
      setCreateRole("user");
      setCurrentCursor(null);
      setCursorHistory([]);
      setReloadVersion((value) => value + 1);
      setFeedback({
        message: `สร้างบัญชี ${payload.user.email} แล้ว กรุณาส่งรหัสผ่านชั่วคราวให้เจ้าของบัญชีอย่างปลอดภัย`,
        tone: "success",
      });
      temporaryPasswordCopyVersion.current += 1;
      setTemporaryPassword(payload.temporaryPassword);
    } catch (caughtError) {
      setCreateError(
        errorMessageFor(
          caughtError,
          "ไม่สามารถสร้างบัญชีได้ กรุณาตรวจสอบข้อมูลแล้วลองใหม่อีกครั้ง"
        )
      );
    } finally {
      setCreating(false);
    }
  };

  const requestConfirmation = (user: AdminUser, kind: ConfirmKind) => {
    if (creating || listLoading || pendingAction || confirmingAction) {
      return;
    }
    clearTransientState();
    let email: string | undefined;
    if (kind === "email") {
      email = normalizeEmail(editingEmail);
      if (!EMAIL_PATTERN.test(email)) {
        setEmailEditError("กรุณากรอกอีเมลที่ถูกต้อง");
        return;
      }
      if (email === user.email) {
        setEmailEditError("อีเมลนี้เป็นอีเมลปัจจุบันอยู่แล้ว");
        return;
      }
    }

    restoreFocusKeyRef.current = actionKey(user.id, kind);
    setConfirmingAction({ email, kind, user });
  };

  const mutateUser = async (
    user: AdminUser,
    kind: ActionKind,
    nextEmail?: string
  ) => {
    if (creating || listLoading || pendingAction) {
      return;
    }

    const key = actionKey(user.id, kind);
    setPendingAction(key);
    clearTransientState();
    let succeeded = false;
    try {
      if (kind === "reset") {
        const payload = await apiPost<AdminUserCredentialResponse>(
          `/api/admin/users/${encodeURIComponent(user.id)}/password-reset`
        );
        setUsers((currentUsers) =>
          currentUsers.map((currentUser) =>
            currentUser.id === payload.user.id ? payload.user : currentUser
          )
        );
        setFeedback({
          message: `ตั้งรหัสผ่านใหม่สำหรับ ${payload.user.email} แล้ว กรุณาส่งรหัสผ่านชั่วคราวให้เจ้าของบัญชีอย่างปลอดภัย`,
          tone: "success",
        });
        temporaryPasswordCopyVersion.current += 1;
        setTemporaryPassword(payload.temporaryPassword);
        setReloadVersion((value) => value + 1);
        succeeded = true;
        return;
      }
      if (kind === "delete") {
        await apiDelete(`/api/admin/users/${encodeURIComponent(user.id)}`, {
          confirm: true,
        });
        if (authenticatedUser?.id === user.id) {
          clearSession();
          await navigate({
            replace: true,
            search: { returnTo: "/admin/users" },
            to: "/login",
          });
          return;
        }
        setUsers((currentUsers) =>
          currentUsers.filter((currentUser) => currentUser.id !== user.id)
        );
        setFeedback({
          message: actionSuccessMessage(kind),
          tone: "success",
        });
        succeeded = true;
        return;
      }

      let body: { enabled: boolean } | { email: string } | { role: Role };
      if (kind === "enable" || kind === "disable") {
        body = { enabled: kind === "enable" };
      } else if (kind === "email") {
        const email = normalizeEmail(nextEmail ?? editingEmail);
        if (!EMAIL_PATTERN.test(email)) {
          setEmailEditError("กรุณากรอกอีเมลที่ถูกต้อง");
          return;
        }
        body = { email };
      } else {
        body = { role: kind === "promote" ? "admin" : "user" };
      }

      const payload = await apiPatch<AdminUserMutationResponse>(
        `/api/admin/users/${encodeURIComponent(user.id)}`,
        body
      );
      if (await leaveAfterOwnMutation(payload.user.id)) {
        return;
      }
      setUsers((currentUsers) =>
        currentUsers.map((currentUser) =>
          currentUser.id === payload.user.id ? payload.user : currentUser
        )
      );
      setFeedback({ message: actionSuccessMessage(kind), tone: "success" });
      focusListAfterLoadRef.current = true;
      setReloadVersion((value) => value + 1);
      succeeded = true;
    } catch (caughtError) {
      setFeedback({
        message: errorMessageFor(caughtError, actionFailureMessage(kind)),
        tone: "danger",
      });
    } finally {
      setPendingAction(null);
      setConfirmingAction(null);
      if (succeeded && kind === "email") {
        setEditingEmailId(null);
        setEmailEditError(null);
      }
    }
  };

  const beginEmailEdit = (user: AdminUser) => {
    if (creating || listLoading || pendingAction || confirmingAction) {
      return;
    }
    clearTransientState();
    setEditingEmailId(user.id);
    setEditingEmail(user.email);
  };

  const cancelEmailEdit = (user: AdminUser) => {
    restoreFocusKeyRef.current = actionKey(user.id, "email");
    setEditingEmailId(null);
    setEmailEditError(null);
  };

  const submitEmailEdit = (
    event: FormEvent<HTMLFormElement>,
    user: AdminUser
  ) => {
    event.preventDefault();
    requestConfirmation(user, "email");
  };

  const loadMoreLegacyAccountLinks = async () => {
    const cursor = legacyAccountLinkNextCursor;
    if (!cursor || legacyAccountLinkLoading || legacyAccountLinkAction) {
      return;
    }
    setLegacyAccountLinkLoading(true);
    setLegacyAccountLinkError(null);
    try {
      const payload = await apiGet<AdminLegacyAccountLinkListResponse>(
        `/api/admin/account-links?cursor=${encodeURIComponent(cursor)}`
      );
      setLegacyAccountLinkRequests((current) => [
        ...current,
        ...payload.requests,
      ]);
      setLegacyAccountLinkNextCursor(payload.nextCursor);
    } catch (caughtError) {
      setLegacyAccountLinkError(
        errorMessageFor(caughtError, "ไม่สามารถโหลดคำขอเพิ่มเติมได้ กรุณาลองใหม่อีกครั้ง")
      );
    } finally {
      setLegacyAccountLinkLoading(false);
    }
  };

  const reviewLegacyAccountLink = async (
    request: AdminLegacyAccountLinkRequest,
    decision: "approve" | "reject"
  ) => {
    const action = `${request.id}:${decision}`;
    setLegacyAccountLinkAction(action);
    setLegacyAccountLinkError(null);
    setLegacyAccountLinkFeedback(null);
    try {
      await apiPost<AdminLegacyAccountLinkMutationResponse>(
        `/api/admin/account-links/${encodeURIComponent(request.id)}/${decision}`,
        { reviewedGeneration: request.reviewedGeneration }
      );
      setLegacyAccountLinkRequests((current) =>
        current.filter((currentRequest) => currentRequest.id !== request.id)
      );
      setLegacyAccountLinkFeedback(
        decision === "approve"
          ? "อนุมัติคำขอแล้ว ผู้ใช้ต้องเริ่มเข้าสู่ระบบผ่านระบบเดิมอีกครั้ง"
          : "ปฏิเสธคำขอเชื่อมบัญชีแล้ว"
      );
      window.requestAnimationFrame(() => {
        legacyAccountLinksHeadingRef.current?.focus();
      });
    } catch (caughtError) {
      if (caughtError instanceof ApiError && caughtError.status === 409) {
        setLegacyAccountLinkLoading(true);
        setReloadVersion((value) => value + 1);
      }
      setLegacyAccountLinkError(
        errorMessageFor(
          caughtError,
          "ไม่สามารถตรวจสอบคำขอเชื่อมบัญชีได้ กรุณาลองใหม่อีกครั้ง"
        )
      );
    } finally {
      setLegacyAccountLinkAction(null);
    }
  };

  return (
    <>
      <PageHeader
        title="จัดการผู้ใช้"
        description="สร้างบัญชี จัดการบทบาทและสถานะ พร้อมควบคุมเซสชันของผู้ใช้ในที่เดียว"
      />

      {feedback ? (
        <Notice tone={feedback.tone}>{feedback.message}</Notice>
      ) : null}

      <TemporaryPasswordNotice
        copyTemporaryPassword={copyTemporaryPassword}
        temporaryPassword={temporaryPassword}
        temporaryPasswordCopyFeedback={temporaryPasswordCopyFeedback}
        temporaryPasswordRef={temporaryPasswordRef}
      />

      <div className="mt-6 grid gap-6 lg:grid-cols-[minmax(0,0.8fr)_minmax(0,1.2fr)]">
        <Card className="p-5 sm:p-6">
          <div className="mb-5 flex items-start gap-3">
            <span className="grid size-10 shrink-0 place-items-center rounded-[10px] bg-[var(--accent-soft)]">
              <Plus size={19} />
            </span>
            <div>
              <h2 className="text-xl font-bold tracking-[-0.03em]">
                สร้างบัญชีใหม่
              </h2>
              <p className="mt-1 text-sm text-[var(--ink-soft)]">
                ระบบจะสร้างรหัสผ่านชั่วคราวและบังคับให้เปลี่ยนเมื่อเข้าสู่ระบบครั้งแรก
              </p>
            </div>
          </div>
          <div
            id="create-user-error"
            className="mb-4"
            hidden={createError === null}
          >
            <Notice tone="danger">{createError}</Notice>
          </div>
          <form
            className="space-y-4"
            onSubmit={createUser}
            noValidate
            aria-busy={creating}
          >
            <label
              className="block text-sm font-semibold"
              htmlFor="create-user-name"
            >
              ชื่อผู้ใช้
              <Input
                id="create-user-name"
                className="mt-2"
                value={createName}
                onChange={(event) => setCreateName(event.target.value)}
                aria-describedby="create-user-error"
                aria-invalid={Boolean(createError)}
                autoComplete="name"
                required
              />
            </label>
            <label
              className="block text-sm font-semibold"
              htmlFor="create-user-email"
            >
              อีเมล
              <Input
                id="create-user-email"
                className="mt-2"
                type="email"
                value={createEmail}
                onChange={(event) => setCreateEmail(event.target.value)}
                aria-describedby="create-user-error"
                aria-invalid={Boolean(createError)}
                autoComplete="email"
                required
              />
            </label>
            <label
              className="block text-sm font-semibold"
              htmlFor="create-user-role"
            >
              บทบาท
              <select
                id="create-user-role"
                className={`${selectClass} mt-2`}
                value={createRole}
                onChange={(event) => setCreateRole(event.target.value as Role)}
              >
                <option value="user">ผู้ใช้</option>
                <option value="admin">ผู้ดูแลระบบ</option>
              </select>
            </label>
            <div className="border-t border-[var(--line)] pt-4">
              <Button type="submit" disabled={rowActionsDisabled}>
                {creating ? <Spinner /> : <Plus size={16} />}
                {creating ? "กำลังสร้างบัญชี…" : "สร้างบัญชี"}
              </Button>
            </div>
          </form>
        </Card>

        <Card className="p-5 sm:p-6">
          <div className="mb-5 flex items-start gap-3">
            <span className="grid size-10 shrink-0 place-items-center rounded-[10px] bg-[var(--success-soft)] text-[var(--success)]">
              <Search size={19} />
            </span>
            <div>
              <h2 className="text-xl font-bold tracking-[-0.03em]">
                ค้นหาและกรองบัญชี
              </h2>
              <p className="mt-1 text-sm text-[var(--ink-soft)]">
                ค้นหาอีเมลแบบไม่สนใจตัวพิมพ์ใหญ่เล็ก และกรองตามบทบาทหรือสถานะ
              </p>
            </div>
          </div>
          <form className="grid gap-4 sm:grid-cols-2" onSubmit={applyFilters}>
            <label
              className="block text-sm font-semibold sm:col-span-2"
              htmlFor="user-email-filter"
            >
              ค้นหาจากอีเมล
              <Input
                id="user-email-filter"
                className="mt-2"
                type="search"
                value={draftEmailFilter}
                onChange={(event) => setDraftEmailFilter(event.target.value)}
                placeholder="เช่น team@example.com"
                aria-describedby="user-email-filter-help"
              />
              <span
                id="user-email-filter-help"
                className="mt-1 block text-xs font-normal text-[var(--ink-soft)]"
              >
                ระบบจะตัดช่องว่างและแปลงเป็นตัวพิมพ์เล็กก่อนค้นหา
              </span>
            </label>
            <label
              className="block text-sm font-semibold"
              htmlFor="user-role-filter"
            >
              บทบาท
              <select
                id="user-role-filter"
                className={`${selectClass} mt-2`}
                value={draftRoleFilter}
                onChange={(event) =>
                  setDraftRoleFilter(event.target.value as FilterRole)
                }
              >
                <option value="all">ทุกบทบาท</option>
                <option value="admin">ผู้ดูแลระบบ</option>
                <option value="user">ผู้ใช้</option>
              </select>
            </label>
            <label
              className="block text-sm font-semibold"
              htmlFor="user-enabled-filter"
            >
              สถานะ
              <select
                id="user-enabled-filter"
                className={`${selectClass} mt-2`}
                value={draftEnabledFilter}
                onChange={(event) =>
                  setDraftEnabledFilter(event.target.value as FilterEnabled)
                }
              >
                <option value="all">ทุกสถานะ</option>
                <option value="true">เปิดใช้งาน</option>
                <option value="false">ปิดใช้งาน</option>
              </select>
            </label>
            <div className="flex flex-wrap items-center gap-2 sm:col-span-2">
              <Button type="submit" disabled={rowActionsDisabled}>
                <Search size={16} />
                ค้นหา
              </Button>
              <Button
                type="button"
                variant="secondary"
                onClick={resetFilters}
                disabled={rowActionsDisabled}
              >
                <X size={16} />
                ล้างตัวกรอง
              </Button>
            </div>
          </form>
        </Card>
      </div>

      <LegacyAccountLinks
        requests={legacyAccountLinkRequests}
        nextCursor={legacyAccountLinkNextCursor}
        loading={legacyAccountLinkLoading}
        error={legacyAccountLinkError}
        feedback={legacyAccountLinkFeedback}
        pendingAction={legacyAccountLinkAction}
        headingRef={legacyAccountLinksHeadingRef}
        reload={() => setReloadVersion((value) => value + 1)}
        review={reviewLegacyAccountLink}
        loadMore={loadMoreLegacyAccountLinks}
      />

      <Card className="mt-6 overflow-hidden" aria-busy={listLoading}>
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[var(--line)] px-5 py-4 sm:px-6">
          <div>
            <h2
              id="admin-users-list-heading"
              ref={usersHeadingRef}
              tabIndex={-1}
              className="text-xl font-bold tracking-[-0.03em] focus:outline-none"
            >
              รายการบัญชี
            </h2>
            <p className="mt-1 text-sm text-[var(--ink-soft)]">
              แสดงครั้งละไม่เกิน 20 บัญชี
            </p>
          </div>
          <Button
            type="button"
            variant="ghost"
            size="sm"
            onClick={() => {
              clearTransientState();
              setReloadVersion((value) => value + 1);
            }}
            disabled={
              rowActionsDisabled ||
              legacyAccountLinkLoading ||
              legacyAccountLinkAction !== null
            }
          >
            {listLoading ? <Spinner /> : <RefreshCw size={15} />}
            โหลดใหม่
          </Button>
        </div>

        {listError ? (
          <div className="p-5 sm:p-6">
            <Notice tone="danger">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <span>{listError}</span>
                <Button
                  type="button"
                  variant="secondary"
                  size="sm"
                  onClick={() => setReloadVersion((value) => value + 1)}
                  disabled={
                    rowActionsDisabled ||
                    legacyAccountLinkLoading ||
                    legacyAccountLinkAction !== null
                  }
                >
                  ลองใหม่
                </Button>
              </div>
            </Notice>
          </div>
        ) : null}

        {listLoading && users.length === 0 ? (
          <div className="grid min-h-56 place-items-center gap-3 p-8 text-sm text-[var(--ink-soft)]">
            <Spinner />
            <span>กำลังโหลดรายการบัญชี…</span>
          </div>
        ) : null}
        {listLoading === false && users.length === 0 && listError === null ? (
          <div className="grid min-h-56 place-items-center p-8 text-center">
            <div>
              <UsersRound
                className="mx-auto mb-3 text-[var(--ink-soft)]"
                size={30}
              />
              <h3 className="font-semibold">ไม่พบบัญชีผู้ใช้</h3>
              <p className="mt-1 text-sm text-[var(--ink-soft)]">
                ลองเปลี่ยนตัวกรอง หรือสร้างบัญชีใหม่
              </p>
            </div>
          </div>
        ) : null}
        {users.length > 0 ? (
          <div className="overflow-x-auto">
            <table className="w-full min-w-[930px] text-left text-sm">
              <caption className="sr-only">รายการบัญชีผู้ใช้ในระบบ</caption>
              <thead className="bg-[var(--muted-soft)] text-xs uppercase tracking-[0.08em] text-[var(--ink-soft)]">
                <tr>
                  <th className="px-5 py-3 font-semibold sm:px-6" scope="col">
                    บัญชี
                  </th>
                  <th className="px-5 py-3 font-semibold" scope="col">
                    บทบาท
                  </th>
                  <th className="px-5 py-3 font-semibold" scope="col">
                    สถานะ
                  </th>
                  <th className="px-5 py-3 font-semibold" scope="col">
                    วันที่อัปเดต
                  </th>
                  <th className="px-5 py-3 font-semibold" scope="col">
                    การดำเนินการ
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[var(--line)]">
                {users.map((user) => {
                  const confirmationForRow =
                    confirmingAction?.user.id === user.id
                      ? confirmingAction
                      : null;
                  const isEditingEmail = editingEmailId === user.id;
                  const rowPending = pendingAction?.startsWith(`${user.id}:`);
                  const rowDisabled = rowActionsDisabled;
                  return (
                    <tr key={user.id} className="align-top">
                      <td className="px-5 py-5 sm:px-6">
                        <div className="flex min-w-64 items-start gap-3">
                          <span className="grid size-10 shrink-0 place-items-center rounded-full bg-[var(--accent-soft)] text-[var(--ink)]">
                            <UserRound size={18} />
                          </span>
                          <div className="min-w-0">
                            <p className="font-semibold text-[var(--ink)]">
                              {user.name}
                            </p>
                            {isEditingEmail ? (
                              <form
                                className="mt-2 max-w-sm space-y-2"
                                onSubmit={(event) =>
                                  submitEmailEdit(event, user)
                                }
                                noValidate
                                aria-busy={
                                  pendingAction === actionKey(user.id, "email")
                                }
                              >
                                <label
                                  className="sr-only"
                                  htmlFor={`edit-email-${user.id}`}
                                >
                                  อีเมลของ {user.name}
                                </label>
                                <Input
                                  id={`edit-email-${user.id}`}
                                  data-edit-email-id={user.id}
                                  type="email"
                                  value={editingEmail}
                                  onChange={(event) => {
                                    setEditingEmail(event.target.value);
                                    setEmailEditError(null);
                                  }}
                                  aria-describedby={
                                    emailEditError
                                      ? `email-edit-error-${user.id}`
                                      : undefined
                                  }
                                  aria-invalid={
                                    emailEditError ? "true" : undefined
                                  }
                                  autoComplete="email"
                                  required
                                />
                                {emailEditError ? (
                                  <p
                                    id={`email-edit-error-${user.id}`}
                                    className="text-xs font-semibold text-[var(--danger)]"
                                  >
                                    {emailEditError}
                                  </p>
                                ) : null}
                                <div className="flex flex-wrap gap-2">
                                  <Button
                                    type="submit"
                                    size="sm"
                                    data-user-action-id={actionKey(
                                      user.id,
                                      "email"
                                    )}
                                    disabled={rowDisabled}
                                  >
                                    {rowPending ? (
                                      <Spinner />
                                    ) : (
                                      <Mail size={14} />
                                    )}
                                    {rowPending ? "กำลังบันทึก…" : "บันทึกอีเมล"}
                                  </Button>
                                  <Button
                                    type="button"
                                    variant="secondary"
                                    size="sm"
                                    onClick={() => cancelEmailEdit(user)}
                                    disabled={rowDisabled}
                                  >
                                    ยกเลิก
                                  </Button>
                                </div>
                              </form>
                            ) : (
                              <p className="mt-1 break-all text-[var(--ink-soft)]">
                                {user.email}
                              </p>
                            )}
                            <p className="mt-2 text-xs text-[var(--ink-soft)]">
                              สร้างเมื่อ {formatDate(user.createdAt)} · อัปเดต{" "}
                              {formatDate(user.updatedAt)}
                            </p>
                            {user.mustChangePassword ? (
                              <div className="mt-2">
                                <Badge tone="warning">ต้องเปลี่ยนรหัสผ่าน</Badge>
                              </div>
                            ) : null}
                          </div>
                        </div>
                      </td>
                      <td className="px-5 py-5">
                        <Badge tone={roleBadgeTones[user.role]}>
                          {roleBadgeLabels[user.role]}
                        </Badge>
                      </td>
                      <td className="px-5 py-5">
                        <Badge tone={user.enabled ? "success" : "danger"}>
                          {user.enabled ? "เปิดใช้งาน" : "ปิดใช้งาน"}
                        </Badge>
                      </td>
                      <td className="whitespace-nowrap px-5 py-5 text-[var(--ink-soft)]">
                        {formatDate(user.updatedAt)}
                      </td>
                      <td className="px-5 py-5">
                        {confirmationForRow ? (
                          <div
                            className="max-w-sm space-y-3 rounded-[10px] border border-[var(--danger)]/30 bg-[var(--danger-soft)] p-3"
                            role="group"
                            aria-describedby={`confirm-description-${user.id}`}
                            aria-label={confirmationTitle(
                              confirmationForRow.kind
                            )}
                          >
                            <div className="flex items-start gap-2">
                              <ShieldCheck
                                className="mt-0.5 shrink-0 text-[var(--danger)]"
                                size={16}
                              />
                              <p
                                id={`confirm-description-${user.id}`}
                                className="text-sm text-[var(--ink)]"
                              >
                                {confirmationDescription(
                                  confirmationForRow.kind
                                )}
                              </p>
                            </div>
                            <div className="flex flex-wrap gap-2">
                              <Button
                                type="button"
                                variant="danger"
                                size="sm"
                                data-confirm-action-id={actionKey(
                                  user.id,
                                  confirmationForRow.kind
                                )}
                                onClick={() => {
                                  void mutateUser(
                                    user,
                                    confirmationForRow.kind,
                                    confirmationForRow.email
                                  );
                                }}
                                disabled={Boolean(pendingAction)}
                              >
                                {rowPending ? <Spinner /> : null}
                                {rowPending ? "กำลังดำเนินการ…" : "ยืนยัน"}
                              </Button>
                              <Button
                                type="button"
                                variant="secondary"
                                size="sm"
                                onClick={() => {
                                  restoreFocusKeyRef.current = actionKey(
                                    user.id,
                                    confirmationForRow.kind
                                  );
                                  setConfirmingAction(null);
                                }}
                                disabled={Boolean(pendingAction)}
                              >
                                ยกเลิก
                              </Button>
                            </div>
                          </div>
                        ) : (
                          <div className="flex max-w-sm flex-wrap gap-2">
                            {user.enabled ? (
                              <Button
                                type="button"
                                variant="danger"
                                size="sm"
                                data-user-action-id={actionKey(
                                  user.id,
                                  "disable"
                                )}
                                onClick={() =>
                                  requestConfirmation(user, "disable")
                                }
                                disabled={rowDisabled}
                                aria-label={`ปิดใช้งาน ${user.email}`}
                              >
                                ปิดใช้งาน
                              </Button>
                            ) : (
                              <Button
                                type="button"
                                variant="secondary"
                                size="sm"
                                data-user-action-id={actionKey(
                                  user.id,
                                  "enable"
                                )}
                                onClick={() => {
                                  restoreFocusKeyRef.current = actionKey(
                                    user.id,
                                    "enable"
                                  );
                                  void mutateUser(user, "enable");
                                }}
                                disabled={rowDisabled}
                                aria-label={`เปิดใช้งาน ${user.email}`}
                              >
                                เปิดใช้งาน
                              </Button>
                            )}
                            {user.role === "admin" ? (
                              <Button
                                type="button"
                                variant="secondary"
                                size="sm"
                                data-user-action-id={actionKey(
                                  user.id,
                                  "demote"
                                )}
                                onClick={() =>
                                  requestConfirmation(user, "demote")
                                }
                                disabled={rowDisabled}
                                aria-label={`ลดสิทธิ์ ${user.email} เป็นผู้ใช้`}
                              >
                                ลดสิทธิ์
                              </Button>
                            ) : (
                              <Button
                                type="button"
                                variant="secondary"
                                size="sm"
                                data-user-action-id={actionKey(
                                  user.id,
                                  "promote"
                                )}
                                onClick={() =>
                                  requestConfirmation(user, "promote")
                                }
                                disabled={rowDisabled}
                                aria-label={`เลื่อน ${user.email} เป็นผู้ดูแลระบบ`}
                              >
                                เลื่อนเป็นผู้ดูแล
                              </Button>
                            )}
                            {isEditingEmail ? null : (
                              <Button
                                type="button"
                                variant="secondary"
                                size="sm"
                                data-user-action-id={actionKey(
                                  user.id,
                                  "email"
                                )}
                                onClick={() => beginEmailEdit(user)}
                                disabled={rowDisabled}
                                aria-label={`แก้ไขอีเมล ${user.email}`}
                              >
                                <Pencil size={14} />
                                แก้ไขอีเมล
                              </Button>
                            )}
                            <Button
                              type="button"
                              variant="secondary"
                              size="sm"
                              data-user-action-id={actionKey(user.id, "reset")}
                              onClick={() => requestConfirmation(user, "reset")}
                              disabled={resetIsDisabled(
                                rowDisabled,
                                user.id,
                                authenticatedUser?.id
                              )}
                              aria-label={`ตั้งรหัสผ่านใหม่ให้ ${user.email}`}
                            >
                              <KeyRound size={14} />
                              ตั้งรหัสผ่านใหม่
                            </Button>
                            <Button
                              type="button"
                              variant="danger"
                              size="sm"
                              data-user-action-id={actionKey(user.id, "delete")}
                              onClick={() =>
                                requestConfirmation(user, "delete")
                              }
                              disabled={rowDisabled}
                              aria-label={`ลบบัญชี ${user.email} ถาวร`}
                            >
                              <Trash2 size={14} />
                              ลบบัญชีถาวร
                            </Button>
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        ) : null}

        <nav
          className="flex flex-wrap items-center justify-between gap-3 border-t border-[var(--line)] px-5 py-4 sm:px-6"
          aria-label="การแบ่งหน้ารายการผู้ใช้"
        >
          <span className="text-sm text-[var(--ink-soft)]">
            หน้า {cursorHistory.length + 1}
          </span>
          <div className="flex items-center gap-2">
            <Button
              type="button"
              variant="secondary"
              size="sm"
              onClick={goToPreviousPage}
              disabled={cursorHistory.length === 0 || rowActionsDisabled}
            >
              <ChevronLeft size={16} />
              ก่อนหน้า
            </Button>
            <Button
              type="button"
              variant="secondary"
              size="sm"
              onClick={goToNextPage}
              disabled={!nextCursor || rowActionsDisabled}
            >
              ถัดไป
              <ChevronRight size={16} />
            </Button>
          </div>
        </nav>
      </Card>
    </>
  );
};

export const Route = createFileRoute("/admin/users")({
  component: AdminUsersRoute,
});
