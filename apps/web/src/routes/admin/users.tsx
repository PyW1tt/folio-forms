import {
  createFileRoute,
  useBlocker,
  useNavigate,
} from "@tanstack/react-router";
import type { FormEvent } from "react";
import { useEffect, useRef, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import { Notice } from "@/components/ui";
import {
  CreateUserPanel,
  TemporaryPasswordNotice,
  UserFilters,
} from "@/features/admin-users/account-panels";
import { LegacyAccountLinks } from "@/features/admin-users/legacy-account-links";
import {
  actionFailureMessage,
  actionKey,
  actionSuccessMessage,
  actionsAreDisabled,
  credentialRequestIsInFlight,
  EMAIL_PATTERN,
  errorMessageFor,
  normalizeEmail,
} from "@/features/admin-users/model";
import type {
  ActionKind,
  Confirmation,
  ConfirmKind,
  FilterEnabled,
  FilterRole,
} from "@/features/admin-users/model";
import { UserList } from "@/features/admin-users/user-list";
import { ApiError, apiDelete, apiGet, apiPatch, apiPost } from "@/lib/api";
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

  const reloadUsers = () => {
    clearTransientState();
    setReloadVersion((value) => value + 1);
  };

  const retryUsers = () => {
    setReloadVersion((value) => value + 1);
  };

  const changeEditingEmail = (value: string) => {
    setEditingEmail(value);
    setEmailEditError(null);
  };

  const confirmMutation = (
    user: AdminUser,
    kind: ConfirmKind,
    nextEmail?: string
  ) => {
    void mutateUser(user, kind, nextEmail);
  };

  const cancelConfirmation = (user: AdminUser, kind: ConfirmKind) => {
    restoreFocusKeyRef.current = actionKey(user.id, kind);
    setConfirmingAction(null);
  };

  const enableUser = (user: AdminUser) => {
    restoreFocusKeyRef.current = actionKey(user.id, "enable");
    void mutateUser(user, "enable");
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
        <CreateUserPanel
          createName={createName}
          createEmail={createEmail}
          createRole={createRole}
          createError={createError}
          creating={creating}
          rowActionsDisabled={rowActionsDisabled}
          setCreateName={setCreateName}
          setCreateEmail={setCreateEmail}
          setCreateRole={setCreateRole}
          createUser={createUser}
        />

        <UserFilters
          draftEmailFilter={draftEmailFilter}
          draftRoleFilter={draftRoleFilter}
          draftEnabledFilter={draftEnabledFilter}
          setDraftEmailFilter={setDraftEmailFilter}
          setDraftRoleFilter={setDraftRoleFilter}
          setDraftEnabledFilter={setDraftEnabledFilter}
          applyFilters={applyFilters}
          resetFilters={resetFilters}
          rowActionsDisabled={rowActionsDisabled}
        />
      </div>

      <LegacyAccountLinks
        requests={legacyAccountLinkRequests}
        nextCursor={legacyAccountLinkNextCursor}
        loading={legacyAccountLinkLoading}
        error={legacyAccountLinkError}
        feedback={legacyAccountLinkFeedback}
        pendingAction={legacyAccountLinkAction}
        headingRef={legacyAccountLinksHeadingRef}
        reload={retryUsers}
        review={reviewLegacyAccountLink}
        loadMore={loadMoreLegacyAccountLinks}
      />

      <UserList
        users={users}
        listLoading={listLoading}
        listError={listError}
        nextCursor={nextCursor}
        cursorHistory={cursorHistory}
        usersHeadingRef={usersHeadingRef}
        rowActionsDisabled={rowActionsDisabled}
        legacyAccountLinkLoading={legacyAccountLinkLoading}
        legacyAccountLinkAction={legacyAccountLinkAction}
        confirmingAction={confirmingAction}
        editingEmailId={editingEmailId}
        editingEmail={editingEmail}
        emailEditError={emailEditError}
        pendingAction={pendingAction}
        authenticatedUserId={authenticatedUser?.id}
        reload={reloadUsers}
        retry={retryUsers}
        submitEmailEdit={submitEmailEdit}
        changeEditingEmail={changeEditingEmail}
        cancelEmailEdit={cancelEmailEdit}
        confirmMutation={confirmMutation}
        cancelConfirmation={cancelConfirmation}
        requestConfirmation={requestConfirmation}
        beginEmailEdit={beginEmailEdit}
        enableUser={enableUser}
        goToPreviousPage={goToPreviousPage}
        goToNextPage={goToNextPage}
      />
    </>
  );
};

export const Route = createFileRoute("/admin/users")({
  component: AdminUsersRoute,
});
