import {
  createRootRoute,
  Navigate,
  Outlet,
  useLocation,
  useNavigate,
} from "@tanstack/react-router";
import { useEffect, useState } from "react";

import { Button, Notice, Spinner } from "@/components/ui";
import { legacySsoReturnPath, safeReturnPath, startLegacySso } from "@/lib/api";
import { AuthProvider, roleFor, useAuth } from "@/lib/auth";
import {
  afterEditorSave,
  SESSION_WARNING_WINDOW_MS,
  shouldWarnBeforeSessionExpiry,
} from "@/lib/form-lifecycle";

import "@/index.css";

const MAX_TIMEOUT_MS = 2_147_483_647;

const currentReturnPath = () => {
  if (typeof window === "undefined") {
    return null;
  }
  return safeReturnPath(`${window.location.pathname}${window.location.search}`);
};

// oxlint-disable-next-line complexity -- Coordinates session expiry, reauthentication, and routing in the app shell.
const AuthGate = () => {
  const { error: authError, expiresAt, loading, signOut, user } = useAuth();
  const location = useLocation();
  const navigate = useNavigate();
  const [sessionExpiringSoon, setSessionExpiringSoon] = useState(false);
  const [reauthenticating, setReauthenticating] = useState(false);
  const [reauthenticationFailed, setReauthenticationFailed] = useState(false);
  const { pathname } = location;
  const usesLegacySso = Boolean(user && roleFor(user) === "user");

  useEffect(() => {
    const expiresAtMs = expiresAt ? Date.parse(expiresAt) : Number.NaN;
    if (!Number.isFinite(expiresAtMs)) {
      setSessionExpiringSoon(false);
      return;
    }

    let timeoutId: number | undefined;
    const updateWarning = () => {
      const remaining = expiresAtMs - Date.now();
      setSessionExpiringSoon(
        shouldWarnBeforeSessionExpiry(expiresAt ?? undefined, Date.now())
      );
      if (remaining > SESSION_WARNING_WINDOW_MS) {
        timeoutId = window.setTimeout(
          updateWarning,
          Math.min(remaining - SESSION_WARNING_WINDOW_MS, MAX_TIMEOUT_MS)
        );
      }
    };

    updateWarning();
    return () => {
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
    };
  }, [expiresAt]);

  const reauthenticate = async () => {
    if (reauthenticating) {
      return;
    }
    setReauthenticating(true);
    setReauthenticationFailed(false);
    try {
      await afterEditorSave(async () => {
        if (usesLegacySso) {
          const { authorizationUrl } = await startLegacySso(
            legacySsoReturnPath(
              window.location.pathname,
              window.location.search
            )
          );
          window.location.assign(authorizationUrl);
          return;
        }
        const returnTo = currentReturnPath() ?? undefined;
        await signOut();
        await navigate({
          replace: true,
          search: { returnTo },
          to: "/login",
        });
      });
    } catch {
      setReauthenticationFailed(true);
    } finally {
      setReauthenticating(false);
    }
  };

  if (loading) {
    return (
      <div
        className="grid min-h-screen place-items-center gap-3 text-sm text-[var(--ink-soft)]"
        aria-busy="true"
      >
        <Spinner />
        <span>กำลังตรวจสอบเซสชัน…</span>
      </div>
    );
  }

  const authenticationRoute =
    pathname === "/login" ||
    pathname === "/change-password" ||
    pathname === "/handoff";
  if (!user && !authenticationRoute) {
    return (
      <Navigate
        to="/login"
        search={{ returnTo: currentReturnPath() ?? undefined }}
        replace
      />
    );
  }

  const restricted =
    user?.mustChangePassword &&
    pathname !== "/login" &&
    pathname !== "/change-password" &&
    pathname !== "/handoff";
  if (restricted) {
    return (
      <Navigate
        to="/change-password"
        search={{ returnTo: currentReturnPath() ?? undefined }}
        replace
      />
    );
  }

  let reauthenticationLabel = usesLegacySso ? "เข้าสู่ระบบระบบเดิม" : "เข้าสู่ระบบใหม่";
  let reauthenticationMessage = usesLegacySso
    ? "เซสชันจะหมดอายุภายใน 5 นาที กรุณาบันทึกงานก่อน แล้วเข้าสู่ระบบระบบเดิม"
    : "เซสชันจะหมดอายุภายใน 5 นาที กรุณาบันทึกงานก่อน แล้วเข้าสู่ระบบใหม่";
  if (reauthenticationFailed) {
    reauthenticationMessage = usesLegacySso
      ? "ไม่สามารถเชื่อมต่อบัญชีระบบเดิมได้ กรุณาลองอีกครั้ง"
      : "ไม่สามารถยกเลิกเซสชันเดิมได้ กรุณาลองอีกครั้ง";
  }
  if (reauthenticating) {
    reauthenticationLabel = "กำลังเตรียมเซสชันใหม่…";
  } else if (reauthenticationFailed) {
    reauthenticationLabel = "ลองอีกครั้ง";
  }

  const showSessionWarning =
    Boolean(user) &&
    sessionExpiringSoon &&
    pathname !== "/login" &&
    pathname !== "/change-password";
  const showSignOutCleanupFailure =
    !user && pathname === "/login" && authError === "sign_out_cleanup_failed";
  return (
    <>
      {showSignOutCleanupFailure ? (
        <div className="mx-auto max-w-[1240px] px-5 pt-4 lg:px-8">
          <Notice tone="danger">ออกจากระบบแล้ว แต่การล้างข้อมูล AI ไม่สำเร็จ</Notice>
        </div>
      ) : null}

      {showSessionWarning ? (
        <div className="mx-auto max-w-[1240px] px-5 pt-4 lg:px-8">
          <Notice tone="danger">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <span>{reauthenticationMessage}</span>
              <Button
                variant="secondary"
                size="sm"
                type="button"
                disabled={reauthenticating}
                onClick={reauthenticate}
              >
                {reauthenticating ? <Spinner /> : null}
                {reauthenticationLabel}
              </Button>
            </div>
          </Notice>
        </div>
      ) : null}
      <Outlet />
    </>
  );
};

const RootLayout = () => (
  <AuthProvider>
    <AuthGate />
  </AuthProvider>
);

export const Route = createRootRoute({ component: RootLayout });
