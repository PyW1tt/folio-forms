import {
  createRootRoute,
  Navigate,
  Outlet,
  useLocation,
  useNavigate,
} from "@tanstack/react-router";
import { useEffect, useState } from "react";

import { Button, Notice, Spinner } from "@/components/ui";
import { safeReturnPath } from "@/lib/api";
import { AuthProvider, useAuth } from "@/lib/auth";
import {
  SESSION_WARNING_WINDOW_MS,
  createDeferred,
  shouldWarnBeforeSessionExpiry,
} from "@/lib/form-lifecycle";

import "@/index.css";

const MAX_TIMEOUT_MS = 2_147_483_647;
interface ReauthenticationRequestDetail {
  handled: boolean;
  resolve: (allowed: boolean) => void;
}

const requestEditorSave = (): Promise<boolean> => {
  const result = createDeferred<boolean>();
  const detail: ReauthenticationRequestDetail = {
    handled: false,
    resolve: result.resolve,
  };
  window.dispatchEvent(
    new CustomEvent("folio:before-reauth", {
      detail,
    })
  );
  return detail.handled ? result.promise : Promise.resolve(true);
};

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
      if (!(await requestEditorSave())) {
        return;
      }
      const returnTo = currentReturnPath() ?? undefined;
      await signOut();
      await navigate({
        replace: true,
        search: { returnTo },
        to: "/login",
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

  let reauthenticationLabel = "เข้าสู่ระบบใหม่";
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
    !user &&
    pathname === "/login" &&
    authError === "sign_out_cleanup_failed";
  return (
    <>
      {showSignOutCleanupFailure ? (
        <div className="mx-auto max-w-[1240px] px-5 pt-4 lg:px-8">
          <Notice tone="danger">
            ออกจากระบบแล้ว แต่การล้างข้อมูล AI ไม่สำเร็จ
          </Notice>
        </div>
      ) : null}

      {showSessionWarning ? (
        <div className="mx-auto max-w-[1240px] px-5 pt-4 lg:px-8">
          <Notice tone="danger">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <span>
                {reauthenticationFailed
                  ? "ไม่สามารถยกเลิกเซสชันเดิมได้ กรุณาลองอีกครั้ง"
                  : "เซสชันจะหมดอายุภายใน 5 นาที กรุณาบันทึกงานก่อน แล้วเข้าสู่ระบบใหม่"}
              </span>
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
