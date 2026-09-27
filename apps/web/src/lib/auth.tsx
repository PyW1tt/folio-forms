import { createContext, useContext, useEffect, useState } from "react";

import {
  ApiError,
  claimLegacySsoSession,
  clearToken,
  getSession,
  getToken,
  replacePassword as requestReplacePassword,
  setToken,
  signIn as requestSignIn,
  signOut as requestSignOut,
} from "@/lib/api";
import type { Role, SessionUser } from "@/lib/api";

interface AuthContextValue {
  user: SessionUser | null;
  loading: boolean;
  error: string | null;
  expiresAt: string | null;
  clearSession: () => void;
  signIn: (email: string, password: string) => Promise<SessionUser>;
  replacePassword: (
    currentPassword: string,
    newPassword: string
  ) => Promise<void>;
  signOut: () => Promise<void>;
}

const AuthContext = createContext<AuthContextValue | null>(null);
const MAX_TIMEOUT_MS = 2_147_483_647;

export const authErrorMessage = (
  error: unknown,
  action: "signIn" | "changePassword"
) => {
  const apiError = error instanceof ApiError ? error : null;
  if (action === "signIn") {
    if (apiError?.code === "login_throttled" || apiError?.status === 429) {
      return "พยายามเข้าสู่ระบบบ่อยเกินไป กรุณารอสักครู่แล้วลองใหม่";
    }
    if (apiError?.code === "invalid_credentials" || apiError?.status === 401) {
      return "อีเมลหรือรหัสผ่านไม่ถูกต้อง";
    }
    return "ไม่สามารถเข้าสู่ระบบได้ กรุณาลองใหม่อีกครั้ง";
  }

  if (apiError?.code === "invalid_current_password") {
    return "รหัสผ่านปัจจุบันไม่ถูกต้อง";
  }
  if (apiError?.code === "password_too_short") {
    return "รหัสผ่านใหม่ต้องมีอย่างน้อย 12 ตัวอักษร";
  }
  if (apiError?.code === "password_too_long") {
    return "รหัสผ่านใหม่ต้องมีไม่เกิน 128 ตัวอักษร";
  }
  if (apiError?.status === 401) {
    return "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่";
  }
  return "ไม่สามารถเปลี่ยนรหัสผ่านได้ กรุณาลองใหม่อีกครั้ง";
};

export const AuthProvider = ({ children }: { children: React.ReactNode }) => {
  const [user, setUser] = useState<SessionUser | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [expiresAt, setExpiresAt] = useState<string | null>(null);

  useEffect(() => {
    let active = true;
    const loadSession = async () => {
      let sessionToken = getToken();
      try {
        if (!sessionToken) {
          sessionToken = await claimLegacySsoSession();
          if (sessionToken) {
            setToken(sessionToken);
          }
        }
        if (!sessionToken) {
          return;
        }
        const session = await getSession();
        if (active && getToken() === sessionToken) {
          setUser(session.user);
          setExpiresAt(session.session.expiresAt);
        }
      } catch {
        if (active && sessionToken && getToken() === sessionToken) {
          clearToken();
          setUser(null);
          setExpiresAt(null);
        }
      } finally {
        if (active) {
          setLoading(false);
        }
      }
    };

    void loadSession();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!expiresAt) {
      return;
    }
    const expiresAtMs = Date.parse(expiresAt);
    if (!Number.isFinite(expiresAtMs)) {
      return;
    }

    const sessionToken = getToken();
    let timeoutId: number | undefined;
    const expireSession = () => {
      if (getToken() !== sessionToken) {
        return;
      }
      clearToken();
      setUser(null);
      setExpiresAt(null);
    };
    const scheduleExpiry = () => {
      const remaining = expiresAtMs - Date.now();
      if (remaining <= 0) {
        expireSession();
        return;
      }
      timeoutId = window.setTimeout(
        scheduleExpiry,
        Math.min(remaining, MAX_TIMEOUT_MS)
      );
    };

    scheduleExpiry();
    return () => {
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
    };
  }, [expiresAt, user]);

  const signIn = async (email: string, password: string) => {
    setError(null);
    try {
      await requestSignIn(email, password);
      const session = await getSession();
      setUser(session.user);
      setExpiresAt(session.session.expiresAt);
      return session.user;
    } catch (caughtError) {
      setError(
        caughtError instanceof ApiError ? caughtError.code : "sign_in_failed"
      );
      throw caughtError;
    }
  };

  const replacePassword = async (
    currentPassword: string,
    newPassword: string
  ) => {
    setError(null);
    try {
      await requestReplacePassword(currentPassword, newPassword);
      clearToken();
      setUser(null);
      setExpiresAt(null);
    } catch (caughtError) {
      setError(
        caughtError instanceof ApiError
          ? caughtError.code
          : "password_change_failed"
      );
      throw caughtError;
    }
  };

  const clearSession = () => {
    clearToken();
    setUser(null);
    setExpiresAt(null);
  };

  const signOut = async () => {
    await requestSignOut();
    clearSession();
  };

  return (
    <AuthContext.Provider
      value={{
        clearSession,
        error,
        expiresAt,
        loading,
        replacePassword,
        signIn,
        signOut,
        user,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error("useAuth must be used within AuthProvider");
  }
  return context;
};

export const roleFor = (user: SessionUser | null): Role =>
  user?.role === "admin" ? "admin" : "user";
