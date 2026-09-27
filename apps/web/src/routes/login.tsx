import {
  createFileRoute,
  useNavigate,
  useSearch,
} from "@tanstack/react-router";
import { ArrowRight, KeyRound, LockKeyhole } from "lucide-react";
import type { FormEvent } from "react";
import { useEffect, useRef, useState } from "react";

import { Button, Card, Input, Notice, Spinner } from "@/components/ui";
import {
  API_ORIGIN,
  legacySsoEnabled as checkLegacySsoEnabled,
  safeReturnPath,
} from "@/lib/api";
import { authErrorMessage, roleFor, useAuth } from "@/lib/auth";

const LoginRoute = () => {
  const { loading: authLoading, signIn, user } = useAuth();
  const navigate = useNavigate();
  const { legacySso, returnTo } = useSearch({ from: "/login" });
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [legacySsoAvailable, setLegacySsoAvailable] = useState(false);
  const errorRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (error) {
      errorRef.current?.focus();
    }
  }, [error]);

  useEffect(() => {
    let active = true;
    void checkLegacySsoEnabled()
      .then((enabled) => {
        if (active) {
          setLegacySsoAvailable(enabled);
        }
      })
      .catch(() => {
        if (active) {
          setLegacySsoAvailable(false);
        }
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (legacySso === "failed") {
      setError("ไม่สามารถเข้าสู่ระบบผ่านระบบเดิมได้ กรุณาลองอีกครั้ง");
    }
  }, [legacySso]);

  useEffect(() => {
    if (!user) {
      return;
    }

    const redirect = async () => {
      if (user.mustChangePassword) {
        await navigate({
          replace: true,
          search: { returnTo },
          to: "/change-password",
        });
        return;
      }
      if (returnTo) {
        await navigate({ replace: true, to: returnTo });
        return;
      }
      await navigate({
        replace: true,
        to: roleFor(user) === "admin" ? "/admin" : "/dashboard",
      });
    };

    void redirect();
  }, [navigate, returnTo, user]);

  if (authLoading) {
    return (
      <div className="grid min-h-screen place-items-center gap-3 bg-[var(--ink)] text-sm text-[var(--paper)]">
        <Spinner />
        <span>กำลังตรวจสอบเซสชัน…</span>
      </div>
    );
  }

  if (user) {
    return (
      <div className="grid min-h-screen place-items-center gap-3 bg-[var(--ink)] text-sm text-[var(--paper)]">
        <Spinner />
        <span>กำลังเปิดพื้นที่ทำงาน…</span>
      </div>
    );
  }

  const submit = async (event: FormEvent) => {
    event.preventDefault();
    setError(null);
    setSuccess(null);
    if (!email.trim() || !password) {
      setError("กรุณากรอกอีเมลและรหัสผ่าน");
      return;
    }
    setBusy(true);
    try {
      const next = await signIn(email.trim(), password);
      if (next.mustChangePassword) {
        setSuccess("เข้าสู่ระบบสำเร็จ กำลังไปตั้งรหัสผ่านใหม่…");
        await navigate({
          replace: true,
          search: { returnTo },
          to: "/change-password",
        });
        return;
      }
      setSuccess("เข้าสู่ระบบสำเร็จ กำลังเปิดพื้นที่ทำงาน…");
      const destination =
        returnTo ?? (roleFor(next) === "admin" ? "/admin" : "/dashboard");
      await navigate({ replace: true, to: destination });
    } catch (caughtError) {
      setError(authErrorMessage(caughtError, "signIn"));
    } finally {
      setBusy(false);
    }
  };
  const startLegacySso = () => {
    const startUrl = new URL("/api/legacy-sso/start", API_ORIGIN);
    if (returnTo) {
      startUrl.searchParams.set("returnTo", returnTo);
    }
    window.location.assign(startUrl.href);
  };

  return (
    <div className="grid min-h-screen place-items-center bg-[var(--ink)] px-5 py-12">
      <div className="w-full max-w-[440px] animate-rise">
        <div className="mb-8 flex items-center gap-3 text-[var(--paper)]">
          <span className="grid size-10 place-items-center rounded-[10px] bg-[var(--accent)] text-[var(--ink)]">
            <LockKeyhole size={19} />
          </span>
          <span className="text-xl font-bold tracking-[-0.03em]">
            Folio Forms
          </span>
        </div>
        <Card className="p-6 sm:p-8">
          <h1 className="text-2xl font-bold tracking-[-0.04em]">เข้าสู่ระบบ</h1>
          <p className="mt-2 text-sm text-[var(--ink-soft)]">
            เข้าสู่พื้นที่ทำงานของคุณเพื่อดำเนินการต่อ
          </p>
          {error ? (
            <div ref={errorRef} id="login-error" className="mt-5" tabIndex={-1}>
              <Notice tone="danger">{error}</Notice>
            </div>
          ) : null}
          {success ? (
            <div className="mt-5">
              <Notice tone="success">{success}</Notice>
            </div>
          ) : null}
          <form
            className="mt-7 space-y-5"
            onSubmit={submit}
            aria-busy={busy}
            noValidate
          >
            <label className="block text-sm font-semibold" htmlFor="email">
              อีเมล
              <Input
                id="email"
                className="mt-2"
                type="email"
                autoComplete="email"
                autoFocus
                value={email}
                onChange={(event) => setEmail(event.target.value)}
                aria-describedby={error ? "login-error" : undefined}
                aria-invalid={error ? "true" : undefined}
                required
              />
            </label>
            <label className="block text-sm font-semibold" htmlFor="password">
              รหัสผ่าน
              <Input
                id="password"
                className="mt-2"
                type="password"
                autoComplete="current-password"
                minLength={12}
                maxLength={128}
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                aria-describedby={error ? "login-error" : undefined}
                aria-invalid={error ? "true" : undefined}
                required
              />
            </label>
            <Button className="w-full" size="lg" type="submit" disabled={busy}>
              {busy ? <Spinner /> : <KeyRound size={17} />}
              {busy ? "กำลังเข้าสู่ระบบ…" : "เข้าสู่ระบบ"}
              <ArrowRight size={16} />
            </Button>
          </form>
          {legacySsoAvailable ? (
            <div className="mt-5 space-y-4">
              <div className="flex items-center gap-3 text-xs text-[var(--ink-soft)]">
                <span className="h-px flex-1 bg-[var(--line)]" />
                <span>หรือเข้าสู่ระบบผ่านระบบเดิม</span>
                <span className="h-px flex-1 bg-[var(--line)]" />
              </div>
              <Button
                className="w-full"
                size="lg"
                type="button"
                variant="secondary"
                onClick={startLegacySso}
              >
                <LockKeyhole size={17} />
                เข้าสู่ระบบผ่านระบบเดิม
              </Button>
            </div>
          ) : null}
        </Card>
        <p className="mt-5 text-center text-xs text-[#b8c5c5]">
          พื้นที่ทำงานภายใน · ใช้โทเค็นเซสชันแบบไม่เปิดเผยข้อมูล
        </p>
      </div>
    </div>
  );
};

export const Route = createFileRoute("/login")({
  component: LoginRoute,
  validateSearch: (search) => ({
    returnTo: safeReturnPath(search.returnTo) ?? undefined,
    ...(search.legacySso === "failed" ? { legacySso: "failed" as const } : {}),
  }),
});
