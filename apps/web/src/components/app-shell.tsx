import { Link, useLocation, useNavigate } from "@tanstack/react-router";
import {
  ClipboardList,
  FilePlus2,
  FileText,
  KeyRound,
  LogOut,
  Menu,
  ScrollText,
  ShieldCheck,
  UserRound,
  UsersRound,
  WandSparkles,
  X,
} from "lucide-react";
import { useState } from "react";

import { LegacySsoSwitchButton } from "@/components/legacy-sso-switch-button";
import { Button, Notice, Spinner } from "@/components/ui";
import { ApiError } from "@/lib/api";
import { roleFor, useAuth } from "@/lib/auth";

export const AppShell = ({ children }: { children: React.ReactNode }) => {
  const { user, signOut } = useAuth();
  const [logoutError, setLogoutError] = useState<"cleanup" | "sign-out" | null>(
    null
  );
  const [loggingOut, setLoggingOut] = useState(false);
  const [open, setOpen] = useState(false);
  const navigate = useNavigate();
  const location = useLocation();
  const role = roleFor(user);
  const formsAreActive =
    location.pathname === "/admin" ||
    location.pathname.startsWith("/admin/forms/");
  const auditAreActive = location.pathname.startsWith("/admin/audit");
  const resultsAreActive = location.pathname.startsWith("/admin/results");
  const authoringIsActive = location.pathname.startsWith("/admin/ai-authoring");
  const [legacySsoError, setLegacySsoError] = useState(false);

  const logout = async () => {
    if (loggingOut) {
      return;
    }
    setLoggingOut(true);
    setLogoutError(null);
    try {
      await signOut();
      await navigate({
        replace: true,
        search: { returnTo: undefined },
        to: "/login",
      });
    } catch (error) {
      setLogoutError(
        error instanceof ApiError && error.sessionRevoked
          ? "cleanup"
          : "sign-out"
      );
    } finally {
      setLoggingOut(false);
    }
  };

  return (
    <div className="min-h-screen bg-[var(--canvas)] text-[var(--ink)]">
      <header className="border-b border-[var(--line)] bg-[var(--paper)]">
        <div className="mx-auto flex max-w-[1240px] items-center justify-between px-5 py-4 lg:px-8">
          <Link
            to={role === "admin" ? "/admin" : "/dashboard"}
            className="flex items-center gap-3"
            onClick={() => setOpen(false)}
          >
            <span className="grid size-9 place-items-center rounded-[10px] bg-[var(--ink)] text-[var(--accent)]">
              <FileText size={18} strokeWidth={2.5} />
            </span>
            <span className="text-lg font-bold tracking-[-0.03em]">
              Folio Forms
            </span>
          </Link>
          <button
            className="rounded-lg p-2 lg:hidden"
            type="button"
            aria-label="เปิดหรือปิดเมนูนำทาง"
            aria-expanded={open}
            aria-controls="main-navigation"
            onClick={() => setOpen((value) => !value)}
          >
            {open ? <X size={20} /> : <Menu size={20} />}
          </button>
          <nav
            id="main-navigation"
            className={`${open ? "absolute inset-x-0 top-[73px] z-10 flex border-b border-[var(--line)] bg-[var(--paper)] p-5" : "hidden"} flex-col gap-2 lg:static lg:flex lg:flex-row lg:items-center lg:border-0 lg:bg-transparent lg:p-0`}
            aria-label="เมนูหลัก"
          >
            {role === "admin" ? (
              <>
                <Link
                  to="/admin"
                  className={`rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)] ${
                    formsAreActive
                      ? "bg-[var(--accent-soft)] text-[var(--ink)]"
                      : ""
                  }`}
                  activeOptions={{ exact: true }}
                  aria-current={formsAreActive ? "page" : undefined}
                  onClick={() => setOpen(false)}
                >
                  แบบฟอร์ม
                </Link>
                <Link
                  to="/admin/results"
                  search={{ form: undefined }}
                  className={`rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)] ${
                    resultsAreActive
                      ? "bg-[var(--accent-soft)] text-[var(--ink)]"
                      : ""
                  }`}
                  aria-current={resultsAreActive ? "page" : undefined}
                  onClick={() => setOpen(false)}
                >
                  <ClipboardList className="mr-1 inline" size={15} />
                  ผลลัพธ์
                </Link>
                <Link
                  to="/admin/audit"
                  className={`rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)] ${
                    auditAreActive
                      ? "bg-[var(--accent-soft)] text-[var(--ink)]"
                      : ""
                  }`}
                  aria-current={auditAreActive ? "page" : undefined}
                  onClick={() => setOpen(false)}
                >
                  <ScrollText className="mr-1 inline" size={15} />
                  Audit Trail
                </Link>
                <Link
                  to="/admin/users"
                  className="rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)]"
                  activeProps={{
                    className: "bg-[var(--accent-soft)] text-[var(--ink)]",
                  }}
                  onClick={() => setOpen(false)}
                >
                  <UsersRound className="mr-1 inline" size={15} />
                  จัดการผู้ใช้
                </Link>
                <Link
                  to="/admin/forms/new"
                  className="rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)]"
                  onClick={() => setOpen(false)}
                >
                  <FilePlus2 className="mr-1 inline" size={15} />
                  สร้างแบบฟอร์ม
                </Link>
                <Link
                  to="/admin/ai-authoring"
                  className={`rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)] ${
                    authoringIsActive
                      ? "bg-[var(--accent-soft)] text-[var(--ink)]"
                      : ""
                  }`}
                  aria-current={authoringIsActive ? "page" : undefined}
                  onClick={() => setOpen(false)}
                >
                  <WandSparkles className="mr-1 inline" size={15} />
                  AI Authoring
                </Link>
              </>
            ) : (
              <Link
                to="/dashboard"
                className="rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)]"
                activeProps={{
                  className: "bg-[var(--accent-soft)] text-[var(--ink)]",
                }}
                onClick={() => setOpen(false)}
              >
                My responses
              </Link>
            )}
            {role === "user" ? (
              <LegacySsoSwitchButton
                onError={() => setLegacySsoError(true)}
                onStart={() => {
                  setOpen(false);
                  setLegacySsoError(false);
                }}
              />
            ) : null}
            <Link
              to="/change-password"
              search={{
                returnTo: role === "admin" ? "/admin" : "/dashboard",
              }}
              className="rounded-lg px-3 py-2 text-sm font-semibold text-[var(--ink-soft)] hover:bg-[var(--accent-soft)]"
              onClick={() => setOpen(false)}
            >
              <KeyRound className="mr-1 inline" size={15} />
              เปลี่ยนรหัสผ่าน
            </Link>
            <span className="mx-1 hidden h-6 w-px bg-[var(--line)] lg:block" />
            <span className="flex items-center gap-2 px-3 py-2 text-sm text-[var(--ink-soft)]">
              <span className="grid size-7 place-items-center rounded-full bg-[var(--success-soft)] text-[var(--success)]">
                {role === "admin" ? (
                  <ShieldCheck size={15} />
                ) : (
                  <UserRound size={15} />
                )}
              </span>
              <span className="max-w-36 truncate">
                {user?.name ?? user?.email}
              </span>
            </span>
            <Button
              variant="ghost"
              size="sm"
              type="button"
              disabled={loggingOut}
              onClick={logout}
            >
              {loggingOut ? <Spinner /> : <LogOut size={15} />}
              {loggingOut ? "กำลังออกจากระบบ…" : "ออกจากระบบ"}
            </Button>
          </nav>
        </div>
      </header>
      {logoutError ? (
        <div className="mx-auto max-w-[1240px] px-5 pt-4 lg:px-8">
          <Notice tone="danger">
            {logoutError === "cleanup"
              ? "ออกจากระบบแล้ว แต่การล้างข้อมูล AI ไม่สำเร็จ"
              : "ไม่สามารถออกจากระบบได้ กรุณาลองใหม่อีกครั้ง"}
          </Notice>
        </div>
      ) : null}
      {legacySsoError ? (
        <div className="mx-auto max-w-[1240px] px-5 pt-4 lg:px-8">
          <Notice tone="danger">ไม่สามารถเชื่อมต่อบัญชีระบบเดิมได้ กรุณาลองอีกครั้ง</Notice>
        </div>
      ) : null}
      <main className="mx-auto max-w-[1240px] px-5 py-8 lg:px-8 lg:py-12">
        {children}
      </main>
    </div>
  );
};

export const PageHeader = ({
  title,
  description,
  action,
}: {
  title: string;
  description?: string;
  action?: React.ReactNode;
}) => (
  <div className="mb-8 flex flex-col justify-between gap-4 sm:flex-row sm:items-end">
    <div>
      <h1 className="text-3xl font-bold tracking-[-0.04em] lg:text-4xl">
        {title}
      </h1>
      {description ? (
        <p className="mt-2 max-w-[68ch] text-[var(--ink-soft)]">
          {description}
        </p>
      ) : null}
    </div>
    {action}
  </div>
);
