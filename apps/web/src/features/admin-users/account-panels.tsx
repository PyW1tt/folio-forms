import { Copy, Plus, Search, X } from "lucide-react";
import type { FormEvent, RefObject } from "react";

import { Button, Card, Input, Notice, Spinner } from "@/components/ui";
import type { Role } from "@/lib/api";

import { selectClass } from "./model";
import type { FilterEnabled, FilterRole } from "./model";

export const TemporaryPasswordNotice = ({
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

export const CreateUserPanel = ({
  createName,
  createEmail,
  createRole,
  createError,
  creating,
  rowActionsDisabled,
  setCreateName,
  setCreateEmail,
  setCreateRole,
  createUser,
}: {
  createName: string;
  createEmail: string;
  createRole: Role;
  createError: string | null;
  creating: boolean;
  rowActionsDisabled: boolean;
  setCreateName: (value: string) => void;
  setCreateEmail: (value: string) => void;
  setCreateRole: (value: Role) => void;
  createUser: (event: FormEvent<HTMLFormElement>) => Promise<void>;
}) => (
  <Card className="p-5 sm:p-6">
    <div className="mb-5 flex items-start gap-3">
      <span className="grid size-10 shrink-0 place-items-center rounded-[10px] bg-[var(--accent-soft)]">
        <Plus size={19} />
      </span>
      <div>
        <h2 className="text-xl font-bold tracking-[-0.03em]">สร้างบัญชีใหม่</h2>
        <p className="mt-1 text-sm text-[var(--ink-soft)]">
          ระบบจะสร้างรหัสผ่านชั่วคราวและบังคับให้เปลี่ยนเมื่อเข้าสู่ระบบครั้งแรก
        </p>
      </div>
    </div>
    <div id="create-user-error" className="mb-4" hidden={createError === null}>
      <Notice tone="danger">{createError}</Notice>
    </div>
    <form
      className="space-y-4"
      onSubmit={createUser}
      noValidate
      aria-busy={creating}
    >
      <label className="block text-sm font-semibold" htmlFor="create-user-name">
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
      <label className="block text-sm font-semibold" htmlFor="create-user-role">
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
);

export const UserFilters = ({
  draftEmailFilter,
  draftRoleFilter,
  draftEnabledFilter,
  setDraftEmailFilter,
  setDraftRoleFilter,
  setDraftEnabledFilter,
  applyFilters,
  resetFilters,
  rowActionsDisabled,
}: {
  draftEmailFilter: string;
  draftRoleFilter: FilterRole;
  draftEnabledFilter: FilterEnabled;
  setDraftEmailFilter: (value: string) => void;
  setDraftRoleFilter: (value: FilterRole) => void;
  setDraftEnabledFilter: (value: FilterEnabled) => void;
  applyFilters: (event: FormEvent<HTMLFormElement>) => void;
  resetFilters: () => void;
  rowActionsDisabled: boolean;
}) => (
  <Card className="p-5 sm:p-6">
    <div className="mb-5 flex items-start gap-3">
      <span className="grid size-10 shrink-0 place-items-center rounded-[10px] bg-[var(--success-soft)] text-[var(--success)]">
        <Search size={19} />
      </span>
      <div>
        <h2 className="text-xl font-bold tracking-[-0.03em]">ค้นหาและกรองบัญชี</h2>
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
      <label className="block text-sm font-semibold" htmlFor="user-role-filter">
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
);
