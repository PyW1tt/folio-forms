import {
  ChevronLeft,
  ChevronRight,
  KeyRound,
  Mail,
  Pencil,
  RefreshCw,
  ShieldCheck,
  Trash2,
  UserRound,
  UsersRound,
} from "lucide-react";
import type { FormEvent, RefObject } from "react";

import { Badge, Button, Card, Input, Notice, Spinner } from "@/components/ui";
import { formatDate } from "@/lib/api";
import type { AdminUser } from "@/lib/api";

import {
  actionKey,
  confirmationDescription,
  confirmationTitle,
  resetIsDisabled,
  roleBadgeLabels,
  roleBadgeTones,
} from "./model";
import type { Confirmation, ConfirmKind } from "./model";

export const UserList = ({
  users,
  listLoading,
  listError,
  nextCursor,
  cursorHistory,
  usersHeadingRef,
  rowActionsDisabled,
  legacyAccountLinkLoading,
  legacyAccountLinkAction,
  confirmingAction,
  editingEmailId,
  editingEmail,
  emailEditError,
  pendingAction,
  authenticatedUserId,
  reload,
  retry,
  submitEmailEdit,
  changeEditingEmail,
  cancelEmailEdit,
  confirmMutation,
  cancelConfirmation,
  requestConfirmation,
  beginEmailEdit,
  enableUser,
  goToPreviousPage,
  goToNextPage,
}: {
  users: AdminUser[];
  listLoading: boolean;
  listError: string | null;
  nextCursor: string | null;
  cursorHistory: (string | null)[];
  usersHeadingRef: RefObject<HTMLHeadingElement | null>;
  rowActionsDisabled: boolean;
  legacyAccountLinkLoading: boolean;
  legacyAccountLinkAction: string | null;
  confirmingAction: Confirmation | null;
  editingEmailId: string | null;
  editingEmail: string;
  emailEditError: string | null;
  pendingAction: string | null;
  authenticatedUserId: string | undefined;
  reload: () => void;
  retry: () => void;
  submitEmailEdit: (event: FormEvent<HTMLFormElement>, user: AdminUser) => void;
  changeEditingEmail: (value: string) => void;
  cancelEmailEdit: (user: AdminUser) => void;
  confirmMutation: (
    user: AdminUser,
    kind: ConfirmKind,
    nextEmail?: string
  ) => void;
  cancelConfirmation: (user: AdminUser, kind: ConfirmKind) => void;
  requestConfirmation: (user: AdminUser, kind: ConfirmKind) => void;
  beginEmailEdit: (user: AdminUser) => void;
  enableUser: (user: AdminUser) => void;
  goToPreviousPage: () => void;
  goToNextPage: () => void;
}) => (
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
        onClick={reload}
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
              onClick={retry}
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
                confirmingAction?.user.id === user.id ? confirmingAction : null;
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
                            onSubmit={(event) => submitEmailEdit(event, user)}
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
                              onChange={(event) =>
                                changeEditingEmail(event.target.value)
                              }
                              aria-describedby={
                                emailEditError
                                  ? `email-edit-error-${user.id}`
                                  : undefined
                              }
                              aria-invalid={emailEditError ? "true" : undefined}
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
                                {rowPending ? <Spinner /> : <Mail size={14} />}
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
                        aria-label={confirmationTitle(confirmationForRow.kind)}
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
                            {confirmationDescription(confirmationForRow.kind)}
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
                              confirmMutation(
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
                            onClick={() =>
                              cancelConfirmation(user, confirmationForRow.kind)
                            }
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
                            data-user-action-id={actionKey(user.id, "disable")}
                            onClick={() => requestConfirmation(user, "disable")}
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
                            data-user-action-id={actionKey(user.id, "enable")}
                            onClick={() => enableUser(user)}
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
                            data-user-action-id={actionKey(user.id, "demote")}
                            onClick={() => requestConfirmation(user, "demote")}
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
                            data-user-action-id={actionKey(user.id, "promote")}
                            onClick={() => requestConfirmation(user, "promote")}
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
                            data-user-action-id={actionKey(user.id, "email")}
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
                            authenticatedUserId
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
                          onClick={() => requestConfirmation(user, "delete")}
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
);
