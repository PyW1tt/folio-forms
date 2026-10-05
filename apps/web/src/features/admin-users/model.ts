import { ApiError } from "@/lib/api";
import type { AdminUser, Role } from "@/lib/api";

export type FilterRole = Role | "all";
export type FilterEnabled = "all" | "true" | "false";
export type ConfirmKind =
  | "delete"
  | "disable"
  | "email"
  | "promote"
  | "demote"
  | "reset";
export type ActionKind = ConfirmKind | "enable";

export interface Confirmation {
  user: AdminUser;
  kind: ConfirmKind;
  email?: string;
}

export const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+$/u;
export const selectClass =
  "min-h-11 w-full rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-[var(--ink)] shadow-sm focus:border-[var(--ink)] focus:outline-none";

export const normalizeEmail = (value: string) => value.trim().toLowerCase();

export const actionsAreDisabled = (
  loading: boolean,
  creating: boolean,
  pendingAction: string | null,
  confirmation: Confirmation | null
): boolean =>
  loading || creating || pendingAction !== null || confirmation !== null;

export const resetIsDisabled = (
  actionsDisabled: boolean,
  targetUserId: string,
  authenticatedUserId: string | undefined
): boolean => actionsDisabled || targetUserId === authenticatedUserId;
export const credentialRequestIsInFlight = (
  creating: boolean,
  pendingAction: string | null
): boolean => creating || pendingAction?.endsWith(":reset") === true;
export const actionKey = (userId: string, kind: ActionKind) =>
  `${userId}:${kind}`;

export const roleBadgeLabels: Record<Role, string> = {
  admin: "ผู้ดูแลระบบ",
  user: "ผู้ใช้",
};
export const roleBadgeTones: Record<Role, "neutral" | "warning"> = {
  admin: "warning",
  user: "neutral",
};
export const errorMessageFor = (caughtError: unknown, fallback: string) => {
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

export const confirmationTitle = (kind: ConfirmKind) => {
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

export const confirmationDescription = (kind: ConfirmKind) => {
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

export const actionSuccessMessage = (kind: ActionKind) => {
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

export const actionFailureMessage = (kind: ActionKind) => {
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
