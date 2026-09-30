// oxlint-disable unicorn/filename-case -- TanStack Router requires this dynamic route filename.
// oxlint-disable complexity -- Form editor coordinates loading, publishing, and editor state.
import {
  createFileRoute,
  Link,
  useNavigate,
  useParams,
} from "@tanstack/react-router";
import {
  ArrowLeft,
  Check,
  Copy,
  ExternalLink,
  FileText,
  Globe2,
  Save,
  Send,
} from "lucide-react";
import type { FormEvent } from "react";
import { useEffect, useRef, useState } from "react";

import type {
  EditorBridgeMessage,
  OnlyOfficeEditorState,
} from "@/components/onlyoffice-editor";
import { OnlyOfficeEditor } from "@/components/onlyoffice-editor";
import {
  Badge,
  Button,
  Input,
  Notice,
  Spinner,
  Textarea,
} from "@/components/ui";
import {
  ApiError,
  apiGet,
  apiPatch,
  apiPost,
  waitForOperation,
} from "@/lib/api";
import type { FillMethod, FormDetail, FormSummary } from "@/lib/api";

interface FormDetailResponse {
  editorConfigUrl: string;
  form: FormSummary;
}

interface FormMutationResponse {
  form: FormSummary;
}

interface AdminEditorConfig {
  bridge?: {
    capabilities?: Partial<Record<"publish" | "save-template", string>>;
  };
  config?: {
    document?: {
      key?: unknown;
    };
  };
}

type OperationViewStatus = "pending" | "processing" | "completed" | "failed";

const formStatusDetails = {
  archived: { label: "เก็บถาวร", tone: "neutral" as const },
  draft: { label: "ร่าง", tone: "warning" as const },
  published: { label: "เผยแพร่แล้ว", tone: "success" as const },
};

const detailErrorMessage = (caughtError: unknown, fallback: string): string => {
  if (caughtError instanceof ApiError) {
    switch (caughtError.code) {
      case "editor_in_use": {
        return "เอกสารนี้กำลังถูกแก้ไขโดยผู้ใช้รายอื่น กรุณารอแล้วลองใหม่";
      }
      case "document_unavailable": {
        return "ยังไม่มีเอกสารต้นแบบสำหรับดำเนินการ กรุณาตรวจสอบต้นแบบของแบบฟอร์ม";
      }
      case "stale_document": {
        return "เอกสารมีการเปลี่ยนแปลงแล้ว กรุณาลองใหม่เพื่อใช้เอกสารปัจจุบัน";
      }
      case "editor_lease_inactive": {
        return "เซสชันตัวแก้ไขหมดอายุ กรุณาลองใหม่";
      }
      case "operation_in_progress": {
        return "มีการบันทึกเอกสารกำลังดำเนินการ กรุณารอแล้วลองใหม่";
      }
      case "form_not_published": {
        return "เก็บถาวรหรือยกเลิกเก็บถาวรได้เฉพาะแบบฟอร์มที่เผยแพร่แล้ว";
      }
      case "native_fill_unsupported": {
        return "Native ใช้ได้เมื่อ Field ที่เผยแพร่ทุกช่องเป็น Text, Checkbox, Date, Dropdown หรือ Combo และไม่มี Picture";
      }
      case "form_unavailable": {
        return "แบบฟอร์มนี้ยังไม่พร้อมรับคำตอบใหม่";
      }
      case "invalid_request": {
        return "ข้อมูลชื่อหรือคำอธิบายไม่ถูกต้อง กรุณาตรวจสอบความยาวแล้วลองใหม่";
      }
      case "not_found": {
        return "ไม่พบแบบฟอร์มนี้ อาจถูกลบไปแล้ว";
      }
      case "unauthorized": {
        return "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่";
      }
      case "forbidden": {
        return "คุณไม่มีสิทธิ์ดำเนินการนี้";
      }
      case "password_change_required": {
        return "กรุณาเปลี่ยนรหัสผ่านก่อนแก้ไขแบบฟอร์ม";
      }
      case "published_immutable": {
        return "แบบฟอร์มนี้เผยแพร่แล้ว สัญญาเอกสารและการตั้งค่า Field ไม่สามารถแก้ไขในที่เดิมได้";
      }
      case "internal_error": {
        return "ระบบไม่สามารถดำเนินการได้ กรุณาลองใหม่อีกครั้ง";
      }
      default: {
        break;
      }
    }
  }
  return fallback;
};

const loadFormDetail = async (publicId: string): Promise<FormDetail> => {
  const payload = await apiGet<FormDetailResponse>(
    `/api/admin/forms/${publicId}`
  );
  return {
    ...payload.form,
    editorConfigUrl: payload.editorConfigUrl,
  };
};

const FormEditorRoute = () => {
  const { formId: publicId } = useParams({
    from: "/admin/forms/$formId",
  });
  const navigate = useNavigate();
  const [form, setForm] = useState<FormDetail | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [metadataBusy, setMetadataBusy] = useState<
    "archive" | "fill-method" | "save" | "duplicate" | null
  >(null);
  const [busy, setBusy] = useState<"save" | "publish" | null>(null);
  const [title, setTitle] = useState("");
  const [description, setDescription] = useState("");
  const [copied, setCopied] = useState(false);
  const [publishPrompt, setPublishPrompt] = useState(false);
  const [editorState, setEditorState] =
    useState<OnlyOfficeEditorState>("loading");
  const [operationStatus, setOperationStatus] =
    useState<OperationViewStatus | null>(null);
  const [editorRevision, setEditorRevision] = useState(0);
  const [reloadToken, setReloadToken] = useState(0);
  const busyGuardRef = useRef(false);
  const metadataGuardRef = useRef(false);
  const feedbackRef = useRef<HTMLDivElement>(null);
  const restorePublishFocusRef = useRef(false);
  useEffect(() => {
    if (!publishPrompt) {
      if (restorePublishFocusRef.current) {
        restorePublishFocusRef.current = false;
        document.querySelector<HTMLElement>("#publish-trigger")?.focus();
      }
      return;
    }
    document.querySelector<HTMLElement>("#publish-cancel")?.focus();
    const handleEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        restorePublishFocusRef.current = true;
        setPublishPrompt(false);
      }
    };
    window.addEventListener("keydown", handleEscape);
    return () => window.removeEventListener("keydown", handleEscape);
  }, [publishPrompt]);

  useEffect(() => {
    if (error || notice || operationStatus === "pending") {
      feedbackRef.current?.focus();
    }
  }, [error, notice, operationStatus]);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setForm(null);
    setTitle("");
    setDescription("");
    setError(null);
    setNotice(null);
    setOperationStatus(null);
    setMetadataBusy(null);
    setEditorState("loading");

    const loadForm = async () => {
      try {
        const payload = await loadFormDetail(publicId);
        if (!cancelled) {
          setForm(payload);
          setTitle(payload.title);
          setDescription(payload.description);
        }
      } catch (caughtError) {
        if (!cancelled) {
          setError(
            detailErrorMessage(
              caughtError,
              "ไม่สามารถโหลดแบบฟอร์มได้ กรุณาลองใหม่อีกครั้ง"
            )
          );
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    void loadForm();
    return () => {
      cancelled = true;
    };
  }, [publicId, reloadToken]);

  if (loading) {
    return (
      <div
        className="grid min-h-64 place-items-center gap-3"
        aria-busy="true"
        role="status"
      >
        <Spinner />
        <span className="text-sm text-[var(--ink-soft)]">กำลังโหลดแบบฟอร์ม…</span>
      </div>
    );
  }

  const loadedForm = form;
  if (!loadedForm) {
    return (
      <div ref={feedbackRef} className="space-y-4" tabIndex={-1}>
        <Notice tone="danger">{error ?? "ไม่พบแบบฟอร์มนี้"}</Notice>
        <Button
          variant="secondary"
          type="button"
          onClick={() => setReloadToken((value) => value + 1)}
        >
          ลองโหลดอีกครั้ง
        </Button>
      </div>
    );
  }

  const canEditTemplate = loadedForm.status === "draft";
  const shareUrl = `${window.location.origin}/forms/${publicId}/fill`;
  const { activeDraftCount } = loadedForm;
  const { editorConfigUrl } = loadedForm;
  const operationBusy =
    operationStatus === "pending" || operationStatus === "processing";
  const canAct =
    canEditTemplate &&
    editorState === "ready" &&
    !busy &&
    !metadataBusy &&
    !operationBusy;
  const metadataCanAct = !busy && !metadataBusy && !operationBusy;
  const canDuplicate =
    loadedForm.status === "draft" || loadedForm.status === "published";

  const perform = async (action: "save" | "publish") => {
    if (!canAct || busyGuardRef.current || !editorConfigUrl) {
      return;
    }
    busyGuardRef.current = true;
    setBusy(action);
    setOperationStatus("pending");
    setError(null);
    setNotice(null);
    try {
      const editorConfig = await apiGet<AdminEditorConfig>(editorConfigUrl);
      const documentKey = editorConfig.config?.document?.key;
      if (typeof documentKey !== "string" || !documentKey) {
        throw new Error("document_key_unavailable");
      }
      const capabilityAction = action === "save" ? "save-template" : "publish";
      const capability = editorConfig.bridge?.capabilities?.[capabilityAction];
      if (!capability) {
        throw new Error("editor_capability_unavailable");
      }
      const result = await apiPost<{ operationId?: string }>(
        `/api/admin/forms/${publicId}/${action}`,
        { documentKey },
        capability
      );
      if (result.operationId) {
        await waitForOperation(result.operationId, (operation) => {
          setOperationStatus(operation.status);
        });
      }
      if (action === "save") {
        setEditorState("loading");
        setEditorRevision((value) => value + 1);
      }
      setOperationStatus("completed");
      setNotice(
        action === "publish"
          ? "เผยแพร่แบบฟอร์มแล้ว คำตอบใหม่จะใช้เอกสารฉบับนี้"
          : "บันทึกแบบร่างเรียบร้อยแล้ว"
      );
      try {
        setForm(await loadFormDetail(publicId));
      } catch (refreshError) {
        setError(
          detailErrorMessage(
            refreshError,
            "ดำเนินการสำเร็จแล้ว แต่โหลดข้อมูลล่าสุดไม่สำเร็จ กรุณาลองใหม่"
          )
        );
      }
    } catch (caughtError) {
      setOperationStatus("failed");
      setError(
        detailErrorMessage(
          caughtError,
          action === "publish"
            ? "เผยแพร่แบบฟอร์มไม่สำเร็จ กรุณาลองใหม่"
            : "บันทึกแบบร่างไม่สำเร็จ กรุณาลองใหม่"
        )
      );
    } finally {
      setBusy(null);
      busyGuardRef.current = false;
    }
  };
  const requestAction = (action: "save" | "publish") => {
    if (!canAct || busyGuardRef.current) {
      return;
    }
    if (action === "publish" && activeDraftCount > 0) {
      setPublishPrompt(true);
      return;
    }
    void perform(action);
  };

  const updateMetadata = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!metadataCanAct || metadataGuardRef.current) {
      return;
    }
    const nextTitle = title.trim();
    if (!nextTitle) {
      setError("กรุณากรอกชื่อแบบฟอร์ม");
      setNotice(null);
      return;
    }
    metadataGuardRef.current = true;
    setMetadataBusy("save");
    setError(null);
    setNotice(null);
    setOperationStatus(null);
    try {
      const payload = await apiPatch<FormMutationResponse>(
        `/api/admin/forms/${publicId}`,
        {
          description: description.trim() || null,
          title: nextTitle,
        }
      );
      setForm((currentForm) =>
        currentForm ? { ...currentForm, ...payload.form } : currentForm
      );
      setTitle(payload.form.title);
      setDescription(payload.form.description);
      setNotice("บันทึกข้อมูลแบบฟอร์มเรียบร้อยแล้ว");
    } catch (caughtError) {
      setError(
        detailErrorMessage(
          caughtError,
          "บันทึกข้อมูลแบบฟอร์มไม่สำเร็จ กรุณาลองใหม่อีกครั้ง"
        )
      );
    } finally {
      setMetadataBusy(null);
      metadataGuardRef.current = false;
    }
  };
  const updateFillMethod = async (fillMethod: FillMethod) => {
    if (
      !metadataCanAct ||
      metadataGuardRef.current ||
      fillMethod === loadedForm.fillMethod
    ) {
      return;
    }
    metadataGuardRef.current = true;
    setMetadataBusy("fill-method");
    setError(null);
    setNotice(null);
    setOperationStatus(null);
    try {
      const payload = await apiPatch<FormMutationResponse>(
        `/api/admin/forms/${publicId}`,
        { fillMethod }
      );
      setForm((currentForm) =>
        currentForm ? { ...currentForm, ...payload.form } : currentForm
      );
      setNotice(
        fillMethod === "native"
          ? "เปลี่ยนเป็นแบบฟอร์ม Native แล้ว"
          : "เปลี่ยนเป็น ONLYOFFICE แล้ว"
      );
    } catch (caughtError) {
      setError(
        detailErrorMessage(
          caughtError,
          "เปลี่ยนวิธีกรอกแบบฟอร์มไม่สำเร็จ กรุณาลองใหม่อีกครั้ง"
        )
      );
    } finally {
      setMetadataBusy(null);
      metadataGuardRef.current = false;
    }
  };
  const toggleArchive = async () => {
    if (
      !metadataCanAct ||
      metadataGuardRef.current ||
      (loadedForm.status !== "published" && loadedForm.status !== "archived")
    ) {
      return;
    }
    const nextStatus =
      loadedForm.status === "archived" ? "published" : "archived";
    metadataGuardRef.current = true;
    setMetadataBusy("archive");
    setError(null);
    setNotice(null);
    setOperationStatus(null);
    try {
      const payload = await apiPatch<FormMutationResponse>(
        `/api/admin/forms/${publicId}`,
        { status: nextStatus }
      );
      setForm((currentForm) =>
        currentForm ? { ...currentForm, ...payload.form } : currentForm
      );
      setNotice(
        nextStatus === "archived"
          ? "เก็บแบบฟอร์มเรียบร้อยแล้ว คำตอบเดิมยังเปิดได้"
          : "ยกเลิกเก็บถาวรเรียบร้อยแล้ว แบบฟอร์มรับคำตอบใหม่ได้"
      );
    } catch (caughtError) {
      setError(
        detailErrorMessage(
          caughtError,
          nextStatus === "archived"
            ? "เก็บแบบฟอร์มไม่สำเร็จ กรุณาลองใหม่อีกครั้ง"
            : "ยกเลิกเก็บถาวรไม่สำเร็จ กรุณาลองใหม่อีกครั้ง"
        )
      );
    } finally {
      setMetadataBusy(null);
      metadataGuardRef.current = false;
    }
  };

  const duplicateForm = async () => {
    if (!canDuplicate || !metadataCanAct || metadataGuardRef.current) {
      return;
    }
    metadataGuardRef.current = true;
    setMetadataBusy("duplicate");
    setError(null);
    setNotice(null);
    setOperationStatus(null);
    try {
      const payload = await apiPost<FormMutationResponse>(
        `/api/admin/forms/${publicId}/duplicate`,
        {}
      );
      navigate({
        params: { formId: payload.form.publicId },
        to: "/admin/forms/$formId",
      });
    } catch (caughtError) {
      setError(
        detailErrorMessage(
          caughtError,
          "สร้างสำเนาแบบฟอร์มไม่สำเร็จ กรุณาลองใหม่อีกครั้ง"
        )
      );
    } finally {
      setMetadataBusy(null);
      metadataGuardRef.current = false;
    }
  };

  const handleEditorBridgeMessage = async (message: EditorBridgeMessage) => {
    if (
      message.type !== "operation" ||
      (message.action !== "save-template" && message.action !== "publish")
    ) {
      return;
    }
    setOperationStatus(
      message.status === "pending" ? "pending" : message.status
    );
    if (message.status === "failed") {
      setError(
        message.action === "publish"
          ? "เผยแพร่แบบฟอร์มไม่สำเร็จ กรุณาลองใหม่"
          : "บันทึกแบบร่างไม่สำเร็จ กรุณาลองใหม่"
      );
      setNotice(null);
      return;
    }
    if (message.status !== "completed") {
      return;
    }
    if (message.action === "save-template") {
      setEditorState("loading");
      setEditorRevision((value) => value + 1);
    }
    setOperationStatus("completed");
    setError(null);
    setNotice(
      message.action === "publish"
        ? "เผยแพร่แบบฟอร์มแล้ว คำตอบใหม่จะใช้เอกสารฉบับนี้"
        : "บันทึกแบบร่างเรียบร้อยแล้ว"
    );
    try {
      setForm(await loadFormDetail(publicId));
    } catch (refreshError) {
      setError(
        detailErrorMessage(
          refreshError,
          "ดำเนินการสำเร็จแล้ว แต่โหลดข้อมูลล่าสุดไม่สำเร็จ กรุณาลองใหม่"
        )
      );
    }
  };

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(shareUrl);
      setCopied(true);
      setError(null);
      setNotice("คัดลอกลิงก์แบบฟอร์มแล้ว");
      window.setTimeout(() => setCopied(false), 1800);
    } catch {
      setError("คัดลอกลิงก์ไม่สำเร็จ กรุณาลองใหม่");
      setNotice(null);
    }
  };
  const archiveLabel =
    loadedForm.status === "archived" ? "ยกเลิกเก็บถาวร" : "เก็บแบบฟอร์ม";
  const archiveButtonLabel =
    metadataBusy === "archive" ? "กำลังเปลี่ยนสถานะ…" : archiveLabel;

  const status = formStatusDetails[loadedForm.status];

  return (
    <div>
      <div className="mb-7 flex flex-wrap items-center justify-between gap-4">
        <button
          className="inline-flex items-center gap-2 text-sm font-semibold text-[var(--ink-soft)] hover:text-[var(--ink)]"
          type="button"
          onClick={() => navigate({ to: "/admin" })}
        >
          <ArrowLeft size={15} />
          กลับไปยังรายการแบบฟอร์ม
        </button>
        <div className="flex flex-wrap gap-2">
          <Button
            variant="secondary"
            size="sm"
            type="button"
            onClick={duplicateForm}
            disabled={!canDuplicate || !metadataCanAct}
          >
            {metadataBusy === "duplicate" ? <Spinner /> : <Copy size={15} />}
            {metadataBusy === "duplicate" ? "กำลังสร้างสำเนา…" : "สร้างสำเนา"}
          </Button>
          {loadedForm.status === "published" ||
          loadedForm.status === "archived" ? (
            <Button
              variant="secondary"
              size="sm"
              type="button"
              onClick={toggleArchive}
              disabled={!metadataCanAct}
              aria-label={
                loadedForm.status === "archived"
                  ? "ยกเลิกเก็บถาวรแบบฟอร์ม"
                  : "เก็บแบบฟอร์ม"
              }
            >
              {metadataBusy === "archive" ? <Spinner /> : null}
              {archiveButtonLabel}
            </Button>
          ) : null}
          <Button
            variant="secondary"
            size="sm"
            type="button"
            onClick={() => requestAction("save")}
            disabled={!canAct}
          >
            {busy === "save" ? <Spinner /> : <Save size={15} />}
            {busy === "save" ? "กำลังบันทึก…" : "บันทึกแบบร่าง"}
          </Button>
          <Button
            id="publish-trigger"
            size="sm"
            type="button"
            onClick={() => requestAction("publish")}
            disabled={!canAct}
          >
            {busy === "publish" ? <Spinner /> : <Send size={15} />}
            {busy === "publish" ? "กำลังเผยแพร่…" : "เผยแพร่"}
          </Button>
        </div>
      </div>
      {publishPrompt ? (
        <div
          className="mb-5 flex flex-wrap items-center justify-between gap-4 rounded-[var(--radius)] border border-[var(--warning)] bg-[var(--warning-soft)] p-4"
          aria-labelledby="publish-warning-title"
          role="alertdialog"
        >
          <div>
            <p id="publish-warning-title" className="font-semibold">
              การเผยแพร่จะยกเลิกฉบับร่างที่กำลังแก้ไข {activeDraftCount} รายการ
            </p>
            <p className="mt-1 text-sm text-[var(--ink-soft)]">
              คำตอบที่ส่งแล้วจะไม่เปลี่ยนแปลง
            </p>
          </div>
          <div className="flex gap-2">
            <Button
              id="publish-cancel"
              variant="secondary"
              size="sm"
              type="button"
              onClick={() => {
                restorePublishFocusRef.current = true;
                setPublishPrompt(false);
              }}
            >
              ยกเลิก
            </Button>
            <Button
              size="sm"
              type="button"
              onClick={() => {
                setPublishPrompt(false);
                void perform("publish");
              }}
            >
              เผยแพร่ต่อ
            </Button>
          </div>
        </div>
      ) : null}
      <div className="mb-7 flex flex-col justify-between gap-4 sm:flex-row sm:items-end">
        <div>
          <div className="mb-2 flex items-center gap-2">
            <Badge tone={status.tone}>{status.label}</Badge>
            <span className="text-sm text-[var(--ink-soft)]">
              ตัวแก้ไขสำหรับผู้ดูแล
            </span>
          </div>
          <h1 className="text-3xl font-bold tracking-[-0.04em]">
            {loadedForm.title}
          </h1>
          <p className="mt-2 max-w-2xl text-[var(--ink-soft)]">
            {loadedForm.description || "จัดรูปแบบเอกสาร แล้วเผยแพร่เมื่อพร้อม"}
          </p>
        </div>
        <Link
          to="/admin/results"
          search={{ form: publicId }}
          className="inline-flex items-center gap-2 text-sm font-semibold underline underline-offset-4"
        >
          ดูคำตอบ <ExternalLink size={15} />
        </Link>
      </div>
      {error || notice || operationBusy ? (
        <div ref={feedbackRef} className="mb-4 space-y-2" tabIndex={-1}>
          {error ? <Notice tone="danger">{error}</Notice> : null}
          {notice ? <Notice tone="success">{notice}</Notice> : null}
          {operationBusy ? (
            <Notice>
              <span className="inline-flex items-center gap-2">
                <Spinner />
                {operationStatus === "processing"
                  ? "กำลังประมวลผลเอกสาร…"
                  : "กำลังบันทึกการเปลี่ยนแปลง…"}
              </span>
            </Notice>
          ) : null}
        </div>
      ) : null}
      {canEditTemplate ? null : (
        <Notice>
          {loadedForm.status === "archived"
            ? "แบบฟอร์มนี้เก็บถาวรแล้ว คำตอบเดิมยังเปิดดูและดำเนินการต่อได้"
            : "แบบฟอร์มนี้เผยแพร่แล้ว สัญญาเอกสารและการตั้งค่า Field ไม่สามารถแก้ไขในที่เดิมได้ หากต้องการเปลี่ยนโครงสร้างให้สร้าง Form ใหม่"}
        </Notice>
      )}
      <section
        className="mb-4 rounded-[var(--radius)] border border-[var(--line)] bg-[var(--paper)] p-5"
        aria-labelledby="metadata-title"
      >
        <div className="max-w-2xl">
          <h2 id="metadata-title" className="text-lg font-semibold">
            ข้อมูลแบบฟอร์ม
          </h2>
          <p className="mt-1 text-sm text-[var(--ink-soft)]">
            แก้ไขชื่อและคำอธิบายได้ทั้งแบบร่างและแบบฟอร์มที่เผยแพร่แล้ว
          </p>
          <form
            className="mt-4 space-y-4"
            onSubmit={updateMetadata}
            aria-busy={metadataBusy === "save"}
          >
            <div className="space-y-2">
              <label className="text-sm font-semibold" htmlFor="form-title">
                ชื่อแบบฟอร์ม
              </label>
              <Input
                id="form-title"
                name="title"
                value={title}
                onChange={(event) => setTitle(event.target.value)}
                maxLength={200}
                required
                aria-describedby="form-title-help"
                disabled={!metadataCanAct}
              />
              <p
                id="form-title-help"
                className="text-xs text-[var(--ink-soft)]"
              >
                ต้องระบุชื่อ ความยาวไม่เกิน 200 ตัวอักษร
              </p>
            </div>
            <div className="space-y-2">
              <label
                className="text-sm font-semibold"
                htmlFor="form-description"
              >
                คำอธิบาย
              </label>
              <Textarea
                id="form-description"
                name="description"
                value={description}
                onChange={(event) => setDescription(event.target.value)}
                maxLength={2000}
                rows={4}
                aria-describedby="form-description-help"
                disabled={!metadataCanAct}
              />
              <p
                id="form-description-help"
                className="text-xs text-[var(--ink-soft)]"
              >
                ใส่คำอธิบายเพิ่มเติมได้ไม่เกิน 2,000 ตัวอักษร
              </p>
            </div>
            <div className="flex justify-end">
              <Button type="submit" disabled={!metadataCanAct}>
                {metadataBusy === "save" ? <Spinner /> : <Save size={15} />}
                {metadataBusy === "save" ? "กำลังบันทึกข้อมูล…" : "บันทึกข้อมูลแบบฟอร์ม"}
              </Button>
            </div>
          </form>
          <div
            className="mt-5 space-y-2"
            aria-busy={metadataBusy === "fill-method"}
          >
            <label
              className="text-sm font-semibold"
              htmlFor="form-fill-method"
            >
              Fill Method
            </label>
            <select
              className="min-h-11 w-full rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-[var(--ink)] shadow-sm focus:border-[var(--ink)] focus:outline-none"
              disabled={!metadataCanAct || loadedForm.status === "draft"}
              id="form-fill-method"
              onChange={(event) => {
                const value = event.target.value;
                if (value === "native" || value === "onlyoffice") {
                  void updateFillMethod(value);
                }
              }}
              value={loadedForm.fillMethod}
            >
              <option value="onlyoffice">ONLYOFFICE</option>
              <option
                disabled={!loadedForm.nativeFillAvailable}
                value="native"
              >
                Native form
              </option>
            </select>
            <p
              className="text-xs text-[var(--ink-soft)]"
              id="form-fill-method-help"
            >
              {loadedForm.status === "draft"
                ? "Fill Method can be changed after publishing."
                : loadedForm.nativeFillAvailable
                  ? "This method applies to new and existing Draft responses."
                  : "Native supports text, checkbox, date, dropdown, and combo Fields; Pictures remain ONLYOFFICE-only."}
            </p>
            {metadataBusy === "fill-method" ? (
              <p className="inline-flex items-center gap-2 text-xs" role="status">
                <Spinner />
                Updating Fill Method…
              </p>
            ) : null}
          </div>
        </div>
      </section>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3 rounded-[var(--radius)] border border-[var(--line)] bg-[var(--paper)] p-4">
        <div className="flex min-w-0 items-center gap-3">
          <Globe2 className="shrink-0 text-[var(--success)]" size={18} />
          <div className="min-w-0">
            <p className="text-sm font-semibold">ลิงก์แบบฟอร์ม</p>
            <p className="truncate text-xs text-[var(--ink-soft)]">
              {shareUrl}
            </p>
          </div>
        </div>
        <Button
          variant="secondary"
          size="sm"
          type="button"
          onClick={copyLink}
          disabled={loadedForm.status !== "published"}
        >
          {copied ? <Check size={15} /> : <Copy size={15} />}
          {copied ? "คัดลอกแล้ว" : "คัดลอกลิงก์"}
        </Button>
      </div>
      <div className="mb-4 flex items-center gap-2 text-sm text-[var(--ink-soft)]">
        <FileText size={16} />
        ตัวแก้ไขเอกสาร
        <span className="text-xs">· แนะนำให้ใช้คอมพิวเตอร์</span>
      </div>
      {canEditTemplate ? (
        <div className="overflow-hidden rounded-[var(--radius)] border border-[var(--line-strong)] bg-[var(--muted)] shadow-inner">
          <OnlyOfficeEditor
            key={publicId}
            onBridgeMessage={handleEditorBridgeMessage}
            onStateChange={setEditorState}
            configUrl={editorConfigUrl}
            revision={editorRevision}
            title={`ตัวแก้ไขเอกสาร ${loadedForm.title}`}
          />
        </div>
      ) : null}
    </div>
  );
};

export const Route = createFileRoute("/admin/forms/$formId/")({
  component: FormEditorRoute,
});
