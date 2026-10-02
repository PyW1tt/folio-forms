// oxlint-disable unicorn/filename-case -- TanStack Router requires this dynamic route filename.
import {
  createFileRoute,
  Navigate,
  useBlocker,
  useNavigate,
  useParams,
  useSearch,
} from "@tanstack/react-router";
import { ArrowLeft, CheckCircle2, Monitor } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import { LegacySsoSwitchButton } from "@/components/legacy-sso-switch-button";
import { NativeForm } from "@/components/native-form";
import type { NativeFormField } from "@/components/native-form";
import { OnlyOfficeEditor } from "@/components/onlyoffice-editor";
import type { EditorBridgeMessage } from "@/components/onlyoffice-editor";
import { Button, Notice, Spinner } from "@/components/ui";
import {
  ApiError,
  apiDelete,
  apiGet,
  apiPost,
  apiPostFormData,
  downloadArtifact,
  safeReturnPath,
  waitForOperation,
} from "@/lib/api";
import type { FillMethod, Operation } from "@/lib/api";
import { roleFor, useAuth } from "@/lib/auth";
import {
  createDeferred,
  isSaveFlowBusy,
  saveThenDownload,
  shouldBlockDirtyNavigation,
} from "@/lib/form-lifecycle";

const formRequestError = (error: unknown, fallback: string) => {
  if (error instanceof ApiError) {
    if (error.code === "handoff_unavailable") {
      return "ลิงก์เปิดแบบฟอร์มหมดอายุหรือใช้ไม่ได้ กรุณากลับไปยังระบบต้นทางแล้วลองใหม่";
    }
    if (error.code === "prefill_required") {
      return "แบบฟอร์มนี้ต้องเปิดจากระบบต้นทาง กรุณากลับไปยังระบบต้นทางแล้วลองใหม่";
    }
    if (error.status === 401) {
      return "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่";
    }
    if (error.code === "form_unavailable") {
      return "แบบฟอร์มนี้เก็บถาวรแล้วและยังไม่รับคำตอบใหม่";
    }
    if (error.code === "fill_method_changed") {
      return "วิธีกรอกแบบฟอร์มเปลี่ยนแล้ว กรุณาโหลดแบบฟอร์มใหม่ก่อนบันทึก";
    }
  }
  return fallback;
};

const isHandoffError = (error: unknown): error is ApiError =>
  error instanceof ApiError &&
  (error.code === "handoff_unavailable" || error.code === "prefill_required");

const returnToSource = () => {
  if (typeof window === "undefined") {
    return;
  }
  if (window.history.length > 1) {
    window.history.back();
    return;
  }
  window.location.assign("/");
};
type ExitIntent = "dashboard" | "navigation" | "reauth";

interface ReauthenticationRequest {
  handled: boolean;
  resolve: (allowed: boolean) => void;
}
type ExportFormat = "docx" | "pdf";
class DraftSaveError extends Error {
  constructor() {
    super("Draft save failed");
    this.name = "DraftSaveError";
  }
}

interface PublicForm {
  title: string;
  description?: string;
  fillMethod: FillMethod;
}

interface NativeEditorConfig {
  capabilities: Record<"save-draft" | "submit", string>;
  data: Record<string, unknown>;
  documentKey: string;
  fields: NativeFormField[];
  fillMethod: "native";
  lockedFields: Record<string, boolean>;
  pictures: Record<string, boolean>;
  responseId: string;
}

const nativeValuesFromConfig = (
  config: NativeEditorConfig
): Record<string, unknown> => {
  const values: Record<string, unknown> = {};
  for (const field of config.fields) {
    const value = config.data[field.tag];
    switch (field.type) {
      case "checkbox": {
        values[field.tag] = value === true;
        break;
      }
      case "combo":
      case "dropdown": {
        values[field.tag] = typeof value === "string" ? value : null;
        break;
      }
      case "picture": {
        values[field.tag] = "";
        break;
      }
      default: {
        values[field.tag] = typeof value === "string" ? value : "";
      }
    }
  }
  return values;
};

const draftSaveSuccessMessage = (
  exportFormat: ExportFormat | null | undefined
): string => {
  if (exportFormat === "docx") {
    return "บันทึกและดาวน์โหลด DOCX แล้ว";
  }
  if (exportFormat === "pdf") {
    return "บันทึกและดาวน์โหลด PDF แล้ว";
  }
  return "บันทึกฉบับร่างคำตอบแล้ว";
};

const nativeSaveErrorMessage = (
  action: "save-draft" | "submit",
  saved: boolean,
  exportFormat: ExportFormat | null | undefined
): string => {
  if (action === "submit") {
    return "ส่งแบบฟอร์มไม่สำเร็จ กรุณาลองใหม่";
  }
  if (!saved) {
    return "บันทึกฉบับร่างไม่สำเร็จ กรุณาลองใหม่";
  }
  if (exportFormat) {
    return "บันทึกแล้ว แต่ดาวน์โหลดไฟล์ไม่สำเร็จ กรุณาลองใหม่";
  }
  return "บันทึกแล้ว แต่โหลดคำตอบล่าสุดไม่สำเร็จ กรุณาโหลดแบบฟอร์มใหม่";
};

const downloadDraftArtifact = (
  responseId: string,
  format: ExportFormat
): Promise<void> =>
  downloadArtifact(
    `/api/responses/${responseId}/draft/${format}`,
    `response-${responseId}.${format}`
  );
// oxlint-disable-next-line complexity -- Coordinates the public form editor, draft lifecycle, and exit confirmation.
const FillRoute = () => {
  const { publicId } = useParams({ from: "/forms/$publicId/fill" });
  const { responseId } = useSearch({ from: "/forms/$publicId/fill" });
  const { user, loading: authLoading } = useAuth();
  const navigate = useNavigate();
  const [form, setForm] = useState<PublicForm | null>(null);
  const [editorConfigUrl, setEditorConfigUrl] = useState<string | null>(null);
  const [nativeConfig, setNativeConfig] = useState<NativeEditorConfig | null>(
    null
  );
  const [nativeValues, setNativeValues] = useState<Record<string, unknown>>({});
  const [nativePictureFiles, setNativePictureFiles] = useState<
    Record<string, File>
  >({});
  const [pictureInputKey, setPictureInputKey] = useState(0);
  const [nativeSaving, setNativeSaving] = useState(false);
  const [editorRevision, setEditorRevision] = useState(0);
  const [activeResponseId, setActiveResponseId] = useState(responseId);
  const [loading, setLoading] = useState(true);
  const [startBusy, setStartBusy] = useState(false);
  const [startAttempt, setStartAttempt] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [handoffError, setHandoffError] = useState(false);
  const [operationError, setOperationError] = useState<string | null>(null);
  const [operation, setOperation] = useState<Operation | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [dirty, setDirty] = useState(false);
  const [exitIntent, setExitIntent] = useState<ExitIntent | null>(null);
  const [saveBeforeExit, setSaveBeforeExit] = useState(false);
  const [saveRequest, setSaveRequest] = useState(0);
  const [exportAfterSave, setExportAfterSave] = useState<ExportFormat | null>(
    null
  );
  const [clearDirtyRequest, setClearDirtyRequest] = useState(0);
  const [discardBusy, setDiscardBusy] = useState(false);
  const startErrorRef = useRef<HTMLDivElement>(null);
  const exportSaveResolverRef = useRef<{
    reject: (reason?: unknown) => void;
    resolve: (value: boolean | PromiseLike<boolean>) => void;
  } | null>(null);
  const nativeSaveGuardRef = useRef(false);
  const reauthResolverRef = useRef<((allowed: boolean) => void) | null>(null);
  const allowNavigationRef = useRef(false);
  const navigationBlockerRef = useRef<{
    proceed: () => void;
    reset: () => void;
  } | null>(null);
  const navigationBlocker = useBlocker({
    disabled: !dirty,
    enableBeforeUnload: dirty,
    shouldBlockFn: () =>
      shouldBlockDirtyNavigation(dirty, allowNavigationRef.current),
    withResolver: true,
  });
  useEffect(() => {
    if (navigationBlocker.status === "blocked") {
      navigationBlockerRef.current = navigationBlocker;
      setExitIntent("navigation");
      setSaveBeforeExit(false);
      setOperationError(null);
      return;
    }
    navigationBlockerRef.current = null;
  }, [navigationBlocker]);

  useEffect(() => {
    if (authLoading) {
      return;
    }
    if (!user) {
      setLoading(false);
      return;
    }

    let cancelled = false;
    setLoading(true);
    setForm(null);
    setEditorConfigUrl(null);
    setNativeConfig(null);
    setNativeValues({});
    setNativePictureFiles({});
    setError(null);
    setHandoffError(false);
    const loadForm = async () => {
      try {
        const payload = await apiGet<{ form: PublicForm } | PublicForm>(
          `/api/forms/${publicId}`
        );
        if (!cancelled) {
          setForm("form" in payload ? payload.form : payload);
        }
      } catch (caughtError) {
        if (!cancelled) {
          setError(
            formRequestError(
              caughtError,
              "ไม่พบแบบฟอร์มนี้ หรือแบบฟอร์มยังไม่พร้อมใช้งาน"
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
  }, [authLoading, publicId, user]);

  useEffect(() => {
    const handleReauthentication = (event: Event) => {
      const { detail } = event as CustomEvent<ReauthenticationRequest>;
      if (!detail || !dirty || !activeResponseId) {
        return;
      }
      detail.handled = true;
      reauthResolverRef.current = detail.resolve;
      setExitIntent("reauth");
      setSaveBeforeExit(false);
      setError(null);
      setOperationError(null);
    };
    window.addEventListener("folio:before-reauth", handleReauthentication);
    return () =>
      window.removeEventListener("folio:before-reauth", handleReauthentication);
  }, [activeResponseId, dirty]);
  useEffect(() => {
    if (!exitIntent) {
      return;
    }
    document.querySelector<HTMLElement>("#unsaved-cancel")?.focus();
  }, [exitIntent]);
  useEffect(() => {
    if (error) {
      startErrorRef.current?.focus();
    }
  }, [error]);
  useEffect(() => {
    if (authLoading || !user || !form || editorConfigUrl) {
      return;
    }

    let cancelled = false;
    const startResponse = async () => {
      setLoading(true);
      setStartBusy(true);
      setError(null);
      setHandoffError(false);
      try {
        const result = await apiPost<{
          editorConfigUrl?: string;
          fillMethod: FillMethod;
          response?: { id: string };
          submissionId?: string;
        }>(`/api/forms/${publicId}/start`, {
          responseId: activeResponseId,
        });
        if (result.submissionId) {
          await navigate({
            params: { submissionId: result.submissionId },
            to: "/receipt/$submissionId",
          });
          return;
        }
        if (cancelled) {
          return;
        }
        if (result.fillMethod === "native" && !result.editorConfigUrl) {
          throw new Error("Native form configuration is unavailable");
        }
        const config =
          result.fillMethod === "native" && result.editorConfigUrl
            ? await apiGet<NativeEditorConfig | { fillMethod: "onlyoffice" }>(
                result.editorConfigUrl
              )
            : null;
        if (cancelled) {
          return;
        }
        const fillMethod = config?.fillMethod ?? result.fillMethod;
        setActiveResponseId(result.response?.id ?? activeResponseId);
        setForm((current) => (current ? { ...current, fillMethod } : current));
        if (config?.fillMethod === "native") {
          setNativeConfig(config);
          setNativeValues(nativeValuesFromConfig(config));
        } else {
          setNativeConfig(null);
          setNativeValues({});
          setNativePictureFiles({});
        }
        setEditorConfigUrl(result.editorConfigUrl ?? null);
      } catch (caughtError) {
        if (!cancelled) {
          setHandoffError(isHandoffError(caughtError));
          setError(
            formRequestError(caughtError, "ไม่สามารถเริ่มคำตอบนี้ได้ กรุณาลองใหม่")
          );
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
          setStartBusy(false);
        }
      }
    };

    void startResponse();
    return () => {
      cancelled = true;
    };
  }, [
    activeResponseId,
    authLoading,
    editorConfigUrl,
    form,
    publicId,
    startAttempt,
    user,
  ]);
  const resolveReauthentication = (allowed: boolean): void => {
    reauthResolverRef.current?.(allowed);
    reauthResolverRef.current = null;
  };
  // oxlint-disable-next-line complexity -- Applies bridge updates, operation state, and save/export transitions.
  const handleBridgeMessage = async (message: EditorBridgeMessage) => {
    if (message.type === "dirty-state") {
      setDirty(message.dirty);
      return;
    }
    if (message.type !== "operation") {
      return;
    }
    const { status } = message;
    setOperation(
      message.operationId
        ? {
            error: message.error,
            id: message.operationId,
            status,
          }
        : null
    );
    if (status === "failed") {
      if (message.action === "save-draft") {
        exportSaveResolverRef.current?.reject(new DraftSaveError());
        exportSaveResolverRef.current = null;
      }
      setSaveBeforeExit(false);
      setExportAfterSave(null);
      setOperationError(
        message.action === "save-draft"
          ? "บันทึกฉบับร่างไม่สำเร็จ กรุณาลองใหม่"
          : (message.error ?? "การดำเนินการกับเอกสารไม่สำเร็จ กรุณาลองใหม่")
      );
      setSuccess(null);
      if (message.action === "save-draft" && exitIntent === "reauth") {
        setExitIntent(null);
        resolveReauthentication(false);
      }
      return;
    }

    if (status !== "completed") {
      return;
    }

    setDirty(false);
    if (message.action === "save-draft") {
      setEditorRevision((value) => value + 1);
    }
    if (message.action === "submit") {
      const submissionId = message.operation?.result?.submissionId;
      if (submissionId) {
        allowNavigationRef.current = true;
        await navigate({
          params: { submissionId },
          to: "/receipt/$submissionId",
        });
        return;
      }
    }

    if (message.action === "save-draft" && exportSaveResolverRef.current) {
      const resolver = exportSaveResolverRef.current;
      exportSaveResolverRef.current = null;
      resolver.resolve(true);
      return;
    }
    setError(null);
    setOperationError(null);
    setSuccess(
      message.action === "save-draft"
        ? "บันทึกฉบับร่างคำตอบแล้ว"
        : "ดำเนินการกับเอกสารเรียบร้อยแล้ว"
    );
    if (message.action === "save-draft" && saveBeforeExit) {
      const intent = exitIntent;
      setSaveBeforeExit(false);
      if (intent === "reauth") {
        setExitIntent(null);
        resolveReauthentication(true);
      } else if (intent === "dashboard") {
        allowNavigationRef.current = true;
        setExitIntent(null);
        await navigate({ to: "/dashboard" });
      } else {
        allowNavigationRef.current = true;
        navigationBlockerRef.current?.proceed();
        setExitIntent(null);
      }
    }
  };
  const loadLatestNativeConfig = async (
    currentConfig: NativeEditorConfig,
    configUrl: string
  ): Promise<NativeEditorConfig> => {
    const latestConfig = await apiGet<
      NativeEditorConfig | { fillMethod: "onlyoffice" }
    >(configUrl);
    if (latestConfig.fillMethod !== "native") {
      setForm((current) =>
        current ? { ...current, fillMethod: "onlyoffice" } : current
      );
      setNativeConfig(null);
      setNativePictureFiles({});
      setNativeValues({});
      throw new Error("The form Fill Method changed");
    }
    if (latestConfig.documentKey !== currentConfig.documentKey) {
      setNativeConfig(latestConfig);
      if (!dirty) {
        setNativeValues(nativeValuesFromConfig(latestConfig));
      }
      throw new Error("The saved response changed in another session");
    }
    setNativeConfig(latestConfig);
    return latestConfig;
  };

  const submitNativeResponse = async (
    action: "save-draft" | "submit",
    config: NativeEditorConfig
  ): Promise<Operation> => {
    const payload = {
      data: Object.fromEntries(
        config.fields
          .filter((field) => field.type !== "picture")
          .map((field): [string, unknown] => {
            const value = nativeValues[field.tag];
            return [
              field.tag,
              field.type === "date" && value === "" ? null : value,
            ];
          })
      ),
      documentKey: config.documentKey,
      fillMethod: "native",
      responseId: config.responseId,
    };
    const formData = new FormData();
    formData.set("payload", JSON.stringify(payload));
    for (const [tag, file] of Object.entries(nativePictureFiles)) {
      formData.set(`picture:${tag}`, file);
    }
    const endpoint = action === "save-draft" ? "draft" : "submit";
    const { operationId } = await apiPostFormData<{ operationId: string }>(
      `/api/forms/${publicId}/${endpoint}`,
      formData,
      config.capabilities[action]
    );
    setOperation({ id: operationId, status: "pending" });
    const completed = await waitForOperation(operationId, setOperation);
    setOperation(completed);
    return completed;
  };

  const refreshNativeResponse = async (configUrl: string): Promise<void> => {
    const refreshedConfig = await apiGet<
      NativeEditorConfig | { fillMethod: "onlyoffice" }
    >(configUrl);
    setForm((current) =>
      current ? { ...current, fillMethod: refreshedConfig.fillMethod } : current
    );
    setNativePictureFiles({});
    setPictureInputKey((key) => key + 1);
    if (refreshedConfig.fillMethod === "native") {
      setNativeConfig(refreshedConfig);
      setNativeValues(nativeValuesFromConfig(refreshedConfig));
    } else {
      setNativeConfig(null);
      setNativeValues({});
    }
    setDirty(false);
    setOperationError(null);
  };

  const finishNativeSaveExit = async (): Promise<void> => {
    const intent = exitIntent;
    setSaveBeforeExit(false);
    setExitIntent(null);
    if (intent === "reauth") {
      resolveReauthentication(true);
      return;
    }
    allowNavigationRef.current = true;
    if (intent === "dashboard") {
      await navigate({ to: "/dashboard" });
      return;
    }
    navigationBlockerRef.current?.proceed();
  };
  const finishNativeResponse = async (
    action: "save-draft" | "submit",
    completed: Operation,
    options: { exitAfterSave?: boolean; exportFormat?: ExportFormat },
    configUrl: string,
    savedResponseId: string
  ): Promise<void> => {
    if (action === "submit") {
      const submissionId = completed.result?.submissionId;
      if (typeof submissionId !== "string") {
        throw new TypeError("The submitted receipt is unavailable");
      }
      allowNavigationRef.current = true;
      await navigate({
        params: { submissionId },
        to: "/receipt/$submissionId",
      });
      return;
    }
    await refreshNativeResponse(configUrl);
    setSuccess(draftSaveSuccessMessage(options.exportFormat));
    if (options.exportFormat) {
      await downloadDraftArtifact(savedResponseId, options.exportFormat);
      setExportAfterSave(null);
    }
    if (options.exitAfterSave) {
      await finishNativeSaveExit();
    }
  };

  const nativeResponseSaveBlocked = (): boolean =>
    nativeSaveGuardRef.current ||
    operation?.status === "pending" ||
    operation?.status === "processing" ||
    Boolean(exportAfterSave) ||
    saveBeforeExit ||
    discardBusy;

  const saveNativeResponse = async (
    action: "save-draft" | "submit",
    options: { exitAfterSave?: boolean; exportFormat?: ExportFormat } = {}
  ): Promise<void> => {
    if (!activeResponseId || !nativeConfig || !editorConfigUrl) {
      return;
    }
    if (nativeResponseSaveBlocked()) {
      return;
    }
    nativeSaveGuardRef.current = true;
    setNativeSaving(true);
    setSaveBeforeExit(options.exitAfterSave ?? false);
    setExportAfterSave(options.exportFormat ?? null);
    setOperationError(null);
    setSuccess(null);
    setError(null);
    let saved = false;
    try {
      const latestConfig = await loadLatestNativeConfig(
        nativeConfig,
        editorConfigUrl
      );
      const completed = await submitNativeResponse(action, latestConfig);
      saved = true;
      if (action === "submit") {
        setNativePictureFiles({});
        setPictureInputKey((key) => key + 1);
        setDirty(false);
      }
      await finishNativeResponse(
        action,
        completed,
        options,
        editorConfigUrl,
        activeResponseId
      );
    } catch (caughtError) {
      setSaveBeforeExit(false);
      setExportAfterSave(null);
      setOperationError(
        formRequestError(
          caughtError,
          nativeSaveErrorMessage(action, saved, options.exportFormat)
        )
      );
      setSuccess(null);
      if (options.exitAfterSave && exitIntent === "reauth") {
        setExitIntent(null);
        resolveReauthentication(false);
      }
    } finally {
      nativeSaveGuardRef.current = false;
      setNativeSaving(false);
    }
  };

  const fillPath = safeReturnPath(
    `/forms/${publicId}/fill${
      responseId ? `?responseId=${encodeURIComponent(responseId)}` : ""
    }`
  );

  if (authLoading || loading) {
    return (
      <div className="grid min-h-screen place-items-center">
        <Spinner />
      </div>
    );
  }

  if (!user) {
    return (
      <Navigate
        to="/login"
        search={{ returnTo: fillPath ?? undefined }}
        replace
      />
    );
  }

  if (error || !form) {
    return (
      <div className="mx-auto max-w-xl px-5 py-16">
        <div ref={startErrorRef} tabIndex={-1}>
          <Notice tone="danger">
            {error ?? "ไม่พบแบบฟอร์มนี้ หรือแบบฟอร์มยังไม่พร้อมใช้งาน"}
          </Notice>
        </div>
        {handoffError ? (
          <div className="mt-5 flex flex-wrap gap-2">
            <Button
              type="button"
              variant="secondary"
              onClick={() => setStartAttempt((attempt) => attempt + 1)}
              disabled={startBusy}
            >
              {startBusy ? <Spinner /> : null}
              {startBusy ? "กำลังลองใหม่…" : "ลองใหม่"}
            </Button>
            <Button type="button" variant="ghost" onClick={returnToSource}>
              <ArrowLeft size={16} />
              กลับไปยังระบบต้นทาง
            </Button>
          </div>
        ) : null}
      </div>
    );
  }

  const operationBusy =
    operation?.status === "pending" || operation?.status === "processing";
  const saveFlowBusy = isSaveFlowBusy(
    operationBusy || nativeSaving,
    exportAfterSave,
    saveBeforeExit
  );
  const handleExit = async () => {
    if (saveFlowBusy || discardBusy) {
      return;
    }
    if (!dirty) {
      allowNavigationRef.current = true;
      await navigate({ to: "/dashboard" });
      return;
    }
    setExitIntent("dashboard");
    setSaveBeforeExit(false);
    setOperationError(null);
  };
  const saveAndExit = () => {
    if (!activeResponseId || saveFlowBusy || discardBusy) {
      return;
    }
    if (nativeConfig) {
      void saveNativeResponse("save-draft", { exitAfterSave: true });
      return;
    }
    setSaveBeforeExit(true);
    setOperationError(null);
    setSuccess(null);
    setSaveRequest((value) => value + 1);
  };
  const saveAndExport = (format: ExportFormat) => {
    if (
      !activeResponseId ||
      saveFlowBusy ||
      discardBusy ||
      exportSaveResolverRef.current
    ) {
      return;
    }
    if (nativeConfig) {
      void saveNativeResponse("save-draft", { exportFormat: format });
      return;
    }
    const saveCompletion = createDeferred<boolean>();
    exportSaveResolverRef.current = {
      reject: saveCompletion.reject,
      resolve: saveCompletion.resolve,
    };
    setExportAfterSave(format);
    setOperationError(null);
    setSuccess(null);
    const completeExport = async () => {
      try {
        await saveThenDownload(
          () => saveCompletion.promise,
          () => downloadDraftArtifact(activeResponseId, format)
        );
        setExportAfterSave(null);
        setError(null);
        setOperationError(null);
        setSuccess(draftSaveSuccessMessage(format));
      } catch (caughtError: unknown) {
        setExportAfterSave(null);
        setOperationError(
          formRequestError(
            caughtError,
            caughtError instanceof DraftSaveError
              ? "บันทึกฉบับร่างไม่สำเร็จ กรุณาลองใหม่"
              : "ดาวน์โหลดไฟล์ไม่สำเร็จ กรุณาลองใหม่"
          )
        );
      }
    };
    completeExport();
    setSaveRequest((value) => value + 1);
  };
  const cancelExit = () => {
    navigationBlockerRef.current?.reset();
    navigationBlockerRef.current = null;
    resolveReauthentication(false);
    setSaveBeforeExit(false);
    setExitIntent(null);
  };
  const discardDraft = async () => {
    if (!activeResponseId || discardBusy || saveFlowBusy) {
      return;
    }
    setDiscardBusy(true);
    setOperationError(null);
    try {
      await apiDelete(`/api/responses/${activeResponseId}`);
      setDirty(false);
      setClearDirtyRequest((value) => value + 1);
      const intent = exitIntent;
      if (intent === "reauth") {
        resolveReauthentication(true);
      } else if (intent === "dashboard") {
        allowNavigationRef.current = true;
        await navigate({ to: "/dashboard" });
      } else {
        allowNavigationRef.current = true;
        navigationBlockerRef.current?.proceed();
      }
      setExitIntent(null);
    } catch (caughtError) {
      setOperationError(
        formRequestError(caughtError, "ลบฉบับร่างไม่สำเร็จ กรุณาลองใหม่")
      );
    } finally {
      setDiscardBusy(false);
    }
  };

  const exitPrompt = exitIntent ? (
    <div
      className="fixed inset-0 z-50 grid place-items-center bg-black/35 px-5"
      role="presentation"
    >
      <div
        className="w-full max-w-lg rounded-[var(--radius)] border border-[var(--line)] bg-[var(--paper)] p-6 shadow-xl"
        role="dialog"
        aria-modal="true"
        aria-labelledby="unsaved-title"
      >
        <h2 id="unsaved-title" className="text-lg font-semibold">
          มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก
        </h2>
        <p className="mt-2 text-sm text-[var(--ink-soft)]">
          {exitIntent === "reauth"
            ? "บันทึกฉบับร่างก่อนเข้าสู่ระบบเดิม หรืออยู่ต่อเพื่อยกเลิก"
            : "บันทึกฉบับร่างก่อนออกจากแบบฟอร์ม หรือทิ้งฉบับร่างนี้อย่างถาวร"}
        </p>
        {operationError ? (
          <div className="mt-4">
            <Notice tone="danger">{operationError}</Notice>
          </div>
        ) : null}
        <div className="mt-6 flex flex-wrap justify-end gap-2">
          {exitIntent === "reauth" ? null : (
            <Button
              variant="secondary"
              type="button"
              onClick={discardDraft}
              disabled={discardBusy || saveFlowBusy}
            >
              {discardBusy ? <Spinner /> : null}
              {discardBusy ? "กำลังลบ…" : "ทิ้งฉบับร่าง"}
            </Button>
          )}
          <Button
            type="button"
            onClick={saveAndExit}
            disabled={discardBusy || saveFlowBusy}
          >
            {saveBeforeExit ? <Spinner /> : null}
            {saveBeforeExit ? "กำลังบันทึก…" : "บันทึกแล้วออก"}
          </Button>
          <Button
            id="unsaved-cancel"
            variant="ghost"
            type="button"
            onClick={cancelExit}
            disabled={discardBusy || saveFlowBusy}
          >
            อยู่ต่อ
          </Button>
        </div>
      </div>
    </div>
  ) : null;

  return (
    <div className="min-h-screen bg-[var(--canvas)]">
      {exitPrompt}

      <header className="border-b border-[var(--line)] bg-[var(--paper)]">
        <div className="mx-auto flex max-w-[1240px] items-center justify-between px-5 py-4 lg:px-8">
          <Button
            variant="ghost"
            size="sm"
            onClick={handleExit}
            disabled={discardBusy || saveFlowBusy}
          >
            <ArrowLeft />
            ออกจากแบบฟอร์ม
          </Button>
          <div className="flex items-center gap-2 text-sm font-semibold">
            <span className="grid size-7 place-items-center rounded-lg bg-[var(--ink)] text-[var(--accent)]">
              F
            </span>
            Folio Forms
          </div>
          <div className="flex flex-wrap items-center justify-end gap-2">
            <span className="max-w-40 truncate text-xs text-[var(--ink-soft)]">
              {user.email}
            </span>
            {roleFor(user) === "user" ? (
              <LegacySsoSwitchButton compact />
            ) : null}
          </div>
        </div>
      </header>
      <main className="mx-auto max-w-[1240px] px-5 py-8 lg:px-8 lg:py-10">
        <div className="mb-7 flex flex-col justify-between gap-4 sm:flex-row sm:items-end">
          <div>
            <h1 className="text-3xl font-bold tracking-[-0.04em]">
              {form.title}
            </h1>
            <p className="mt-2 max-w-2xl text-[var(--ink-soft)]">
              {form.description || "กรอกข้อมูลด้านล่าง แล้วบันทึกฉบับร่างหรือส่งแบบฟอร์ม"}
            </p>
          </div>
          {nativeConfig ? (
            <span className="text-sm text-[var(--ink-soft)]">
              แบบฟอร์ม Native
            </span>
          ) : (
            <div className="flex items-center gap-2 text-sm text-[var(--ink-soft)]">
              <Monitor />
              แนะนำให้ใช้ตัวแก้ไขบนคอมพิวเตอร์
            </div>
          )}
        </div>
        {error ? (
          <div className="mb-4">
            <Notice tone="danger">{error}</Notice>
          </div>
        ) : null}
        {operationError ? (
          <div className="mb-4">
            <Notice tone="danger">
              {operationError} คุณสามารถแก้ไขข้อมูลแล้วลองใหม่ได้
            </Notice>
          </div>
        ) : null}
        {success ? (
          <div className="mb-4">
            <Notice tone="success">
              <span className="inline-flex items-center gap-2">
                <CheckCircle2 />
                {success}
              </span>
            </Notice>
          </div>
        ) : null}
        {operationBusy || nativeSaving ? (
          <div className="mb-4">
            <Notice>
              <span className="inline-flex items-center gap-2">
                <Spinner />
                {operation?.status === "processing"
                  ? "กำลังเตรียมไฟล์…"
                  : "กำลังบันทึกคำตอบ…"}
              </span>
            </Notice>
          </div>
        ) : null}
        <div className="mb-4 flex items-center justify-between gap-3 rounded-[10px] border border-[var(--line)] bg-[var(--paper)] px-4 py-3">
          <div className="text-sm text-[var(--ink-soft)]">
            สถานะ: {dirty ? "มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก" : "บันทึกแล้ว"}
          </div>
        </div>
        {nativeConfig ? (
          <NativeForm
            fields={nativeConfig.fields}
            lockedFields={nativeConfig.lockedFields}
            operationBusy={
              operationBusy ||
              nativeSaving ||
              Boolean(exportAfterSave) ||
              saveBeforeExit
            }
            values={nativeValues}
            pictures={nativeConfig.pictures}
            pictureFiles={nativePictureFiles}
            pictureInputKey={pictureInputKey}
            onChange={(tag, value) => {
              setNativeValues((current) => ({ ...current, [tag]: value }));
              setDirty(true);
              setOperationError(null);
            }}
            onPictureChange={(tag, file) => {
              setNativePictureFiles((current) => {
                if (file) {
                  return { ...current, [tag]: file };
                }
                return Object.fromEntries(
                  Object.entries(current).filter(
                    ([currentTag]) => currentTag !== tag
                  )
                );
              });
              setDirty(true);
              setOperationError(null);
            }}
            onExportDocx={() => saveAndExport("docx")}
            onExportPdf={() => saveAndExport("pdf")}
            onSave={() => saveNativeResponse("save-draft")}
            onSubmit={() => saveNativeResponse("submit")}
          />
        ) : (
          <>
            <div className="mb-4 rounded-[10px] border border-[var(--accent)]/35 bg-[var(--accent-soft)] px-4 py-3 text-sm text-[var(--ink)]">
              <strong>ใช้แท็บ Form ในตัวแก้ไขเอกสาร</strong>{" "}
              เพื่อเลือกบันทึกฉบับร่างหรือส่งคำตอบ ระบบจะแสดงความคืบหน้าที่นี่
            </div>
            <div className="mb-4 flex flex-wrap items-center justify-end gap-3 rounded-[10px] border border-[var(--line)] bg-[var(--paper)] px-4 py-3">
              <Button
                type="button"
                variant="secondary"
                size="sm"
                onClick={() => saveAndExport("docx")}
                disabled={saveFlowBusy || discardBusy}
              >
                {exportAfterSave === "docx" ? <Spinner /> : null}
                บันทึกและดาวน์โหลด DOCX
              </Button>
              <Button
                type="button"
                variant="secondary"
                size="sm"
                onClick={() => saveAndExport("pdf")}
                disabled={saveFlowBusy || discardBusy}
              >
                {exportAfterSave === "pdf" ? <Spinner /> : null}
                บันทึกและดาวน์โหลด PDF
              </Button>
            </div>
            <div className="overflow-hidden rounded-[var(--radius)] border border-[var(--line-strong)] bg-[var(--muted)] shadow-inner">
              <OnlyOfficeEditor
                clearDirtyRequest={clearDirtyRequest}
                configUrl={editorConfigUrl ?? undefined}
                onBridgeMessage={handleBridgeMessage}
                onDirtyChange={setDirty}
                revision={editorRevision}
                saveRequest={saveRequest}
                title={`กรอกแบบฟอร์ม ${form.title}`}
              />
            </div>
          </>
        )}
      </main>
    </div>
  );
};
export const Route = createFileRoute("/forms/$publicId/fill")({
  component: FillRoute,
  validateSearch: (search) => ({
    responseId:
      typeof search.responseId === "string" ? search.responseId : undefined,
  }),
});
