import { useCallback, useEffect, useId, useRef, useState } from "react";

import { Button, Notice, Spinner } from "@/components/ui";
import { ApiError, API_ORIGIN, apiDelete, apiGet, apiPost } from "@/lib/api";

type EditorAction =
  | "save-template"
  | "publish"
  | "save-draft"
  | "save-correction"
  | "submit"
  | "configure-fields";
type EditorOperationAction = Exclude<EditorAction, "configure-fields">;
type FieldControlType =
  | "text"
  | "checkbox"
  | "date"
  | "dropdown"
  | "combo"
  | "picture"
  | "unsupported";
type EditorOperationStatus = "pending" | "completed" | "failed";
export type OnlyOfficeEditorState = "loading" | "ready" | "blocked" | "error";

interface EditorLease {
  id: string;
  expiresAt: string;
  releaseUrl: string;
  renewUrl: string;
}

const leaseRenewalIntervalMs = 30_000;

interface EditorConfig {
  apiScriptUrl?: string;
  apiUrl?: string;
  bridge?: {
    capabilities?: Partial<Record<EditorAction, string>>;
    id?: string;
    lease?: EditorLease;
    pluginOrigin?: string;
  };
  config?: Record<string, unknown>;
  editorUrl?: string;
  [key: string]: unknown;
}

interface DirtyStateBridgeMessage {
  bridgeId: string;
  dirty: boolean;
  source: "form-bridge";
  type: "dirty-state";
}
interface BridgeReadyMessage {
  bridgeId: string;
  source: "form-bridge";
  type: "bridge-ready";
}

interface CapabilityRequestMessage {
  action: EditorAction;
  bridgeId: string;
  requestId: string;
  source: "form-bridge";
  type: "capability-request";
}
interface FieldSelectionBridgeMessage {
  bridgeId: string;
  controlType: FieldControlType;
  selectionId: string;
  selected: boolean;
  source: "form-bridge";
  tag: string | null;
  type: "field-selection";
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);

const isNonEmptyString = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0;
const isEditorOperationAction = (
  value: unknown
): value is EditorOperationAction =>
  value === "save-template" ||
  value === "publish" ||
  value === "save-draft" ||
  value === "save-correction" ||
  value === "submit";
const isEditorAction = (value: unknown): value is EditorAction =>
  value === "configure-fields" || isEditorOperationAction(value);
const isFieldControlType = (value: unknown): value is FieldControlType =>
  value === "text" ||
  value === "checkbox" ||
  value === "date" ||
  value === "dropdown" ||
  value === "combo" ||
  value === "picture" ||
  value === "unsupported";

const isOperationStatus = (value: unknown): value is EditorOperationStatus =>
  value === "pending" || value === "completed" || value === "failed";

const isBridgeReadyMessage = (
  value: unknown,
  bridgeId: string
): value is BridgeReadyMessage =>
  isRecord(value) &&
  value.bridgeId === bridgeId &&
  value.source === "form-bridge" &&
  value.type === "bridge-ready";

const parseDirtyStateMessage = (
  value: unknown,
  bridgeId: string
): DirtyStateBridgeMessage | null =>
  isRecord(value) &&
  value.bridgeId === bridgeId &&
  typeof value.dirty === "boolean" &&
  value.source === "form-bridge" &&
  value.type === "dirty-state"
    ? (value as unknown as DirtyStateBridgeMessage)
    : null;

const parseCapabilityRequest = (
  value: unknown,
  bridgeId: string
): CapabilityRequestMessage | null => {
  if (
    !isRecord(value) ||
    value.bridgeId !== bridgeId ||
    value.source !== "form-bridge" ||
    value.type !== "capability-request" ||
    typeof value.action !== "string" ||
    !isEditorAction(value.action) ||
    typeof value.requestId !== "string" ||
    !value.requestId
  ) {
    return null;
  }

  return value as unknown as CapabilityRequestMessage;
};

const parseFieldSelectionMessage = (
  value: unknown,
  bridgeId: string
): FieldSelectionBridgeMessage | null => {
  if (
    !isRecord(value) ||
    value.bridgeId !== bridgeId ||
    value.source !== "form-bridge" ||
    value.type !== "field-selection" ||
    !isNonEmptyString(value.selectionId) ||
    typeof value.selected !== "boolean" ||
    (value.tag !== null && typeof value.tag !== "string") ||
    !isFieldControlType(value.controlType)
  ) {
    return null;
  }

  return value as unknown as FieldSelectionBridgeMessage;
};

const parseOperationMessage = (
  value: unknown,
  bridgeId: string
): EditorOperationBridgeMessage | null => {
  if (
    !isRecord(value) ||
    value.bridgeId !== bridgeId ||
    value.source !== "form-bridge" ||
    value.type !== "operation" ||
    typeof value.action !== "string" ||
    !isEditorOperationAction(value.action) ||
    !isOperationStatus(value.status) ||
    (value.operationId !== undefined && !isNonEmptyString(value.operationId)) ||
    (value.status !== "failed" && !isNonEmptyString(value.operationId)) ||
    (value.error !== undefined && typeof value.error !== "string")
  ) {
    return null;
  }

  if (value.operation !== undefined) {
    if (!isRecord(value.operation)) {
      return null;
    }
    if (value.operation.result !== undefined) {
      if (!isRecord(value.operation.result)) {
        return null;
      }
      if (
        value.operation.result.submissionId !== undefined &&
        typeof value.operation.result.submissionId !== "string"
      ) {
        return null;
      }
    }
  }

  return value as unknown as EditorOperationBridgeMessage;
};

interface EditorOperationBridgeMessage {
  action: EditorOperationAction;
  bridgeId: string;
  error?: string;
  operation?: {
    result?: {
      submissionId?: string;
    };
  };
  operationId?: string;
  source: "form-bridge";
  status: EditorOperationStatus;
  type: "operation";
}

export type EditorBridgeMessage =
  | DirtyStateBridgeMessage
  | EditorOperationBridgeMessage
  | FieldSelectionBridgeMessage;

const acknowledgeBridge = (
  source: MessageEventSource,
  pluginOrigin: string,
  bridgeId: string
) => {
  try {
    (source as Window).postMessage(
      {
        bridgeId,
        source: "folio-parent",
        type: "bridge-ack",
      },
      pluginOrigin
    );
  } catch {
    // The plugin may close its frame while the handshake is in flight.
  }
};

interface DocsApi {
  DocEditor: new (
    elementId: string,
    config: Record<string, unknown>
  ) => { destroyEditor?: () => void };
}

declare global {
  interface Window {
    DocsAPI?: DocsApi;
  }
}

export const OnlyOfficeEditor = ({
  clearDirtyRequest = 0,
  configUrl,
  onBridgeMessage,
  onDirtyChange,
  onStateChange,
  readOnly = false,
  revision = 0,
  saveAction = "save-draft",
  saveReason = "",
  saveRequest = 0,
  title,
}: {
  clearDirtyRequest?: number;
  configUrl?: string;
  onBridgeMessage?: (message: EditorBridgeMessage) => void | Promise<void>;
  onDirtyChange?: (dirty: boolean) => void;
  onStateChange?: (state: OnlyOfficeEditorState) => void;
  readOnly?: boolean;
  revision?: number;
  saveAction?: "save-draft" | "save-correction";
  saveReason?: string;
  saveRequest?: number;
  title: string;
}) => {
  const hostRef = useRef<HTMLDivElement>(null);
  const surfaceRef = useRef<HTMLDivElement>(null);
  const feedbackRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<{ destroyEditor?: () => void } | null>(null);
  const onBridgeMessageRef = useRef(onBridgeMessage);
  const onDirtyChangeRef = useRef(onDirtyChange);
  const onStateChangeRef = useRef(onStateChange);
  const pinnedSourceRef = useRef<MessageEventSource | null>(null);
  const lastClearDirtyRequestRef = useRef(0);
  const lastSaveRequestRef = useRef(0);
  const terminalOperationIdsRef = useRef(new Set<string>());
  const editorId = useId().replaceAll(":", "");
  const surfaceId = `${editorId}-surface`;
  const leaseRef = useRef<EditorLease | null>(null);
  const loadedConfigUrlRef = useRef<string | null>(null);
  const [config, setConfig] = useState<EditorConfig | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [editorState, setEditorState] =
    useState<OnlyOfficeEditorState>("loading");
  const [expanded, setExpanded] = useState(false);
  const [retryToken, setRetryToken] = useState(0);
  const [bridgeReadyVersion, setBridgeReadyVersion] = useState(0);

  const reportState = useCallback((nextState: OnlyOfficeEditorState) => {
    setEditorState(nextState);
    onStateChangeRef.current?.(nextState);
  }, []);

  const releaseCurrentLease = useCallback(async (): Promise<void> => {
    const lease = leaseRef.current;
    leaseRef.current = null;
    if (!lease) {
      return;
    }
    try {
      await apiDelete(lease.releaseUrl, undefined, { keepalive: true });
    } catch {
      // Lease expiry remains the fallback if best-effort release is unavailable.
    }
  }, []);

  useEffect(() => {
    if (
      leaseRef.current &&
      loadedConfigUrlRef.current !== (configUrl ?? null)
    ) {
      void releaseCurrentLease();
    }
  }, [configUrl, releaseCurrentLease]);

  useEffect(
    () => () => {
      void releaseCurrentLease();
    },
    [releaseCurrentLease]
  );
  useEffect(() => {
    if (editorState === "blocked" || editorState === "error") {
      feedbackRef.current?.focus();
    }
  }, [editorState]);
  useEffect(() => {
    const surface = surfaceRef.current;
    if (!expanded || !surface) {
      return;
    }

    const obscured: HTMLElement[] = [];
    const focusableAncestors: [HTMLElement, string | null][] = [];
    let branch: HTMLElement = surface;
    while (branch.parentElement) {
      const parent = branch.parentElement;
      for (const sibling of parent.children) {
        if (
          sibling instanceof HTMLElement &&
          sibling !== branch &&
          !sibling.inert
        ) {
          sibling.inert = true;
          obscured.push(sibling);
        }
      }
      branch = parent;
      if (parent.tabIndex >= 0) {
        focusableAncestors.push([parent, parent.getAttribute("tabindex")]);
        parent.tabIndex = -1;
      }
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        setExpanded(false);
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      for (const sibling of obscured) {
        sibling.inert = false;
      }
      for (const [ancestor, tabIndex] of focusableAncestors) {
        if (tabIndex === null) {
          ancestor.removeAttribute("tabindex");
        } else {
          ancestor.setAttribute("tabindex", tabIndex);
        }
      }
    };
  }, [expanded]);

  useEffect(() => {
    const lease =
      loadedConfigUrlRef.current === (configUrl ?? null)
        ? config?.bridge?.lease
        : undefined;
    if (!lease) {
      return;
    }

    let cancelled = false;
    leaseRef.current = lease;

    const renewLease = async () => {
      const currentLease = leaseRef.current;
      if (cancelled || !currentLease) {
        return;
      }

      try {
        const response = await apiPost<{
          lease: Pick<EditorLease, "expiresAt" | "id">;
        }>(currentLease.renewUrl);
        const activeLease = leaseRef.current;
        if (
          !cancelled &&
          activeLease?.id === currentLease.id &&
          response.lease.id === currentLease.id
        ) {
          leaseRef.current = {
            ...activeLease,
            expiresAt: response.lease.expiresAt,
          };
        }
      } catch {
        // A transient heartbeat failure must not tear down unsaved editor state.
      }
    };

    const renewalTimer = window.setInterval(() => {
      void renewLease();
    }, leaseRenewalIntervalMs);

    return () => {
      cancelled = true;
      window.clearInterval(renewalTimer);
      // Preserve the lease reference across a same-document config refresh.
    };
  }, [config, configUrl]);

  useEffect(() => {
    onBridgeMessageRef.current = onBridgeMessage;
    onDirtyChangeRef.current = onDirtyChange;
    onStateChangeRef.current = onStateChange;
  }, [onBridgeMessage, onDirtyChange, onStateChange]);
  useEffect(() => {
    const bridgeId = config?.bridge?.id;
    const pluginOrigin = config?.bridge?.pluginOrigin;
    const source = pinnedSourceRef.current;
    if (
      !source ||
      typeof bridgeId !== "string" ||
      !bridgeId ||
      typeof pluginOrigin !== "string" ||
      !pluginOrigin
    ) {
      return;
    }
    const postCommand = (message: Record<string, unknown>) => {
      try {
        (source as Window).postMessage(
          { ...message, bridgeId, source: "folio-parent" },
          pluginOrigin
        );
      } catch {
        // The editor may close while a command is in flight.
      }
    };
    if (saveRequest > lastSaveRequestRef.current) {
      lastSaveRequestRef.current = saveRequest;
      postCommand({
        action: saveAction,
        ...(saveAction === "save-correction" ? { reason: saveReason } : {}),
        type: "run-action",
      });
    }
    if (clearDirtyRequest > lastClearDirtyRequestRef.current) {
      lastClearDirtyRequestRef.current = clearDirtyRequest;
      postCommand({ type: "clear-dirty" });
    }
  }, [
    bridgeReadyVersion,
    clearDirtyRequest,
    config,
    saveAction,
    saveReason,
    saveRequest,
  ]);

  useEffect(() => {
    let cancelled = false;
    if (!configUrl) {
      setConfig(null);
      setError("ไม่มีการตั้งค่าตัวแก้ไขเอกสาร");
      reportState("error");
      return () => {
        cancelled = true;
      };
    }

    const path = configUrl.startsWith("http")
      ? configUrl.replace(API_ORIGIN, "")
      : configUrl;
    loadedConfigUrlRef.current = null;
    setConfig(null);
    setError(null);
    reportState("loading");

    const loadConfig = async () => {
      try {
        const nextConfig = await apiGet<EditorConfig>(path);
        if (cancelled) {
          return;
        }
        loadedConfigUrlRef.current = configUrl;
        setConfig(nextConfig);
        if (nextConfig.editorUrl) {
          reportState("ready");
        }
      } catch (caughtError) {
        if (cancelled) {
          return;
        }
        if (
          caughtError instanceof ApiError &&
          caughtError.code === "editor_in_use"
        ) {
          setError(null);
          reportState("blocked");
          return;
        }
        setError("ไม่สามารถโหลดตัวแก้ไขเอกสารได้ กรุณาลองใหม่อีกครั้ง");
        reportState("error");
      }
    };

    void loadConfig();
    return () => {
      cancelled = true;
    };
  }, [configUrl, reportState, retryToken, revision]);

  useEffect(() => {
    if (editorState === "blocked" || editorState === "error") {
      feedbackRef.current?.focus();
    }
  }, [editorState]);

  useEffect(() => {
    if (!config) {
      return;
    }

    const bridgeId = isNonEmptyString(config.bridge?.id)
      ? config.bridge.id
      : "";
    const pluginOrigin = isNonEmptyString(config.bridge?.pluginOrigin)
      ? config.bridge.pluginOrigin
      : "";
    const hasBridge = bridgeId.length > 0 && pluginOrigin.length > 0;
    if (
      (readOnly && config.bridge !== undefined) ||
      (!readOnly && !hasBridge && !config.editorUrl)
    ) {
      setError("การตั้งค่าตัวแก้ไขเอกสารไม่ถูกต้อง");
      reportState("error");
      return;
    }
    let cancelled = false;
    let mountedHost: HTMLDivElement | null = null;

    const scriptUrl =
      config.apiScriptUrl ??
      `${config.apiUrl ?? "http://localhost:8080"}/web-apps/apps/api/documents/api.js`;
    const existing = document.querySelector<HTMLScriptElement>(
      `script[src="${scriptUrl}"]`
    );
    pinnedSourceRef.current = null;
    terminalOperationIdsRef.current.clear();
    const respondToCapabilityRequest = (
      request: CapabilityRequestMessage,
      response: { capability: string } | { error: string }
    ) => {
      if (cancelled) {
        return;
      }
      const pinnedSource = pinnedSourceRef.current;
      if (!pinnedSource) {
        return;
      }
      try {
        (pinnedSource as Window).postMessage(
          {
            action: request.action,
            bridgeId,
            requestId: request.requestId,
            source: "folio-parent",
            type: "capability-response",
            ...response,
          },
          pluginOrigin
        );
      } catch {
        // The plugin may close its frame while renewal is in flight.
      }
    };

    const renewCapability = async (request: CapabilityRequestMessage) => {
      if (!configUrl) {
        respondToCapabilityRequest(request, {
          error: "ไม่พบสิทธิ์สำหรับตัวแก้ไขเอกสาร",
        });
        return;
      }

      const requestedLeaseId = leaseRef.current?.id;
      try {
        const path = configUrl.startsWith("http")
          ? configUrl.replace(API_ORIGIN, "")
          : configUrl;
        const fresh = await apiGet<EditorConfig>(path);
        if (cancelled) {
          return;
        }
        const freshLease = fresh.bridge?.lease;
        if (
          freshLease &&
          requestedLeaseId &&
          leaseRef.current?.id === requestedLeaseId
        ) {
          leaseRef.current = freshLease;
        }

        const capability = fresh.bridge?.capabilities?.[request.action];
        if (typeof capability !== "string" || !capability) {
          respondToCapabilityRequest(request, {
            error: "ไม่พบสิทธิ์สำหรับตัวแก้ไขเอกสาร",
          });
          return;
        }
        respondToCapabilityRequest(request, { capability });
      } catch {
        if (!cancelled) {
          respondToCapabilityRequest(request, {
            error: "ไม่สามารถยืนยันสิทธิ์ตัวแก้ไขเอกสารได้",
          });
        }
      }
    };

    // oxlint-disable-next-line complexity -- Dispatches the trusted bridge protocol and capability renewal in one handler.
    const handleBridgeMessage = (event: MessageEvent<unknown>) => {
      if (event.origin !== pluginOrigin) {
        return;
      }

      const { data } = event;
      if (isBridgeReadyMessage(data, bridgeId)) {
        if (!event.source) {
          return;
        }
        if (
          pinnedSourceRef.current &&
          pinnedSourceRef.current !== event.source
        ) {
          return;
        }
        pinnedSourceRef.current = event.source;
        acknowledgeBridge(event.source, pluginOrigin, bridgeId);
        setBridgeReadyVersion((value) => value + 1);
        return;
      }

      const pinnedSource = pinnedSourceRef.current;
      if (!pinnedSource || event.source !== pinnedSource) {
        return;
      }

      const capabilityRequest = parseCapabilityRequest(data, bridgeId);
      if (capabilityRequest) {
        void renewCapability(capabilityRequest);
        return;
      }
      const fieldSelection = parseFieldSelectionMessage(data, bridgeId);
      if (fieldSelection) {
        onBridgeMessageRef.current?.(fieldSelection);
        return;
      }
      const dirtyMessage = parseDirtyStateMessage(data, bridgeId);
      if (dirtyMessage) {
        onDirtyChangeRef.current?.(dirtyMessage.dirty);
        return;
      }
      const message = parseOperationMessage(data, bridgeId);
      if (!message) {
        return;
      }
      const isTerminal =
        message.status === "completed" || message.status === "failed";
      if (
        isTerminal &&
        message.operationId &&
        terminalOperationIdsRef.current.has(message.operationId)
      ) {
        return;
      }
      if (isTerminal && message.operationId) {
        terminalOperationIdsRef.current.add(message.operationId);
      }
      onBridgeMessageRef.current?.(message);
    };

    if (hasBridge) {
      window.addEventListener("message", handleBridgeMessage);
    }

    if (config.editorUrl || !hostRef.current) {
      return () => {
        cancelled = true;
        window.removeEventListener("message", handleBridgeMessage);
        pinnedSourceRef.current = null;
      };
    }

    if (!isRecord(config.config)) {
      setError("การตั้งค่าตัวแก้ไขเอกสารไม่ถูกต้อง");
      reportState("error");
      return () => {
        cancelled = true;
        window.removeEventListener("message", handleBridgeMessage);
        pinnedSourceRef.current = null;
      };
    }

    const editorConfig = config.config;
    const mount = () => {
      if (cancelled) {
        return;
      }
      if (!window.DocsAPI || !hostRef.current) {
        setError("ไม่สามารถเปิดตัวแก้ไขเอกสารได้ กรุณาตรวจสอบบริการ ONLYOFFICE");
        reportState("error");
        return;
      }

      const host = hostRef.current;
      const placeholder = document.createElement("div");
      placeholder.id = editorId;
      placeholder.className = "h-full w-full";
      host.replaceChildren(placeholder);
      mountedHost = host;
      editorRef.current = new window.DocsAPI.DocEditor(editorId, editorConfig);
      reportState("ready");
    };
    const handleScriptError = () => {
      if (!cancelled) {
        setError("ไม่สามารถโหลดตัวแก้ไขเอกสารได้ กรุณาลองใหม่อีกครั้ง");
        reportState("error");
      }
    };

    if (window.DocsAPI) {
      mount();
    } else if (existing) {
      existing.addEventListener("load", mount, { once: true });
      existing.addEventListener("error", handleScriptError, { once: true });
    } else {
      const script = document.createElement("script");
      script.src = scriptUrl;
      script.addEventListener("load", mount, { once: true });
      script.addEventListener("error", handleScriptError, { once: true });
      document.head.append(script);
    }

    return () => {
      cancelled = true;
      window.removeEventListener("message", handleBridgeMessage);
      pinnedSourceRef.current = null;
      const editor = editorRef.current;
      editorRef.current = null;
      try {
        editor?.destroyEditor?.();
      } finally {
        mountedHost?.replaceChildren();
      }
    };
  }, [config, configUrl, editorId, readOnly, reportState]);

  const retry = () => {
    setError(null);
    setConfig(null);
    reportState("loading");
    setRetryToken((value) => value + 1);
  };

  const renderEditorContent = () => {
    if (editorState === "blocked") {
      return (
        <div
          ref={feedbackRef}
          className={`grid ${
            expanded ? "h-full min-h-0" : "min-h-[520px]"
          } place-items-center p-8`}
          tabIndex={-1}
        >
          <Notice tone="danger">
            <div className="space-y-3">
              <p className="font-semibold">เอกสารนี้กำลังถูกแก้ไขโดยผู้ใช้รายอื่น</p>
              <p>ยังไม่เปิดตัวแก้ไขจนกว่าจะเชื่อมต่อใหม่ได้</p>
              <Button type="button" variant="secondary" onClick={retry}>
                ลองเชื่อมต่อใหม่
              </Button>
            </div>
          </Notice>
        </div>
      );
    }

    if (editorState === "error" || error) {
      return (
        <div
          ref={feedbackRef}
          className={`grid ${
            expanded ? "h-full min-h-0" : "min-h-[520px]"
          } place-items-center p-8`}
          tabIndex={-1}
        >
          <Notice tone="danger">
            <div className="space-y-3">
              <p>{error ?? "ไม่สามารถเปิดตัวแก้ไขเอกสารได้"}</p>
              <Button type="button" variant="secondary" onClick={retry}>
                ลองใหม่
              </Button>
            </div>
          </Notice>
        </div>
      );
    }

    if (!config) {
      return (
        <div
          className={`grid ${
            expanded ? "h-full min-h-0" : "min-h-[520px]"
          } place-items-center gap-3 p-8 text-center`}
          aria-busy="true"
          role="status"
        >
          <Spinner />
          <p className="text-sm text-[var(--ink-soft)]">
            กำลังเตรียมตัวแก้ไขเอกสาร…
          </p>
        </div>
      );
    }

    if (config.editorUrl) {
      return (
        <iframe
          title={title}
          src={config.editorUrl}
          className={
            expanded
              ? "h-full min-h-0 w-full border-0"
              : "h-[min(72vh,760px)] min-h-[520px] w-full border-0"
          }
        />
      );
    }

    return (
      <div
        className={
          expanded
            ? "h-full min-h-0 w-full"
            : "h-[min(72vh,760px)] min-h-[520px] w-full"
        }
      >
        <div ref={hostRef} className="h-full w-full" aria-label={title} />
      </div>
    );
  };

  return (
    <div
      aria-label={expanded ? title : undefined}
      aria-modal={expanded ? true : undefined}
      ref={surfaceRef}
      role={expanded ? "dialog" : undefined}
      className={
        expanded
          ? "fixed inset-0 z-[100] flex min-h-0 flex-col overflow-hidden overscroll-none bg-[var(--canvas)] p-3 sm:p-4"
          : "relative"
      }
    >
      <div className="flex shrink-0 justify-end pb-2">
        <Button
          aria-controls={surfaceId}
          aria-expanded={expanded}
          className="focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[var(--ink)]"
          onClick={() => setExpanded((value) => !value)}
          size="sm"
          type="button"
          variant="secondary"
        >
          {expanded ? "คืนค่าขนาดปกติ" : "ขยายพื้นที่เอกสาร"}
        </Button>
      </div>
      <div className={expanded ? "min-h-0 flex-1" : undefined} id={surfaceId}>
        {renderEditorContent()}
      </div>
    </div>
  );
};
