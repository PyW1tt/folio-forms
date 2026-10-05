import { API_ORIGIN } from "@/lib/api";

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

export interface EditorLease {
  id: string;
  expiresAt: string;
  releaseUrl: string;
  renewUrl: string;
}

export const leaseRenewalIntervalMs = 30_000;

export const editorConfigPath = (configUrl: string): string =>
  configUrl.startsWith("http") ? configUrl.replace(API_ORIGIN, "") : configUrl;

export interface EditorConfig {
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

export interface CapabilityRequestMessage {
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

export const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
export const recordOrEmpty = (value: unknown): Record<string, unknown> =>
  isRecord(value) ? value : {};

export const isNonEmptyString = (value: unknown): value is string =>
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

export const isBridgeReadyMessage = (
  value: unknown,
  bridgeId: string
): value is BridgeReadyMessage =>
  isRecord(value) &&
  value.bridgeId === bridgeId &&
  value.source === "form-bridge" &&
  value.type === "bridge-ready";

export const parseDirtyStateMessage = (
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

export const parseCapabilityRequest = (
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

export const parseFieldSelectionMessage = (
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

export const parseOperationMessage = (
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

export const acknowledgeBridge = (
  source: MessageEventSource,
  pluginOrigin: string,
  bridgeId: string,
  editorSaveSupported: boolean
) => {
  try {
    (source as Window).postMessage(
      {
        bridgeId,
        editorSaveSupported,
        source: "folio-parent",
        type: "bridge-ack",
      },
      pluginOrigin
    );
  } catch {
    // The plugin may close its frame while the handshake is in flight.
  }
};
