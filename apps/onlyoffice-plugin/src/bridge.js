// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import {
  ACTIONS,
  BRIDGE_MESSAGE_SOURCE,
  PARENT_MESSAGE_SOURCE,
  BRIDGE_READY_TYPE,
  BRIDGE_ACK_TYPE,
  CAPABILITY_ACTIONS,
  CAPABILITY_REQUEST_TYPE,
  CAPABILITY_RESPONSE_TYPE,
  CAPABILITY_REQUEST_TIMEOUT_MS,
  DIRTY_STATE_TYPE,
  RUN_ACTION_TYPE,
  CLEAR_DIRTY_TYPE,
  OPERATION_MESSAGE_TYPE,
  EDITOR_SAVE_REQUEST_TYPE,
  EDITOR_SAVE_RESPONSE_TYPE,
} from "./constants.js";
import { saveDocumentCommand } from "./document-commands.js";
import { requireOption, isRecord, errorMessage } from "./values.js";

export function createBridge(options, office, callbacks) {
  const {
    runAction,
    loadFieldRules,
    getFieldSelection,
    setStatus,
    setPanelStatus,
  } = callbacks;
  let documentDirty = false;
  let pendingFieldSelection;
  let bridgeAcknowledged = false;
  let bridgeMessageListenerAttached = false;
  let editorSaveSupported = false;
  let editorSaveRequestSequence = 0;
  let pendingEditorSave;
  let bridgeReadySent = false;
  let capabilityRequestSequence = 0;
  const pendingCapabilityRequests = new Map();

  function parentWindow() {
    const target = window.top;
    if (!target || typeof target.postMessage !== "function") {
      throw new Error("The editor parent window is unavailable");
    }
    return target;
  }

  function postBridgeMessage(message) {
    const bridgeId = requireOption(options.bridgeId, "bridgeId");
    const parentOrigin = requireOption(options.parentOrigin, "parentOrigin");
    parentWindow().postMessage(
      {
        ...message,
        bridgeId,
      },
      parentOrigin
    );
  }

  async function flushEditorChanges() {
    if (!bridgeAcknowledged) {
      throw new Error("The editor host bridge is not acknowledged");
    }
    if (!editorSaveSupported) {
      throw new Error("Native editor save synchronization is unavailable");
    }
    // Api.Save requests asc_Save after the command callback; it is not completion.
    if ((await office.callCommandResult(saveDocumentCommand)) !== true) {
      throw new Error("The native editor could not request a document save");
    }
    editorSaveRequestSequence += 1;
    const requestId = `editor-save-${editorSaveRequestSequence}`;
    // This host round trip runs after the native command and its queued state events.
    await new Promise((resolve, reject) => {
      pendingEditorSave = { reject, requestId, resolve };
      try {
        postBridgeMessage({
          requestId,
          source: BRIDGE_MESSAGE_SOURCE,
          type: EDITOR_SAVE_REQUEST_TYPE,
        });
      } catch (error) {
        pendingEditorSave = undefined;
        reject(error);
      }
    });
  }

  function postFieldSelection(message) {
    try {
      postBridgeMessage(message);
    } catch {
      setPanelStatus("ไม่สามารถแจ้งการเลือกฟิลด์ได้", "error");
    }
  }

  function isActionAllowedInMode(action) {
    switch (options.action) {
      case ACTIONS.TEMPLATE_EDIT: {
        return (
          action === ACTIONS.CONFIGURE_FIELDS ||
          action === ACTIONS.PUBLISH ||
          action === ACTIONS.SAVE_TEMPLATE
        );
      }
      case ACTIONS.FILL:
      case ACTIONS.DRAFT: {
        return action === ACTIONS.SAVE_DRAFT || action === ACTIONS.SUBMIT;
      }
      case ACTIONS.CORRECTION: {
        return action === ACTIONS.SAVE_CORRECTION;
      }
      case ACTIONS.SUBMIT: {
        return action === ACTIONS.SUBMIT;
      }
      default: {
        return false;
      }
    }
  }

  function requestActionCapability(action) {
    if (
      !CAPABILITY_ACTIONS.includes(action) ||
      !isActionAllowedInMode(action)
    ) {
      return Promise.reject(
        new Error(`Unsupported form action: ${action || "none"}`)
      );
    }

    if (!bridgeAcknowledged) {
      return Promise.reject(
        new Error("The editor host bridge is not acknowledged")
      );
    }

    capabilityRequestSequence += 1;
    const requestId = `capability-${capabilityRequestSequence}`;

    return new Promise((resolve, reject) => {
      const pending = {
        action,
        reject,
        resolve,
        timeoutId: undefined,
      };
      pendingCapabilityRequests.set(requestId, pending);

      const timeout = () => {
        if (pendingCapabilityRequests.get(requestId) !== pending) {
          return;
        }

        pendingCapabilityRequests.delete(requestId);
        reject(new Error(`Timed out waiting for ${action} capability`));
      };

      try {
        pending.timeoutId = window.setTimeout(
          timeout,
          CAPABILITY_REQUEST_TIMEOUT_MS
        );
        postBridgeMessage({
          action,
          requestId,
          source: BRIDGE_MESSAGE_SOURCE,
          type: CAPABILITY_REQUEST_TYPE,
        });
      } catch (error) {
        pendingCapabilityRequests.delete(requestId);
        window.clearTimeout(pending.timeoutId);
        reject(error);
      }
    });
  }

  function setDirtyState(dirty) {
    if (documentDirty === dirty) {
      return;
    }
    documentDirty = dirty;
    try {
      postBridgeMessage({
        dirty,
        source: BRIDGE_MESSAGE_SOURCE,
        type: DIRTY_STATE_TYPE,
      });
    } catch {
      // The editor can run without a host frame.
    }
  }

  function settleCapabilityRequest(requestId, settle, value) {
    const pending = pendingCapabilityRequests.get(requestId);
    if (!pending) {
      return;
    }

    pendingCapabilityRequests.delete(requestId);
    window.clearTimeout(pending.timeoutId);
    settle(value);
  }

  function handleCapabilityResponse(message) {
    if (
      typeof message.requestId !== "string" ||
      !message.requestId ||
      !CAPABILITY_ACTIONS.includes(message.action)
    ) {
      return;
    }

    const pending = pendingCapabilityRequests.get(message.requestId);
    if (!pending || pending.action !== message.action) {
      return;
    }

    const hasCapability = Object.hasOwn(message, "capability");
    const hasError = Object.hasOwn(message, "error");
    if (hasCapability === hasError) {
      return;
    }

    if (hasCapability) {
      if (
        typeof message.capability !== "string" ||
        !message.capability.trim()
      ) {
        return;
      }

      settleCapabilityRequest(
        message.requestId,
        pending.resolve,
        message.capability
      );
      return;
    }

    if (typeof message.error !== "string" || !message.error.trim()) {
      return;
    }

    settleCapabilityRequest(
      message.requestId,
      pending.reject,
      new Error(message.error)
    );
  }

  function notifyParent(action, status, operationId, payload, error) {
    if (!bridgeAcknowledged) {
      return;
    }

    try {
      postBridgeMessage({
        action,
        error: error || undefined,
        operation: isRecord(payload?.operation) ? payload.operation : undefined,
        operationId: operationId || undefined,
        source: BRIDGE_MESSAGE_SOURCE,
        status,
        type: OPERATION_MESSAGE_TYPE,
      });
    } catch {
      // The editor can run without a host frame.
    }
  }

  function handleParentMessage(event) {
    const message = event?.data;
    let topWindow;

    try {
      topWindow = window.top;
    } catch {
      return;
    }

    if (
      !topWindow ||
      event?.source !== topWindow ||
      event.origin !== options.parentOrigin ||
      !isRecord(message) ||
      message.source !== PARENT_MESSAGE_SOURCE ||
      message.bridgeId !== options.bridgeId ||
      (message.type !== BRIDGE_ACK_TYPE &&
        message.type !== CAPABILITY_RESPONSE_TYPE &&
        message.type !== RUN_ACTION_TYPE &&
        message.type !== CLEAR_DIRTY_TYPE &&
        message.type !== EDITOR_SAVE_RESPONSE_TYPE)
    ) {
      return;
    }
    if (message.type === BRIDGE_ACK_TYPE) {
      bridgeAcknowledged = true;
      editorSaveSupported = message.editorSaveSupported === true;
      if (pendingFieldSelection) {
        const pending = pendingFieldSelection;
        pendingFieldSelection = undefined;
        postFieldSelection(pending);
      }
      if (
        options.action === ACTIONS.TEMPLATE_EDIT &&
        getFieldSelection()?.controlKey &&
        !getFieldSelection().rulesLoaded
      ) {
        void loadFieldRules();
      }
      return;
    }

    if (!bridgeAcknowledged) {
      return;
    }

    if (message.type === EDITOR_SAVE_RESPONSE_TYPE) {
      const pending = pendingEditorSave;
      if (!pending || message.requestId !== pending.requestId) {
        return;
      }
      if (message.saved === true && !Object.hasOwn(message, "error")) {
        pendingEditorSave = undefined;
        pending.resolve();
      } else if (
        !Object.hasOwn(message, "saved") &&
        typeof message.error === "string" &&
        message.error.trim()
      ) {
        pendingEditorSave = undefined;
        pending.reject(new Error(message.error));
      }
      return;
    }

    if (message.type === CLEAR_DIRTY_TYPE) {
      setDirtyState(false);
      return;
    }

    if (message.type === RUN_ACTION_TYPE) {
      if (
        message.action !== ACTIONS.SAVE_DRAFT &&
        message.action !== ACTIONS.SAVE_CORRECTION
      ) {
        return;
      }

      void runAction(
        message.action,
        typeof message.reason === "string" ? message.reason : ""
      );
      return;
    }

    handleCapabilityResponse(message);
  }

  function startBridge() {
    if (!bridgeMessageListenerAttached) {
      window.addEventListener("message", handleParentMessage);
      bridgeMessageListenerAttached = true;
    }

    if (bridgeReadySent) {
      return;
    }

    try {
      postBridgeMessage({
        source: BRIDGE_MESSAGE_SOURCE,
        type: BRIDGE_READY_TYPE,
      });
      bridgeReadySent = true;
    } catch (error) {
      setStatus(
        `Could not connect to editor host: ${errorMessage(error)}`,
        "error"
      );
    }
  }

  function sendFieldSelection(message) {
    if (!bridgeAcknowledged) {
      pendingFieldSelection = message;
      return message;
    }

    postFieldSelection(message);
    return message;
  }

  return {
    flushEditorChanges,
    notifyParent,
    postBridgeMessage,
    requestActionCapability,
    sendFieldSelection,
    setDirtyState,
    startBridge,
  };
}
