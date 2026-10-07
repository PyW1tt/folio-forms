// oxlint-disable prefer-await-to-callbacks node/callback-return unicorn/prefer-dom-node-remove prefer-destructuring unicorn/prefer-response-static-json no-await-in-loop unicorn/consistent-function-scoping no-plusplus unicorn/prefer-array-find -- VM harness intentionally mirrors browser callbacks and DOM shims.
import path from "node:path";
import { runInNewContext } from "node:vm";

import { buildPlugin } from "../../server/src/plugin-bundle.ts";

const pluginDirectory = path.dirname(import.meta.dirname);
const pluginArtifact = await buildPlugin(pluginDirectory);
const pluginSource = await pluginArtifact.text();

export const createHarness = ({
  action = "fill",
  capabilityResponses = [],
  clipboard,
  controls = [],
  editorSaveSupported = true,
  holdEditorSave = false,
  prefill,
  responses = [],
  selection,
  serializeCommands = false,
  tagAliases,
} = {}) => {
  const parentOrigin = "https://web.example.test";
  const bridgeId = "bridge-test-1";
  const messages = [];
  const messageListeners = [];
  const editorEvents = new Map();
  const elements = new Map();
  const requests = [];
  const methodCalls = [];
  const saveCalls = [];
  const timers = [];
  const queuedCapabilityResponses = [...capabilityResponses];
  const queuedResponses = [...responses];
  const selectionState = {
    currentControl: selection?.currentControl,
    properties: selection?.properties,
  };
  let pendingEditorSaveRequest;

  const createElement = (tagName = "div", id = "") => {
    const listeners = new Map();
    const element = {
      addEventListener(type, listener) {
        const callbacks = listeners.get(type) || [];
        callbacks.push(listener);
        listeners.set(type, callbacks);
      },
      append(...children) {
        for (const child of children) {
          if (!child) {
            continue;
          }
          child.parentNode = this;
          this.children.push(child);
        }
      },
      appendChild(child) {
        if (!child) {
          return child;
        }
        child.parentNode = this;
        this.children.push(child);
        return child;
      },
      checked: false,
      children: [],
      dataset: {},
      disabled: false,
      dispatchEvent(event = {}) {
        const callbacks = listeners.get(event.type) || [];
        const dispatched = {
          ...event,
          currentTarget: this,
          preventDefault:
            event.preventDefault ||
            (() => {
              // Browser event cancellation is not modeled by this DOM shim.
            }),
          target: event.target || this,
        };
        for (const callback of callbacks) {
          callback(dispatched);
        }
        return true;
      },
      hidden: false,
      id,
      remove() {
        if (!this.parentNode) {
          return;
        }
        this.parentNode.removeChild(this);
      },
      removeChild(child) {
        const index = this.children.indexOf(child);
        if (index !== -1) {
          this.children.splice(index, 1);
          child.parentNode = undefined;
        }
        return child;
      },
      replaceChildren(...children) {
        this.children = [];
        this.append(...children);
      },
      select() {
        // Text selection is not modeled by this DOM shim.
      },
      setAttribute(name, value) {
        this[name] = String(value);
      },
      style: {},
      tagName: tagName.toUpperCase(),
      textContent: "",
      value: "",
    };
    if (id) {
      elements.set(id, element);
    }
    return element;
  };

  const panelIds = [
    "field-apply-pointer",
    "field-panel",
    "field-picture-help",
    "field-policy-form",
    "field-prefill-policy",
    "field-required",
    "field-save",
    "field-schema-list",
    "field-schema-next",
    "field-schema-query",
    "field-schema-search",
    "field-selection-tag",
    "field-selection-type",
    "field-panel-status",
    "form-bridge-status",
  ];
  for (const id of panelIds) {
    createElement("div", id);
  }
  elements.get("field-panel").hidden = true;
  const statusElement = elements.get("form-bridge-status");

  const dispatch = (event) => {
    for (const listener of messageListeners) {
      listener(event);
    }
  };
  const parentWindow = {
    postMessage(message, targetOrigin) {
      messages.push({ message, targetOrigin });
      if (message.type === "editor-save-request") {
        pendingEditorSaveRequest = message;
        if (!holdEditorSave) {
          queueMicrotask(() => {
            dispatch({
              data: {
                ...message,
                saved: true,
                source: "folio-parent",
                type: "editor-save-response",
              },
              origin: parentOrigin,
              source: parentWindow,
            });
          });
        }
        return;
      }

      if (message.type !== "capability-request") {
        return;
      }

      const response = queuedCapabilityResponses.shift();
      if (response === undefined) {
        return;
      }

      queueMicrotask(() => {
        dispatch({
          data: {
            ...message,
            ...(typeof response === "string"
              ? { capability: response }
              : response),
            source: "folio-parent",
            type: "capability-response",
          },
          origin: parentOrigin,
          source: parentWindow,
        });
      });
    },
  };

  const plugin = {
    attachEditorEvent(name, callback) {
      const callbacks = editorEvents.get(name) || [];
      callbacks.push(callback);
      editorEvents.set(name, callbacks);
    },
    executeMethod(method, ...args) {
      methodCalls.push(method);
      const done = args.at(-1);
      if (typeof done !== "function") {
        return;
      }
      if (selection !== undefined) {
        if (method === "GetCurrentContentControl") {
          done(selectionState.currentControl);
          return;
        }
        if (method === "GetCurrentContentControlPr") {
          done(selectionState.properties);
          return;
        }
      }
      done();
    },
    info: {
      guid: "asc.test-plugin",
      options: {
        action,
        apiBase: "https://api.example.test/",
        bridgeId,
        documentKey: "document-key",
        parentOrigin,
        prefill,
        publicId: "public-id",
        responseId: "response-id",
        tagAliases,
        targetId: "target-id",
      },
    },
  };

  const fetch = (input, init = {}) => {
    const payload = queuedResponses.shift();
    if (payload === undefined) {
      throw new Error("Unexpected plugin request");
    }
    requests.push({
      body: init.body,
      credentials: init.credentials,
      headers: init.headers,
      method: init.method ?? "GET",
      url: String(input),
    });
    if (payload instanceof Response) {
      return payload;
    }
    if (
      payload &&
      typeof payload === "object" &&
      payload.httpStatus !== undefined
    ) {
      return new Response(JSON.stringify(payload.body ?? {}), {
        status: payload.httpStatus,
      });
    }
    return Response.json(payload);
  };

  const window = {
    Asc: { plugin },
    addEventListener(type, listener) {
      if (type === "message") {
        messageListeners.push(listener);
      }
    },
    clearTimeout(timer) {
      if (timer) {
        timer.cancelled = true;
      }
    },
    setTimeout(callback, milliseconds) {
      const timer = { callback, cancelled: false, milliseconds };
      timers.push(timer);
      if (milliseconds === 1000) {
        queueMicrotask(() => {
          if (!timer.cancelled) {
            return callback();
          }
        });
      }
      return timer;
    },
    top: parentWindow,
  };
  const officeControl = {
    GetClassType() {
      return "inlineLvlSdt";
    },
    GetFormType() {
      return selection?.formType || "";
    },
    GetInternalId() {
      return selectionState.currentControl;
    },
    GetTag() {
      return selectionState.properties?.Tag || "";
    },
    SetTag(tag) {
      selectionState.properties.Tag = tag;
      return true;
    },
  };
  const Api = {
    GetDocument() {
      return {
        GetAllContentControls() {
          return controls;
        },
        GetCurrentContentControl() {
          return selection ? officeControl : null;
        },
        GetCurrentContentControlPr() {
          return selectionState.properties || {};
        },
      };
    },
    Save() {
      saveCalls.push("save");
      return true;
    },
  };
  plugin.callCommand = function callCommand(...args) {
    const command = args[0];
    const done = args.at(-1);
    if (typeof done !== "function") {
      return;
    }
    if (serializeCommands && typeof command === "function") {
      done(
        runInNewContext(`(${command.toString()})()`, {
          Api,
          Asc: { scope: window.Asc.scope },
        })
      );
      return;
    }
    if (
      command?.name === "getCurrentContentControlCommand" &&
      selectionState.properties
    ) {
      done(command());
      return;
    }
    if (
      command?.name === "setCurrentContentControlTagCommand" &&
      selectionState.properties
    ) {
      const tag = window.Asc.scope?.formBridgeSelectionTag;
      selectionState.properties.Tag = tag;
      done(JSON.stringify({ ok: true, tag }));
      return;
    }
    if (typeof command === "function") {
      done(command());
      return;
    }
    done(JSON.stringify({ name: "Ada", value: "example" }));
  };
  const document = {
    body: createElement("body"),
    createElement,
    getElementById(id) {
      return elements.get(id) || null;
    },
    querySelector(selector) {
      if (selector.startsWith("#")) {
        return elements.get(selector.slice(1)) || null;
      }
      return statusElement;
    },
  };
  const context = {
    Api,
    Asc: window.Asc,
    Headers,
    Response,
    document,
    fetch,
    navigator: clipboard ? { clipboard } : undefined,
    queueMicrotask,
    window,
  };

  runInNewContext(pluginSource, context, {
    filename: "apps/onlyoffice-plugin/dist/plugin.js",
  });
  window.Asc.plugin.init();

  return {
    bridgeId,
    dispatch,
    editorEvents,
    editorSaveSupported,
    element(id) {
      return elements.get(id);
    },
    emitEditorEvent(name, value) {
      for (const callback of editorEvents.get(name) || []) {
        callback(value);
      }
    },
    expireCapabilityRequests() {
      for (const timer of timers) {
        if (timer.milliseconds !== 5000 || timer.cancelled) {
          continue;
        }
        timer.cancelled = true;
        timer.callback();
      }
    },
    finishEditorSave(error) {
      dispatch({
        data: {
          ...pendingEditorSaveRequest,
          ...(error ? { error } : { saved: true }),
          source: "folio-parent",
          type: "editor-save-response",
        },
        origin: parentOrigin,
        source: parentWindow,
      });
    },
    messages,
    methodCalls,
    parentOrigin,
    parentWindow,
    requests,
    saveCalls,
    setSelection(nextSelection) {
      selectionState.currentControl = nextSelection?.currentControl;
      selectionState.properties = nextSelection?.properties;
    },
    statusElement,
    window,
  };
};

export const completedOperation = (result = {}) => ({
  operation: { result, status: "completed" },
});
export const acknowledgeBridge = (harness) => {
  harness.dispatch({
    data: {
      bridgeId: harness.bridgeId,
      editorSaveSupported: harness.editorSaveSupported,
      source: "folio-parent",
      type: "bridge-ack",
    },
    origin: harness.parentOrigin,
    source: harness.parentWindow,
  });
};
export const flushPlugin = async () => {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    await Promise.resolve();
  }
};

export const selectedControl = (
  tag = "title",
  controlType = "text",
  internalId = "control-1"
) => ({
  currentControl: internalId,
  properties: {
    InternalId: internalId,
    Tag: tag,
    Type: controlType,
  },
});
