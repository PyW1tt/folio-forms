// oxlint-disable prefer-await-to-callbacks node/callback-return unicorn/prefer-dom-node-remove prefer-destructuring unicorn/prefer-response-static-json no-await-in-loop unicorn/consistent-function-scoping no-plusplus unicorn/prefer-array-find -- VM harness intentionally mirrors browser callbacks and DOM shims.
import { expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import { runInNewContext } from "node:vm";

const pluginSource = await readFile(
  new URL("plugin.js", import.meta.url),
  "utf-8"
);

const createHarness = ({
  action = "fill",
  capabilityResponses = [],
  clipboard,
  controls = [],
  prefill,
  responses = [],
  selection,
} = {}) => {
  const parentOrigin = "https://web.example.test";
  const bridgeId = "bridge-test-1";
  const messages = [];
  const messageListeners = [];
  const editorEvents = new Map();
  const elements = new Map();
  const requests = [];
  const methodCalls = [];
  const timers = [];
  const queuedCapabilityResponses = [...capabilityResponses];
  const queuedResponses = [...responses];
  const selectionState = {
    currentControl: selection?.currentControl,
    properties: selection?.properties,
  };

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
          preventDefault: event.preventDefault || (() => {}),
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
      select() {},
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
    callCommand(...args) {
      const command = args[0];
      const done = args.at(-1);
      if (typeof done !== "function") {
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
    filename: "apps/onlyoffice-plugin/plugin.js",
  });
  window.Asc.plugin.init();

  return {
    bridgeId,
    dispatch,
    editorEvents,
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
    messages,
    methodCalls,
    parentOrigin,
    parentWindow,
    requests,
    setSelection(nextSelection) {
      selectionState.currentControl = nextSelection?.currentControl;
      selectionState.properties = nextSelection?.properties;
    },
    statusElement,
    window,
  };
};

const completedOperation = (result = {}) => ({
  operation: { result, status: "completed" },
});
const acknowledgeBridge = (harness) => {
  harness.dispatch({
    data: {
      bridgeId: harness.bridgeId,
      source: "folio-parent",
      type: "bridge-ack",
    },
    origin: harness.parentOrigin,
    source: harness.parentWindow,
  });
};
const flushPlugin = async () => {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    await Promise.resolve();
  }
};

const selectedControl = (
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

test("reports dirty state after a document content change", () => {
  const harness = createHarness();
  acknowledgeBridge(harness);

  harness.emitEditorEvent("onChangeContentControl");

  expect(harness.messages.at(-1)).toEqual({
    message: {
      bridgeId: harness.bridgeId,
      dirty: true,
      source: "form-bridge",
      type: "dirty-state",
    },
    targetOrigin: harness.parentOrigin,
  });
});

test("clears dirty state after successful save and submit actions", async () => {
  const harness = createHarness({
    action: "draft",
    capabilityResponses: ["save-draft-capability", "submit-capability"],
    responses: [{ ok: true }, { ok: true }],
  });
  acknowledgeBridge(harness);

  harness.emitEditorEvent("onChangeContentControl");
  await expect(
    harness.window.FormBridge.runAction("save-draft")
  ).resolves.toMatchObject({ ok: true });
  harness.emitEditorEvent("onChangeContentControl");
  await expect(
    harness.window.FormBridge.runAction("submit")
  ).resolves.toMatchObject({ ok: true });

  expect(
    harness.messages
      .filter(({ message }) => message.type === "dirty-state")
      .map(({ message }) => message.dirty)
  ).toEqual([true, false, true, false]);
});

test("accepts only authenticated parent dirty commands", async () => {
  const harness = createHarness({
    action: "draft",
    capabilityResponses: ["save-draft-capability"],
    responses: [{ ok: true }],
  });
  acknowledgeBridge(harness);
  harness.emitEditorEvent("onChangeContentControl");

  const validRunAction = {
    action: "save-draft",
    bridgeId: harness.bridgeId,
    source: "folio-parent",
    type: "run-action",
  };
  const invalidEvents = [
    {
      data: validRunAction,
      origin: "https://wrong.example.test",
      source: harness.parentWindow,
    },
    {
      data: validRunAction,
      origin: harness.parentOrigin,
      source: {},
    },
    {
      data: { ...validRunAction, bridgeId: "wrong-bridge" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validRunAction, action: "submit" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: {
        bridgeId: harness.bridgeId,
        source: "attacker",
        type: "clear-dirty",
      },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
  ];

  for (const event of invalidEvents) {
    harness.dispatch(event);
  }
  await flushPlugin();
  expect(harness.requests).toHaveLength(0);
  expect(
    harness.messages
      .filter(({ message }) => message.type === "dirty-state")
      .map(({ message }) => message.dirty)
  ).toEqual([true]);

  harness.dispatch({
    data: {
      bridgeId: harness.bridgeId,
      source: "folio-parent",
      type: "clear-dirty",
    },
    origin: harness.parentOrigin,
    source: harness.parentWindow,
  });
  harness.emitEditorEvent("onChangeContentControl");
  harness.dispatch({
    data: validRunAction,
    origin: harness.parentOrigin,
    source: harness.parentWindow,
  });
  await flushPlugin();

  expect(harness.requests).toHaveLength(1);
  expect(
    harness.messages
      .filter(({ message }) => message.type === "dirty-state")
      .map(({ message }) => message.dirty)
  ).toEqual([true, false, true, false]);
});

test("runs correction saves with a reason and correction endpoint", async () => {
  const harness = createHarness({
    action: "correction",
    capabilityResponses: ["save-correction-capability", "poll-capability"],
    responses: [
      {
        operationCapability: "poll-capability",
        operationId: "correction-operation",
      },
      completedOperation({ correctionId: "correction-1", revision: 1 }),
    ],
  });
  expect(harness.methodCalls).toContain("SetEditingRestrictions");
  const restrictionsBeforeReady = harness.methodCalls.filter(
    (method) => method === "SetEditingRestrictions"
  ).length;
  harness.emitEditorEvent("onDocumentContentReady");
  await flushPlugin();
  expect(
    harness.methodCalls.filter((method) => method === "SetEditingRestrictions")
      .length
  ).toBeGreaterThan(restrictionsBeforeReady);
  acknowledgeBridge(harness);

  await expect(
    harness.window.FormBridge.runAction("save-correction", "แก้ไขตามเอกสารต้นฉบับ")
  ).resolves.toMatchObject({ ok: true });

  expect(harness.requests.map(({ method, url }) => [method, url])).toEqual([
    [
      "POST",
      "https://api.example.test/api/admin/results/response-id/correction",
    ],
    ["GET", "https://api.example.test/api/operations/correction-operation"],
  ]);
  expect(JSON.parse(harness.requests[0].body)).toEqual({
    data: {},
    documentKey: "document-key",
    reason: "แก้ไขตามเอกสารต้นฉบับ",
  });
  expect(harness.requests[0].headers.get("x-editor-capability")).toBe(
    "save-correction-capability"
  );
});
test("extracts scalar form values with plugin contract semantics", async () => {
  const control = ({
    checkbox = false,
    checked = false,
    date = false,
    dateValue = null,
    dropdown = false,
    items = [],
    combo = false,
    formType = "",
    picture = false,
    tag,
    text,
  }) => ({
    GetClassType: () => "inlineLvlSdt",
    GetDate: () => dateValue,
    GetDropdownList: () => ({ GetAllItems: () => items }),
    GetFormType: () => formType,
    GetRange: () => ({ GetText: () => text }),
    GetTag: () => tag,
    IsCheckBox: () => checkbox,
    IsCheckBoxChecked: () => checked,
    IsComboBox: () => combo,
    IsDatePicker: () => date,
    IsDropDownList: () => dropdown,
    IsPicture: () => picture,
  });

  const harness = createHarness({
    action: "draft",
    capabilityResponses: ["save-draft-capability"],
    controls: [
      control({ tag: "notes", text: "line one\nline two" }),
      control({ checkbox: true, checked: true, tag: "accept_terms" }),
      control({
        date: true,
        dateValue: new Date(2026, 8, 15),
        tag: "start_date",
      }),
      control({
        dropdown: true,
        items: [
          {
            GetText: () => "Engineering",
            GetValue: () => "engineering",
          },
        ],
        tag: "department",
        text: "Engineering",
      }),
      control({
        combo: true,
        items: [
          {
            GetText: () => "Known",
            GetValue: () => "known",
          },
        ],
        tag: "custom",
        text: "Custom value",
      }),
      control({
        formType: "picture",
        picture: true,
        tag: "photo",
        text: "picture bytes never become scalar data",
      }),
    ],
    responses: [
      { operationCapability: "save-operation-capability", operationId: "save" },
      completedOperation({ saved: true }),
    ],
  });
  acknowledgeBridge(harness);
  await expect(
    harness.window.FormBridge.runAction("save-draft")
  ).resolves.toMatchObject({ ok: true });
  const requestBody = JSON.parse(harness.requests[0].body);
  expect(requestBody.data).toEqual({
    accept_terms: true,
    custom: "Custom value",
    department: "engineering",
    notes: "line one\nline two",
    start_date: "2026-09-15",
  });
});

test("skips native picture controls during prefill without mutating them", async () => {
  let mutationCount = 0;
  const pictureControl = {
    AddText: () => {
      mutationCount += 1;
    },
    GetClassType: () => "inlineLvlSdt",
    GetFormType: () => "picture",
    GetRange: () => ({ SetText: () => mutationCount++ }),
    GetTag: () => "photo",
    IsPicture: () => true,
    RemoveAllElements: () => {
      mutationCount += 1;
    },
    SetLock: () => {
      mutationCount += 1;
    },
  };
  const harness = createHarness({ controls: [pictureControl] });

  await expect(
    harness.window.FormBridge.applyPrefill({
      policies: { photo: "lock-when-available" },
      values: { photo: "https://example.test/image.png" },
    })
  ).resolves.toEqual({
    applied: [],
    failed: [],
    skipped: ["photo"],
  });
  expect(mutationCount).toBe(0);
});
test("locks available Prefill and leaves missing values editable", async () => {
  const controls = ["trusted", "missing"].map((tag) => {
    let lock;
    let text = "";
    return {
      AddText: (value) => {
        text += value;
      },
      GetClassType: () => "inlineLvlSdt",
      GetRange: () => ({ GetText: () => text }),
      GetTag: () => tag,
      RemoveAllElements: () => {
        text = "";
      },
      SetLock: (value) => {
        lock = value;
      },
      read: () => ({ lock, text }),
    };
  });
  const harness = createHarness({ controls });
  await expect(
    harness.window.FormBridge.applyPrefill({
      editableFields: { trusted: false },
      values: { trusted: "Trusted value" },
    })
  ).resolves.toEqual({
    applied: ["trusted"],
    failed: [],
    skipped: ["missing"],
  });
  expect(controls.map((control) => control.read())).toEqual([
    { lock: "sdtContentLocked", text: "Trusted value" },
    { lock: undefined, text: "" },
  ]);
});
test("applies a saved scalar response and reports Thai action status", async () => {
  const createMutableControl = ({ kind, tag, items = [] }) => {
    let checked = false;
    let dateValue = null;
    let textValue = "";
    const optionControls = items.map(({ display, value }) => ({
      GetText: () => display,
      GetValue: () => value,
      Select: () => {
        textValue = display;
      },
    }));
    return {
      AddText: (value) => {
        textValue += value;
      },
      GetClassType: () => "inlineLvlSdt",
      GetDate: () => dateValue,
      GetDropdownList: () => ({
        GetAllItems: () => optionControls,
      }),
      GetRange: () => ({ GetText: () => textValue }),
      GetTag: () => tag,
      IsCheckBox: () => kind === "checkbox",
      IsCheckBoxChecked: () => checked,
      IsComboBox: () => kind === "combo",
      IsDatePicker: () => kind === "date",
      IsDropDownList: () => kind === "dropdown",
      RemoveAllElements: () => {
        textValue = "";
      },
      SetCheckBoxChecked: (value) => {
        checked = value;
      },
      SetDate: (value) => {
        dateValue = value;
      },
      SetLock: () => {},
      read: () => ({ checked, dateValue, textValue }),
    };
  };
  const controls = [
    createMutableControl({ kind: "text", tag: "notes" }),
    createMutableControl({ kind: "checkbox", tag: "accept_terms" }),
    createMutableControl({ kind: "date", tag: "start_date" }),
    createMutableControl({
      items: [{ display: "Engineering", value: "engineering" }],
      kind: "dropdown",
      tag: "department",
    }),
    createMutableControl({
      items: [{ display: "Known", value: "known" }],
      kind: "combo",
      tag: "custom",
    }),
  ];
  const harness = createHarness({
    capabilityResponses: ["save-draft-capability"],
    controls,
    prefill: {
      values: {
        accept_terms: true,
        custom: "Custom value",
        department: "engineering",
        notes: "line one\nline two",
        start_date: "2026-09-15",
      },
    },
    responses: [
      { operationCapability: "save-operation-capability", operationId: "save" },
      completedOperation({ saved: true }),
    ],
  });
  acknowledgeBridge(harness);
  await expect(
    harness.window.FormBridge.runAction("save-draft")
  ).resolves.toMatchObject({ ok: true });
  expect(controls.map((control) => control.read())).toEqual([
    { checked: false, dateValue: null, textValue: "line one\nline two" },
    { checked: true, dateValue: null, textValue: "" },
    {
      checked: false,
      dateValue: new Date(2026, 8, 15),
      textValue: "",
    },
    { checked: false, dateValue: null, textValue: "Engineering" },
    { checked: false, dateValue: null, textValue: "Custom value" },
  ]);
  expect(JSON.parse(harness.requests[0].body).data).toEqual({
    accept_terms: true,
    custom: "Custom value",
    department: "engineering",
    notes: "line one\nline two",
    start_date: "2026-09-15",
  });
  expect(harness.statusElement.textContent).toBe("บันทึกฉบับร่าง สำเร็จ");

  const errorHarness = createHarness({
    action: "draft",
    capabilityResponses: ["save-draft-capability"],
    responses: [
      {
        body: { error: "invalid_response_data", message: "invalid" },
        httpStatus: 422,
      },
    ],
  });
  acknowledgeBridge(errorHarness);
  await expect(
    errorHarness.window.FormBridge.runAction("save-draft")
  ).resolves.toMatchObject({ ok: false });
  expect(errorHarness.statusElement.textContent).toContain("ไม่สำเร็จ");
});
test("uses fresh capabilities and an exact acknowledged bridge", async () => {
  const harness = createHarness({
    capabilityResponses: [
      "fresh-save-draft-capability",
      "fresh-submit-capability",
    ],
    responses: [
      {
        operationCapability: "save-operation-capability",
        operationId: "save-operation",
      },
      { operation: { status: "processing" } },
      completedOperation({ saved: true }),
      {
        operationCapability: "submit-operation-capability",
        operationId: "submit-operation",
      },
      { operation: { status: "processing" } },
      completedOperation({ submissionId: "submission-1" }),
    ],
  });

  expect(harness.messages).toHaveLength(1);
  expect(harness.messages[0]).toEqual({
    message: {
      bridgeId: harness.bridgeId,
      source: "form-bridge",
      type: "bridge-ready",
    },
    targetOrigin: harness.parentOrigin,
  });

  acknowledgeBridge(harness);

  const firstAction = harness.window.FormBridge.runAction("save-draft");
  await expect(harness.window.FormBridge.runAction("submit")).resolves.toEqual({
    ignored: true,
    ok: false,
  });
  await expect(firstAction).resolves.toMatchObject({ ok: true });
  await expect(
    harness.window.FormBridge.runAction("submit")
  ).resolves.toMatchObject({ ok: true });

  const capabilityRequests = harness.messages.filter(
    ({ message }) => message.type === "capability-request"
  );
  expect(capabilityRequests).toEqual([
    {
      message: {
        action: "save-draft",
        bridgeId: harness.bridgeId,
        requestId: "capability-1",
        source: "form-bridge",
        type: "capability-request",
      },
      targetOrigin: harness.parentOrigin,
    },
    {
      message: {
        action: "submit",
        bridgeId: harness.bridgeId,
        requestId: "capability-2",
        source: "form-bridge",
        type: "capability-request",
      },
      targetOrigin: harness.parentOrigin,
    },
  ]);

  expect(
    harness.requests.map(({ headers, method, url }) => [
      method,
      url,
      headers.get("x-editor-capability"),
      headers.has("authorization"),
    ])
  ).toEqual([
    [
      "POST",
      "https://api.example.test/api/forms/public-id/draft",
      "fresh-save-draft-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/save-operation",
      "save-operation-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/save-operation",
      "save-operation-capability",
      false,
    ],
    [
      "POST",
      "https://api.example.test/api/forms/public-id/submit",
      "fresh-submit-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/submit-operation",
      "submit-operation-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/submit-operation",
      "submit-operation-capability",
      false,
    ],
  ]);
  expect(
    harness.requests.every(({ credentials }) => credentials === "omit")
  ).toBe(true);

  const operationMessages = harness.messages.filter(
    ({ message }) => message.type === "operation"
  );
  expect(operationMessages).toHaveLength(4);
  expect(
    operationMessages.map(({ message, targetOrigin }) => [
      message.action,
      message.status,
      message.bridgeId,
      targetOrigin,
    ])
  ).toEqual([
    ["save-draft", "pending", harness.bridgeId, harness.parentOrigin],
    ["save-draft", "completed", harness.bridgeId, harness.parentOrigin],
    ["submit", "pending", harness.bridgeId, harness.parentOrigin],
    ["submit", "completed", harness.bridgeId, harness.parentOrigin],
  ]);
});

test("rejects parent acknowledgements with the wrong source, origin, or bridge id", async () => {
  const invalidEvents = [
    (harness) => ({
      data: {
        bridgeId: harness.bridgeId,
        source: "folio-parent",
        type: "bridge-ack",
      },
      origin: harness.parentOrigin,
      source: {},
    }),
    (harness) => ({
      data: {
        bridgeId: harness.bridgeId,
        source: "folio-parent",
        type: "bridge-ack",
      },
      origin: "https://wrong.example.test",
      source: harness.parentWindow,
    }),
    (harness) => ({
      data: {
        bridgeId: "wrong-bridge",
        source: "folio-parent",
        type: "bridge-ack",
      },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    }),
  ];

  await Promise.all(
    invalidEvents.map(async (createEvent) => {
      const harness = createHarness({
        action: "template-edit",
      });
      harness.dispatch(createEvent(harness));
      await expect(
        harness.window.FormBridge.runAction("save-template")
      ).resolves.toMatchObject({ ok: false });
      expect(harness.requests).toHaveLength(0);
      expect(
        harness.messages.filter(({ message }) => message.type === "operation")
      ).toHaveLength(0);
    })
  );
});

test("fails safely when capability renewal times out or returns an error", async () => {
  const timeoutHarness = createHarness({ action: "template-edit" });
  acknowledgeBridge(timeoutHarness);
  const timedOutAction =
    timeoutHarness.window.FormBridge.runAction("save-template");

  expect(timeoutHarness.requests).toHaveLength(0);
  timeoutHarness.expireCapabilityRequests();
  await expect(timedOutAction).resolves.toMatchObject({
    error: "Timed out waiting for save-template capability",
    ok: false,
  });

  const errorHarness = createHarness({
    action: "template-edit",
    capabilityResponses: [{ error: "capability denied" }],
  });
  acknowledgeBridge(errorHarness);

  await expect(
    errorHarness.window.FormBridge.runAction("save-template")
  ).resolves.toMatchObject({
    error: "capability denied",
    ok: false,
  });
  expect(errorHarness.requests).toHaveLength(0);
});

test("ignores forged or malformed capability responses", async () => {
  const harness = createHarness({
    action: "template-edit",
    responses: [{ ok: true }],
  });
  acknowledgeBridge(harness);

  const action = harness.window.FormBridge.runAction("save-template");
  const request = harness.messages.find(
    ({ message }) => message.type === "capability-request"
  );
  const validResponse = {
    ...request.message,
    capability: "forged-capability",
    source: "folio-parent",
    type: "capability-response",
  };
  const invalidEvents = [
    {
      data: validResponse,
      origin: harness.parentOrigin,
      source: {},
    },
    {
      data: validResponse,
      origin: "https://wrong.example.test",
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, bridgeId: "wrong-bridge" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, requestId: "unknown-request" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, action: "submit" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, action: "unknown-action" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, source: "attacker" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, type: "unknown-response" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, capability: "" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, capability: {} },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: {
        ...harness.messages.find(
          ({ message }) => message.type === "capability-request"
        ).message,
        source: "folio-parent",
        type: "capability-response",
      },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, capability: "forged", error: "also forged" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
  ];

  let settled = false;
  void action.then(() => {
    settled = true;
  });
  for (const event of invalidEvents) {
    harness.dispatch(event);
  }
  await Promise.resolve();

  expect(settled).toBe(false);
  expect(harness.requests).toHaveLength(0);
  harness.expireCapabilityRequests();
  await expect(action).resolves.toMatchObject({ ok: false });
});

test("uses the public Form identifier for Admin template actions", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["action-capability"],
    responses: [{ ok: true }],
  });
  acknowledgeBridge(harness);

  await expect(
    harness.window.FormBridge.runAction("save-template")
  ).resolves.toMatchObject({ ok: true });
  expect(harness.requests[0]?.url).toBe(
    "https://api.example.test/api/admin/forms/public-id/save"
  );
  expect(harness.requests[0]?.url).not.toContain("form-id");
});

test("configures the selected field with exact schema pointers and policy", async () => {
  let copiedPointer = null;
  const pointer = "/account/address/city";
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: [
      "field-capability-1",
      "field-capability-2",
      "field-capability-3",
      "field-capability-4",
    ],
    clipboard: {
      writeText(value) {
        copiedPointer = value;
        return Promise.resolve();
      },
    },
    responses: [
      { rules: [] },
      {
        items: [{ pointer, type: "string" }],
        nextCursor: "schema-cursor-1",
      },
      {
        items: [{ pointer: "/account/address/country", type: "string" }],
        nextCursor: null,
      },
      {
        rule: {
          prefillPointer: pointer,
          prefillPolicy: "lock-when-available",
          required: true,
          tag: pointer,
        },
      },
    ],
    selection: selectedControl("title", "text", "control-1"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.element("field-panel").hidden).toBe(false);
  expect(harness.element("field-selection-tag").textContent).toBe("title");
  expect(harness.requests[0]).toMatchObject({
    credentials: "omit",
    method: "GET",
    url: "https://api.example.test/api/admin/forms/public-id/field-rules",
  });

  const query = harness.element("field-schema-query");
  query.value = "address";
  query.dispatchEvent({ target: query, type: "input" });
  harness.element("field-schema-search").dispatchEvent({ type: "click" });
  await flushPlugin();

  expect(harness.requests[1]).toMatchObject({
    credentials: "omit",
    method: "GET",
    url: "https://api.example.test/api/admin/forms/public-id/schema?q=address",
  });
  expect(harness.requests[1].headers.get("x-editor-capability")).toBe(
    "field-capability-2"
  );
  const list = harness.element("field-schema-list");
  expect(list.children).toHaveLength(1);
  const firstRow = list.children[0];
  const firstActions = firstRow.children[2];
  firstActions.children[0].dispatchEvent({ type: "click" });
  await flushPlugin();
  expect(copiedPointer).toBe(pointer);

  firstRow.children[0].dispatchEvent({ type: "click" });
  harness.element("field-schema-next").dispatchEvent({ type: "click" });
  await flushPlugin();
  expect(harness.requests[2]).toMatchObject({
    credentials: "omit",
    method: "GET",
    url: "https://api.example.test/api/admin/forms/public-id/schema?q=address&cursor=schema-cursor-1",
  });
  expect(list.children).toHaveLength(2);

  harness.element("field-apply-pointer").dispatchEvent({ type: "click" });
  await flushPlugin();
  const required = harness.element("field-required");
  required.checked = true;
  required.dispatchEvent({ target: required, type: "change" });
  const policy = harness.element("field-prefill-policy");
  policy.value = "lock-when-available";
  policy.dispatchEvent({ target: policy, type: "change" });
  harness.element("field-policy-form").dispatchEvent({ type: "submit" });
  await flushPlugin();

  expect(harness.requests[3]).toMatchObject({
    credentials: "omit",
    method: "PATCH",
    url: "https://api.example.test/api/admin/forms/public-id/field-rules",
  });
  expect(JSON.parse(harness.requests[3].body)).toEqual({
    documentKey: "document-key",
    prefillPointer: pointer,
    prefillPolicy: "lock-when-available",
    previousTag: null,
    required: true,
    tag: pointer,
  });
});
test("keeps native picture fields embedded and outside prefill configuration", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["picture-field-capability"],
    responses: [
      {
        rules: [
          {
            prefillPointer: "/photo",
            prefillPolicy: "lock-when-available",
            required: true,
            tag: "photo",
          },
        ],
      },
    ],
    selection: selectedControl("photo", "picture", "picture-control"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.element("field-picture-help").hidden).toBe(false);
  expect(harness.element("field-prefill-policy").disabled).toBe(true);
  expect(harness.element("field-schema-query").disabled).toBe(true);
  expect(harness.window.FormBridge.getPanelState()).toMatchObject({
    currentPointer: null,
    prefillPolicy: "editable",
    required: true,
    selection: { controlType: "picture", tag: "photo" },
  });
});

test("loads persisted field rules into the selected panel", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["field-capability"],
    responses: [
      {
        rules: [
          {
            prefillPointer: "/account/address/city",
            prefillPolicy: "lock-when-available",
            required: true,
            tag: "title",
          },
        ],
      },
    ],
    selection: selectedControl("title", "text", "control-1"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.element("field-required").checked).toBe(true);
  expect(harness.element("field-prefill-policy").value).toBe(
    "lock-when-available"
  );
  expect(harness.window.FormBridge.getPanelState()).toMatchObject({
    currentPointer: "/account/address/city",
    selection: { tag: "title" },
  });
});

test("keeps field configuration available only in template-edit mode", async () => {
  const templateHarness = createHarness({ action: "template-edit" });
  acknowledgeBridge(templateHarness);
  await flushPlugin();
  expect(templateHarness.element("field-panel").hidden).toBe(false);
  expect(typeof templateHarness.window.FormBridge.saveFieldRule).toBe(
    "function"
  );
  expect(
    templateHarness.messages.find(
      ({ message }) => message.type === "field-selection"
    )?.message
  ).toMatchObject({
    controlType: "unsupported",
    selected: false,
    tag: null,
    type: "field-selection",
  });

  const fillHarness = createHarness({ action: "fill" });
  expect(fillHarness.element("field-panel").hidden).toBe(true);
  expect(fillHarness.window.FormBridge.applySchemaPointer).toBeUndefined();
  expect(fillHarness.window.FormBridge.getFieldRules).toBeUndefined();
  expect(fillHarness.window.FormBridge.saveFieldRule).toBeUndefined();
  for (const action of ["configure-fields", "publish", "save-template"]) {
    await expect(
      fillHarness.window.FormBridge.runAction(action)
    ).resolves.toMatchObject({ ok: false });
  }
  expect(fillHarness.requests).toHaveLength(0);
  expect(
    fillHarness.messages.filter(
      ({ message }) => message.type === "capability-request"
    )
  ).toHaveLength(0);
});

test("reports field-rule API failures without exposing a value payload", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["field-capability"],
    responses: [{ body: { error: "unavailable" }, httpStatus: 503 }],
    selection: selectedControl("title", "text", "control-1"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.element("field-panel-status").textContent).toBe(
    "ไม่สามารถโหลดนโยบายฟิลด์ได้"
  );
  expect(harness.requests[0]?.body).toBeUndefined();
  expect(
    harness.messages
      .filter(({ message }) => message.type === "field-selection")
      .every(({ message }) => !Object.hasOwn(message, "value"))
  ).toBe(true);
});

test("retries persisted field rules when bridge acknowledgement is delayed", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["field-capability"],
    responses: [{ rules: [] }],
    selection: selectedControl("title", "text", "control-1"),
  });
  await flushPlugin();
  expect(harness.requests).toHaveLength(0);

  acknowledgeBridge(harness);
  await flushPlugin();
  expect(harness.requests).toHaveLength(1);
  expect(harness.requests[0]?.url).toContain("/field-rules");
});
test("maps real ONLYOFFICE form-type snapshots to supported controls", async () => {
  const cases = [
    ["textForm", "text"],
    ["comboBoxForm", "combo"],
    ["dropDownForm", "dropdown"],
    ["checkBoxForm", "checkbox"],
    ["pictureForm", "picture"],
    ["dateForm", "date"],
    ["radioForm", "unsupported"],
  ];

  for (const [formType, expectedType] of cases) {
    const harness = createHarness({
      action: "template-edit",
      capabilityResponses: ["field-capability"],
      responses: [{ rules: [] }],
      selection: {
        currentControl: "control-1",
        formType,
        properties: {
          Id: "control-1",
          InternalId: "control-1",
          Tag: "title",
        },
      },
    });
    acknowledgeBridge(harness);
    await flushPlugin();

    const selectionMessage = harness.messages
      .filter(({ message }) => message.type === "field-selection")
      .at(-1)?.message;
    expect(selectionMessage).toMatchObject({
      controlType: expectedType,
      selected: true,
      tag: "title",
      type: "field-selection",
    });
  }
});
