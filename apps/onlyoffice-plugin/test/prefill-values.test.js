// oxlint-disable prefer-await-to-callbacks node/callback-return unicorn/prefer-dom-node-remove prefer-destructuring unicorn/prefer-response-static-json no-await-in-loop unicorn/consistent-function-scoping no-plusplus unicorn/prefer-array-find -- VM harness intentionally mirrors browser callbacks and DOM shims.
import { expect, test } from "bun:test";

import {
  acknowledgeBridge,
  completedOperation,
  createHarness,
  flushPlugin,
} from "./harness.js";

test("extracts scalar form values with plugin contract semantics", async () => {
  const control = ({
    checkbox = false,
    checked = false,
    classType = "inlineLvlSdt",
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
    GetClassType: () => classType,
    GetContent: () => ({ GetText: () => text }),
    GetDropdownList: () => ({ GetAllItems: () => items }),
    GetFormType: () => formType,
    GetRange: () => ({ GetText: () => text }),
    GetTag: () => tag,
    ...(classType === "inlineLvlSdt"
      ? {
          GetDate: () => dateValue,
          IsCheckBox: () => checkbox,
          IsCheckBoxChecked: () => checked,
          IsComboBox: () => combo,
          IsDatePicker: () => date,
          IsDropDownList: () => dropdown,
          IsPicture: () => picture,
        }
      : {}),
  });

  const harness = createHarness({
    action: "draft",
    capabilityResponses: ["save-draft-capability"],
    controls: [
      control({ tag: "notes", text: "line one\nline two" }),
      control({
        classType: "blockLvlSdt",
        tag: "accept_terms",
        text: "✓",
      }),
      control({
        classType: "blockLvlSdt",
        tag: "start_date",
        text: "15/09/2026",
      }),
      control({
        classType: "blockLvlSdt",
        tag: "department",
        text: "Engineering",
      }),
      control({
        classType: "blockLvlSdt",
        tag: "custom",
        text: "Custom value",
      }),
      control({
        checkbox: true,
        checked: true,
        tag: "inline_checkbox",
        text: "checked",
      }),
      control({
        date: true,
        dateValue: new Date(2026, 8, 15),
        tag: "inline_date",
        text: "date",
      }),
      control({
        dropdown: true,
        items: [
          {
            GetText: () => "Engineering",
            GetValue: () => "engineering",
          },
        ],
        tag: "inline_department",
        text: "Engineering",
      }),
      control({
        dropdown: true,
        items: [
          {
            GetText: () => "Empty option label",
            GetValue: () => "",
          },
          {
            GetText: () => "Choice B",
            GetValue: () => "Empty option label",
          },
        ],
        tag: "selected_empty",
        text: "Empty option label",
      }),
      control({
        dropdown: true,
        items: [
          {
            GetText: () => "Empty option label",
            GetValue: () => "",
          },
          {
            GetText: () => "Choice B",
            GetValue: () => "Empty option label",
          },
        ],
        tag: "selected_value_collision",
        text: "Choice B",
      }),
      control({
        dropdown: true,
        items: [
          {
            GetText: () => "Empty option label",
            GetValue: () => "",
          },
        ],
        tag: "blank_empty",
        text: "",
      }),
      control({
        combo: true,
        items: [
          {
            GetText: () => "Empty combo label",
            GetValue: () => "",
          },
          {
            GetText: () => "Choice B",
            GetValue: () => "Empty combo label",
          },
        ],
        tag: "selected_empty_combo",
        text: "Empty combo label",
      }),
      control({
        combo: true,
        items: [
          {
            GetText: () => "Empty combo label",
            GetValue: () => "",
          },
          {
            GetText: () => "Choice B",
            GetValue: () => "Empty combo label",
          },
        ],
        tag: "selected_combo_value_collision",
        text: "Choice B",
      }),
      control({
        combo: true,
        items: [
          {
            GetText: () => "Empty combo label",
            GetValue: () => "",
          },
        ],
        tag: "blank_empty_combo",
        text: "",
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
    accept_terms: "✓",
    blank_empty: "",
    blank_empty_combo: "",
    custom: "Custom value",
    department: "Engineering",
    inline_checkbox: true,
    inline_date: "2026-09-15",
    inline_department: "Engineering",
    notes: "line one\nline two",
    selected_combo_value_collision: "Choice B",
    selected_empty: "Empty option label",
    selected_empty_combo: "Empty combo label",
    selected_value_collision: "Choice B",
    start_date: "15/09/2026",
  });
  expect(requestBody.canonicalDateFields).toEqual(["inline_date"]);
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
test("maps namespaced field tags during prefill and save", async () => {
  let lock;
  let text = "";
  const control = {
    AddText(value) {
      text += value;
    },
    GetClassType: () => "inlineLvlSdt",
    GetRange: () => ({ GetText: () => text }),
    GetTag: () => "metadata",
    IsCheckBox: () => false,
    IsComboBox: () => false,
    IsDatePicker: () => false,
    IsDropDownList: () => false,
    RemoveAllElements() {
      text = "";
    },
    SetLock(value) {
      lock = value;
    },
  };
  const harness = createHarness({
    action: "draft",
    capabilityResponses: ["save-draft-capability"],
    controls: [control],
    responses: [
      { operationCapability: "save-operation-capability", operationId: "save" },
      completedOperation({ saved: true }),
    ],
    tagAliases: { metadata: "full_name" },
  });
  acknowledgeBridge(harness);

  await expect(
    harness.window.FormBridge.applyPrefill({
      editableFields: { full_name: false },
      values: { full_name: "Trusted value" },
    })
  ).resolves.toEqual({
    applied: ["full_name"],
    failed: [],
    skipped: [],
  });
  await expect(
    harness.window.FormBridge.runAction("save-draft")
  ).resolves.toMatchObject({ ok: true });

  expect(JSON.parse(harness.requests[0].body).data).toEqual({
    full_name: "Trusted value",
  });
  expect({ lock, text }).toEqual({
    lock: "sdtContentLocked",
    text: "Trusted value",
  });
});
const createMutableControl = ({ kind, tag, items = [] }) => {
  let checked = false;
  let dateValue = null;
  let lock;
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
    SetLock: (value) => {
      lock = value;
    },
    read: () => ({ checked, dateValue, textValue }),
    readLock: () => lock,
  };
};

test("applies a saved scalar response and rejects API failures", async () => {
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
    department: "Engineering",
    notes: "line one\nline two",
    start_date: "2026-09-15",
  });

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
});

test("runs bundled document commands after serialization in an isolated Office context", async () => {
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
    createMutableControl({ kind: "text", tag: "missing" }),
  ];
  const harness = createHarness({
    action: "draft",
    capabilityResponses: ["save-draft-capability"],
    controls,
    responses: [{ ok: true }],
    serializeCommands: true,
  });
  acknowledgeBridge(harness);

  await expect(
    harness.window.FormBridge.applyPrefill({
      editableFields: {
        accept_terms: true,
        custom: true,
        department: true,
        missing: false,
        notes: false,
        start_date: false,
      },
      values: {
        accept_terms: true,
        custom: "Custom value",
        department: "engineering",
        notes: "line one\nline two",
        start_date: "2026-09-15",
      },
    })
  ).resolves.toEqual({
    applied: ["notes", "accept_terms", "start_date", "department", "custom"],
    failed: [],
    skipped: ["missing"],
  });
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
    { checked: false, dateValue: null, textValue: "" },
  ]);
  expect(controls.map((control) => control.readLock())).toEqual([
    "sdtContentLocked",
    "unlocked",
    "sdtContentLocked",
    "unlocked",
    "unlocked",
    undefined,
  ]);

  harness.emitEditorEvent("onChangeContentControl");
  await expect(
    harness.window.FormBridge.runAction("save-draft")
  ).resolves.toMatchObject({ ok: true });
  const { promise: extracted, resolve, reject } = Promise.withResolvers();
  harness.window.FormBridge.extractFormData((error, data) => {
    if (error) {
      reject(error);
      return;
    }
    resolve(data);
  });
  await expect(extracted).resolves.toEqual({
    canonicalDateFields: ["start_date"],
    data: {
      accept_terms: true,
      custom: "Custom value",
      department: "Engineering",
      missing: "",
      notes: "line one\nline two",
      start_date: "2026-09-15",
    },
  });
  expect(harness.saveCalls).toEqual(["save"]);

  const selection = {
    currentControl: "control-1",
    formType: "textForm",
    properties: {
      Id: "control-1",
      InternalId: "control-1",
      Tag: "title",
    },
  };
  const templateHarness = createHarness({
    action: "template-edit",
    capabilityResponses: ["field-capability"],
    responses: [{ rules: [] }],
    selection,
    serializeCommands: true,
  });
  acknowledgeBridge(templateHarness);
  await flushPlugin();
  expect(
    templateHarness.messages
      .filter(({ message }) => message.type === "field-selection")
      .at(-1).message
  ).toMatchObject({
    controlType: "text",
    selected: true,
    tag: "title",
  });
  const pointer = "/account/address/city";
  await expect(
    templateHarness.window.FormBridge.applySchemaPointer(pointer)
  ).resolves.toEqual({ ok: true, tag: pointer });
  expect(selection.properties.Tag).toBe(pointer);
  await expect(
    templateHarness.window.FormBridge.refreshSelection()
  ).resolves.toEqual({
    controlType: "text",
    ok: true,
    selected: true,
    tag: pointer,
  });
});
