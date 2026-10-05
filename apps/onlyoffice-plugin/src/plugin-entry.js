// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import { createApi } from "./api.js";
import { createBridge } from "./bridge.js";
import { ACTIONS, BUTTON_IDS } from "./constants.js";
import { createFieldPanel } from "./field-panel.js";
import { createFields } from "./fields.js";
import { createOffice } from "./office.js";
import { createOperations } from "./operations.js";
import {
  createPrefill,
  normalizePrefill,
  hasPrefillValues,
} from "./prefill.js";
import { isRecord, firstString, actionLabel, errorMessage } from "./values.js";

let runtimeOptions = {};
let pluginInitialized = false;
let initializationPending = false;
let initializationStarted = false;
let pendingContentChange = false;
let office;
let api;
let prefill;
let bridge;
let panel;
let fields;
let operations;

function getStatusElement() {
  const existing = document.querySelector("#form-bridge-status");

  if (existing) {
    return existing;
  }

  if (!document.body) {
    return null;
  }

  const element = document.createElement("div");
  element.id = "form-bridge-status";
  element.setAttribute("role", "status");
  element.setAttribute("aria-live", "polite");
  document.body.append(element);

  return element;
}

function setStatus(message, state) {
  const element = getStatusElement();

  if (element) {
    element.textContent = message;
    element.dataset.state = state || "info";
  }
}

function normalizeRuntimeOptions() {
  let options = window.Asc.plugin.info?.options || {};

  if (typeof options === "string") {
    try {
      options = JSON.parse(options);
    } catch {
      options = {};
    }
  }
  if (!isRecord(options)) {
    options = {};
  }
  return {
    action: firstString(options.action)?.toLowerCase(),
    apiBase: firstString(options.apiBase)?.replace(/\/+$/, ""),
    bridgeId: firstString(options.bridgeId),
    documentKey: firstString(options.documentKey),
    formId: firstString(options.formId),
    operationCapability: firstString(options.operationCapability),
    operationId: firstString(options.operationId),
    parentOrigin: firstString(options.parentOrigin),
    prefill: normalizePrefill(options),
    publicId: firstString(options.publicId),
    responseId: firstString(options.responseId),
    tagAliases: isRecord(options.tagAliases) ? options.tagAliases : {},
    targetId: firstString(options.targetId),
  };
}

function toolbarActionsForMode(mode) {
  switch (mode) {
    case ACTIONS.TEMPLATE_EDIT: {
      return [ACTIONS.SAVE_TEMPLATE, ACTIONS.PUBLISH];
    }
    case ACTIONS.FILL:
    case ACTIONS.DRAFT: {
      return [ACTIONS.SAVE_DRAFT, ACTIONS.SUBMIT];
    }
    case ACTIONS.CORRECTION: {
      return [];
    }
    case ACTIONS.SUBMIT: {
      return [ACTIONS.SUBMIT];
    }
    default: {
      return [];
    }
  }
}

function buttonIdForAction(action) {
  switch (action) {
    case ACTIONS.SAVE_TEMPLATE: {
      return BUTTON_IDS.SAVE_TEMPLATE;
    }
    case ACTIONS.PUBLISH: {
      return BUTTON_IDS.PUBLISH;
    }
    case ACTIONS.SAVE_DRAFT: {
      return BUTTON_IDS.SAVE_DRAFT;
    }
    case ACTIONS.SUBMIT: {
      return BUTTON_IDS.SUBMIT;
    }
    default: {
      return;
    }
  }
}

function addToolbarMenuItems(actions) {
  if (!actions.length) {
    return;
  }

  const items = actions.map((action) => ({
    enableToggle: false,
    hint: actionLabel(action),
    id: buttonIdForAction(action),
    lockInViewMode: false,
    separator: false,
    text: actionLabel(action),
    type: "button",
  }));

  try {
    window.Asc.plugin.executeMethod("AddToolbarMenuItem", [
      {
        guid: window.Asc.plugin.info.guid,
        tabs: [
          {
            id: "form_bridge",
            items,
            text: "Form",
          },
        ],
      },
    ]);
  } catch (error) {
    setStatus(`Could not add form toolbar: ${errorMessage(error)}`, "error");
  }
}

function attachToolbarHandlers(actions) {
  const currentOperations = operations;
  for (const action of actions) {
    const buttonId = buttonIdForAction(action);

    if (!buttonId) {
      continue;
    }

    try {
      window.Asc.plugin.attachToolbarMenuClickEvent(buttonId, () => {
        void currentOperations.runAction(action);
      });
    } catch (error) {
      setStatus(
        `Could not attach ${actionLabel(action)}: ${errorMessage(error)}`,
        "error"
      );
    }
  }
}

function startInitializationTasks() {
  if (initializationStarted) {
    return;
  }
  initializationStarted = true;
  const tasks = [];
  if (prefill.shouldApplyRuntimePrefill()) {
    window.setTimeout(() => {
      prefill
        .ensurePrefill()
        .then((result) => {
          setStatus(
            `เติมข้อมูลล่วงหน้าแล้ว (${result.applied?.length ?? 0} ช่อง, ข้าม ${
              result.skipped?.length ?? 0
            } ช่อง)`,
            "success"
          );
        })
        .catch((error) => {
          setStatus(`เติมข้อมูลล่วงหน้าไม่สำเร็จ: ${errorMessage(error)}`, "error");
        });
    }, 5000);
  }
  if (
    (runtimeOptions.action === ACTIONS.FILL ||
      runtimeOptions.action === ACTIONS.DRAFT ||
      runtimeOptions.action === ACTIONS.CORRECTION ||
      runtimeOptions.action === ACTIONS.SUBMIT) &&
    !hasPrefillValues(runtimeOptions.prefill)
  ) {
    window.setTimeout(() => {
      office.restrictEditorToForms().catch((error) => {
        setStatus(`จำกัดการแก้ไขเอกสารไม่สำเร็จ: ${errorMessage(error)}`, "error");
      });
    }, 3000);
  }

  if (runtimeOptions.operationId) {
    tasks.push(
      operations
        .pollOperation(
          runtimeOptions.operationId,
          "Existing operation",
          runtimeOptions.operationCapability
        )
        .then((result) => {
          setStatus("การดำเนินการเดิมเสร็จแล้ว", "success");
          return result;
        })
        .catch((error) => {
          setStatus(`การดำเนินการเดิมไม่สำเร็จ: ${errorMessage(error)}`, "error");
          return null;
        })
    );
  }

  if (!tasks.length) {
    return;
  }

  initializationPending = true;
  Promise.all(tasks).finally(() => {
    initializationPending = false;
  });
}

function startInitializationWhenReady() {
  const plugin = window.Asc?.plugin;
  if (plugin && typeof plugin.attachEditorEvent === "function") {
    plugin.attachEditorEvent("onDocumentContentReady", () => {
      if (runtimeOptions.action === ACTIONS.CORRECTION) {
        office.restrictEditorToForms().catch((error) => {
          setStatus(
            `จำกัดการแก้ไขเอกสารไม่สำเร็จ: ${errorMessage(error)}`,
            "error"
          );
        });
      }
      startInitializationTasks();
    });
    plugin.attachEditorEvent("onDocumentContentChanged", () => {
      bridge.setDirtyState(true);
    });
    plugin.attachEditorEvent("onChangeContentControl", () => {
      bridge.setDirtyState(true);
    });
  }

  window.setTimeout(startInitializationTasks, 3000);
}

function exposeFormBridge() {
  const formBridge = Object.assign(window.FormBridge || {}, {
    applyPrefill: prefill.applyPrefill,
    extractFormData: office.extractFormData,
    getRuntimeOptions: () => runtimeOptions,
    pollOperation: operations.pollOperation,
    runAction: operations.runAction,
    submitForm: operations.submitForm,
  });
  if (runtimeOptions.action === ACTIONS.TEMPLATE_EDIT) {
    Object.assign(formBridge, {
      applySchemaPointer: fields.applySchemaPointer,
      copySchemaPointer: fields.copySchemaPointer,
      getFieldRules: fields.loadFieldRules,
      getPanelState: fields.getPanelState,
      loadSchemaPage: fields.loadSchemaPage,
      refreshSelection: fields.refreshSelection,
      saveFieldRule: fields.saveFieldRule,
      selectSchemaPointer: fields.selectSchemaPointer,
      setSchemaQuery: fields.setSchemaQuery,
    });
  } else {
    delete formBridge.applySchemaPointer;
    delete formBridge.copySchemaPointer;
    delete formBridge.getFieldRules;
    delete formBridge.getPanelState;
    delete formBridge.loadSchemaPage;
    delete formBridge.refreshSelection;
    delete formBridge.saveFieldRule;
    delete formBridge.selectSchemaPointer;
    delete formBridge.setSchemaQuery;
  }
  window.FormBridge = formBridge;
  window.submitForm = operations.submitForm;
}

function initializePlugin() {
  if (pluginInitialized) {
    return;
  }

  pluginInitialized = true;
  runtimeOptions = normalizeRuntimeOptions();
  office = createOffice(runtimeOptions);
  api = createApi(runtimeOptions);
  prefill = createPrefill(runtimeOptions, office);
  bridge = createBridge(runtimeOptions, office, {
    getFieldSelection: () => fields.state.selection,
    loadFieldRules: () => fields.loadFieldRules(),
    runAction: (...args) => operations.runAction(...args),
    setPanelStatus: (...args) => panel.setPanelStatus(...args),
    setStatus,
  });
  panel = createFieldPanel(() => fields);
  fields = createFields(runtimeOptions, office, bridge, panel, api);
  operations = createOperations(runtimeOptions, office, prefill, bridge, api, {
    isInitializationPending: () => initializationPending,
    setStatus,
  });
  if (pendingContentChange) {
    bridge.setDirtyState(true);
    pendingContentChange = false;
  }
  exposeFormBridge();
  bridge.startBridge();

  if (runtimeOptions.action === ACTIONS.TEMPLATE_EDIT) {
    panel.setupFieldPanel();
  } else {
    panel.hideFieldPanel();
  }

  const actions = toolbarActionsForMode(runtimeOptions.action);
  if (actions.length) {
    addToolbarMenuItems(actions);
    attachToolbarHandlers(actions);
  } else if (runtimeOptions.action === ACTIONS.CORRECTION) {
    setStatus("พร้อมแก้ไขเฉพาะช่องกรอก", "success");
  } else {
    setStatus("ไม่พบการทำงานของฟอร์มที่รองรับ", "error");
  }
  if (runtimeOptions.action === ACTIONS.CORRECTION) {
    office.restrictEditorToForms().catch((error) => {
      setStatus(`จำกัดการแก้ไขเอกสารไม่สำเร็จ: ${errorMessage(error)}`, "error");
    });
  }

  startInitializationWhenReady();
}

/**
 * ONLYOFFICE plugin entry point.
 *
 * The host may install plugin options immediately before invoking init, so
 * wait for the info object without delaying registration of the init hook.
 */
window.Asc = window.Asc || {};
window.Asc.plugin = window.Asc.plugin || {};
window.Asc.plugin.event_onChangeContentControl = () => {
  if (bridge) {
    bridge.setDirtyState(true);
  } else {
    pendingContentChange = true;
  }
};
window.Asc.plugin.init = function init() {
  const start = () => {
    const options = window.Asc.plugin.info?.options;
    if (
      !options ||
      (typeof options === "object" && Object.keys(options).length === 0)
    ) {
      window.setTimeout(start, 50);
      return;
    }
    initializePlugin();
  };
  start();
};
