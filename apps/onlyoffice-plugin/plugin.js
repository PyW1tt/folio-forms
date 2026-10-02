// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
const ACTIONS = Object.freeze({
  CONFIGURE_FIELDS: "configure-fields",
  CORRECTION: "correction",
  DRAFT: "draft",
  FILL: "fill",
  PUBLISH: "publish",
  SAVE_CORRECTION: "save-correction",
  SAVE_DRAFT: "save-draft",
  SAVE_TEMPLATE: "save-template",
  SUBMIT: "submit",
  TEMPLATE_EDIT: "template-edit",
});
const BUTTON_IDS = Object.freeze({
  PUBLISH: "publish-form",
  SAVE_DRAFT: "save-draft",
  SAVE_TEMPLATE: "save-template",
  SUBMIT: "submit-form",
});

const API_ROUTES = Object.freeze({
  ADMIN_FORMS: "/api/admin/forms",
  ADMIN_RESULTS: "/api/admin/results",
  FORMS: "/api/forms",
  OPERATIONS: "/api/operations",
});

const CAPABILITY_ACTIONS = Object.freeze([
  ACTIONS.SAVE_TEMPLATE,
  ACTIONS.PUBLISH,
  ACTIONS.SAVE_DRAFT,
  ACTIONS.SAVE_CORRECTION,
  ACTIONS.SUBMIT,
  ACTIONS.CONFIGURE_FIELDS,
]);

const FIELD_CONTROL_TYPES = Object.freeze([
  "text",
  "checkbox",
  "date",
  "dropdown",
  "combo",
  "picture",
  "unsupported",
]);
const FIELD_RULE_POLICIES = Object.freeze(["editable", "lock-when-available"]);
const PANEL_IDS = Object.freeze({
  applyPointer: "field-apply-pointer",
  copyPrefix: "field-copy-",
  list: "field-schema-list",
  nextPage: "field-schema-next",
  panel: "field-panel",
  policyForm: "field-policy-form",
  pictureHelp: "field-picture-help",
  policySelect: "field-prefill-policy",
  query: "field-schema-query",
  required: "field-required",
  save: "field-save",
  search: "field-schema-search",
  selectionTag: "field-selection-tag",
  selectionType: "field-selection-type",
  status: "field-panel-status",
});

const BRIDGE_MESSAGE_SOURCE = "form-bridge";
const PARENT_MESSAGE_SOURCE = "folio-parent";
const BRIDGE_READY_TYPE = "bridge-ready";
const BRIDGE_ACK_TYPE = "bridge-ack";
const CAPABILITY_REQUEST_TYPE = "capability-request";
const CAPABILITY_RESPONSE_TYPE = "capability-response";
const DIRTY_STATE_TYPE = "dirty-state";
const RUN_ACTION_TYPE = "run-action";
const CLEAR_DIRTY_TYPE = "clear-dirty";
const OPERATION_MESSAGE_TYPE = "operation";
const EDITOR_SAVE_REQUEST_TYPE = "editor-save-request";
const EDITOR_SAVE_RESPONSE_TYPE = "editor-save-response";

const CAPABILITY_REQUEST_TIMEOUT_MS = 5000;
const OPERATION_POLL_INTERVAL_MS = 1000;
const MAX_OPERATION_POLLS = 300;
const PREFILL_MAX_ATTEMPTS = 120;
const MAX_SCHEMA_QUERY_LENGTH = 128;
const MAX_SCHEMA_PAGE_ITEMS = 200;
const MAX_SCHEMA_PAGES = 50;

let runtimeOptions = {};
let actionInFlight = false;
let documentDirty = false;
let panelElements = {};
let panelEventsAttached = false;
let panelSelectionEventAttached = false;
let selectionSequence = 0;
let schemaRequestSequence = 0;
const fieldPanelState = {
  currentPointer: null,
  prefillPolicy: "editable",
  required: false,
  rules: [],
  schemaCursor: null,
  schemaItems: [],
  schemaPageCount: 0,
  schemaQuery: "",
  selection: null,
  selectedPointer: null,
  saving: false,
};
let panelSelectionRequestSequence = 0;
let pendingFieldSelection;
let bridgeAcknowledged = false;
let bridgeMessageListenerAttached = false;
let editorSaveSupported = false;
let editorSaveRequestSequence = 0;
let pendingEditorSave;
let bridgeReadySent = false;
let capabilityRequestSequence = 0;
const pendingCapabilityRequests = new Map();
let initializationPending = false;
let initializationStarted = false;
let prefillApplied = false;
let prefillPromise;
let pluginInitialized = false;

function requireOption(value, name) {
  if (!value) {
    throw new Error(`Missing editor option: ${name}`);
  }

  return value;
}

/**
 * Runs inside ONLYOFFICE document context.
 *
 * IMPORTANT:
 * This function is passed into Asc.plugin.callCommand(), so it can access
 * Api.GetDocument() and Office API objects. Do not move dependencies into the
 * outer plugin scope: ONLYOFFICE serializes this function before execution.
 */
function extractFormDataCommand() {
  const scope = typeof Asc !== "undefined" && Asc.scope ? Asc.scope : {};
  const tagAliases = scope.formBridgeTagAliases || {};
  const doc = Api.GetDocument();
  const controls = doc.GetAllContentControls();
  const data = {};
  const canonicalDateFields = [];
  function isPictureControlInsideCommand(control) {
    if (typeof control.IsPicture === "function") {
      try {
        if (control.IsPicture()) {
          return true;
        }
      } catch {
        // Fall through to the normalized form type.
      }
    }

    if (typeof control.GetFormType === "function") {
      try {
        const formType = control.GetFormType();
        if (typeof formType === "string") {
          const normalized = formType.toLowerCase().replaceAll(/[\s_-]+/g, "");
          return (
            normalized === "picture" ||
            normalized === "pictureform" ||
            normalized === "picturecontentcontrol" ||
            normalized === "image"
          );
        }
      } catch {
        // Treat controls with an unavailable form type as scalar candidates.
      }
    }

    return false;
  }

  function getInlineTextInsideCommand(control) {
    return control
      .GetRange()
      .GetText({
        ParaSeparator: "\n",
        NewLineSeparator: "\n",
      })
      .trim();
  }

  function formatDateInsideCommand(date) {
    if (!date) {
      return null;
    }

    if (typeof date === "string") {
      return date.slice(0, 10);
    }

    try {
      if (typeof date.getFullYear === "function") {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, "0");
        const day = String(date.getDate()).padStart(2, "0");

        return `${year}-${month}-${day}`;
      }

      if (typeof date.toISOString === "function") {
        return date.toISOString().slice(0, 10);
      }
    } catch {
      return null;
    }

    return null;
  }

  for (const control of controls) {
    const receivedTag = control.GetTag();
    const rawTag = typeof receivedTag === "string" ? receivedTag.trim() : "";
    const tag =
      typeof tagAliases[rawTag] === "string" ? tagAliases[rawTag] : rawTag;

    if (!tag) {
      continue;
    }

    if (isPictureControlInsideCommand(control)) {
      continue;
    }

    const classType = control.GetClassType();
    if (classType !== "inlineLvlSdt" && classType !== "blockLvlSdt") {
      continue;
    }

    /**
     * Checkbox
     *
     * Return a real boolean instead of "☒" / "☐".
     */
    if (typeof control.IsCheckBox === "function" && control.IsCheckBox()) {
      data[tag] = Boolean(control.IsCheckBoxChecked());
      continue;
    }

    /**
     * Date picker
     *
     * Normalize to YYYY-MM-DD without changing the calendar date for local
     * Date objects returned by the Office API.
     */
    if (typeof control.IsDatePicker === "function" && control.IsDatePicker()) {
      const dateValue = formatDateInsideCommand(control.GetDate());
      data[tag] = dateValue;
      if (
        typeof dateValue === "string" &&
        /^\d{4}-\d{2}-\d{2}$/u.test(dateValue)
      ) {
        const parsedDate = new Date(`${dateValue}T00:00:00.000Z`);
        if (
          !Number.isNaN(parsedDate.getTime()) &&
          parsedDate.toISOString().slice(0, 10) === dateValue
        ) {
          canonicalDateFields.push(tag);
        }
      }
      continue;
    }

    /**
     * Dropdown / Combo box
     *
     * Send visible text so the server can resolve the published option label.
     */
    if (
      (typeof control.IsDropDownList === "function" &&
        control.IsDropDownList()) ||
      (typeof control.IsComboBox === "function" && control.IsComboBox())
    ) {
      data[tag] = getInlineTextInsideCommand(control);
      continue;
    }
    if (classType === "blockLvlSdt") {
      data[tag] = control
        .GetContent()
        .GetText({
          ParaSeparator: "\n",
          NewLineSeparator: "\n",
        })
        .trim();

      continue;
    }

    /**
     * Plain Text and other inline controls.
     */
    data[tag] = getInlineTextInsideCommand(control);
  }

  /**
   * callCommand() transports primitive/string data cleanly back to the
   * plugin iframe.
   */
  return JSON.stringify({ canonicalDateFields, data });
}

/**
 * Runs inside ONLYOFFICE document context and applies server-provided
 * prefill values. Asc.scope is the supported bridge for passing data into a
 * callCommand function.
 */
function applyPrefillCommand() {
  const scope =
    typeof Asc !== "undefined" && Asc.scope
      ? Asc.scope
      : { formBridgePrefill: null };
  const tagAliases = scope.formBridgeTagAliases || {};
  function normalizedFieldTag(control) {
    const rawTag = control.GetTag();
    const tag = typeof rawTag === "string" ? rawTag.trim() : "";
    return typeof tagAliases[tag] === "string" ? tagAliases[tag] : tag;
  }
  const payload = scope.formBridgePrefill || {};
  const values = payload.values || {};
  const policies = payload.policies || {};
  const doc = Api.GetDocument();
  const controls = doc.GetAllContentControls();
  const applied = [];
  const skipped = [];
  const failed = [];
  const hasOwn = Object.prototype.hasOwnProperty;
  function isPictureControl(control) {
    if (typeof control.IsPicture === "function") {
      try {
        if (control.IsPicture()) {
          return true;
        }
      } catch {
        // Fall through to the normalized form type.
      }
    }

    if (typeof control.GetFormType === "function") {
      try {
        const formType = control.GetFormType();
        if (typeof formType === "string") {
          const normalized = formType.toLowerCase().replaceAll(/[\s_-]+/g, "");
          return (
            normalized === "picture" ||
            normalized === "pictureform" ||
            normalized === "picturecontentcontrol" ||
            normalized === "image"
          );
        }
      } catch {
        // Treat controls with an unavailable form type as scalar candidates.
      }
    }

    return false;
  }

  function setControlText(control, text) {
    const classType =
      typeof control.GetClassType === "function" ? control.GetClassType() : "";

    if (
      classType === "inlineLvlSdt" &&
      typeof control.RemoveAllElements === "function" &&
      typeof control.AddText === "function"
    ) {
      control.RemoveAllElements();
      control.AddText(text);
      return true;
    }

    if (
      classType === "blockLvlSdt" &&
      typeof control.GetContent === "function"
    ) {
      const content = control.GetContent();
      if (content && typeof content.SetText === "function") {
        content.SetText(text);
        return true;
      }
    }

    if (typeof control.SetText === "function") {
      control.SetText(text);
      return true;
    }

    if (
      classType === "inlineLvlSdt" &&
      typeof control.GetRange === "function"
    ) {
      const range = control.GetRange();
      if (range && typeof range.SetText === "function") {
        range.SetText(text);
        return true;
      }
    }

    return false;
  }

  function setControlValue(control, value) {
    const classType =
      typeof control.GetClassType === "function" ? control.GetClassType() : "";
    if (classType !== "inlineLvlSdt" && classType !== "blockLvlSdt") {
      return false;
    }
    if (classType === "inlineLvlSdt" && typeof control.SetLock === "function") {
      control.SetLock("unlocked");
    }
    if (typeof control.IsCheckBox === "function" && control.IsCheckBox()) {
      const checked =
        value === true ||
        value === 1 ||
        value === "1" ||
        (typeof value === "string" && value.toLowerCase() === "true");

      if (typeof control.SetCheckBoxChecked === "function") {
        control.SetCheckBoxChecked(checked);
        return true;
      }

      return setControlText(control, checked ? "☒" : "☐");
    }

    if (typeof control.IsDatePicker === "function" && control.IsDatePicker()) {
      if (value === null || value === "") {
        return setControlText(control, "");
      }

      const dateText = String(value);
      const dateParts = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateText);

      if (dateParts && typeof control.SetDate === "function") {
        control.SetDate(
          new Date(
            Number(dateParts[1]),
            Number(dateParts[2]) - 1,
            Number(dateParts[3])
          )
        );
        return true;
      }

      return setControlText(control, dateText);
    }

    const isDropdown =
      typeof control.IsDropDownList === "function" && control.IsDropDownList();
    const isComboBox =
      typeof control.IsComboBox === "function" && control.IsComboBox();

    if (isDropdown || isComboBox) {
      const valueText = value === null ? "" : String(value);
      if (typeof control.GetDropdownList === "function") {
        const list = control.GetDropdownList();
        if (list && typeof list.GetAllItems === "function") {
          const items = list.GetAllItems();
          for (const item of items) {
            if (!item || String(item.GetValue()) !== valueText) {
              continue;
            }
            if (typeof item.Select === "function") {
              item.Select();
              return true;
            }
          }
        }
      }
      return isComboBox && setControlText(control, valueText);
    }

    return setControlText(control, value === null ? "" : String(value));
  }

  function setControlPolicy(control, policy) {
    if (typeof control.SetLock !== "function" || policy === undefined) {
      return;
    }

    let policyName = "";

    if (typeof policy === "string") {
      policyName = policy.toLowerCase();
    } else if (policy === true) {
      policyName = "locked";
    } else if (policy && typeof policy === "object") {
      if (policy.locked === true || policy.editable === false) {
        policyName = "locked";
      } else if (policy.editable === true || policy.locked === false) {
        policyName = "editable";
      } else if (typeof policy.mode === "string") {
        policyName = policy.mode.toLowerCase();
      }
    }

    if (
      policyName === "locked" ||
      policyName === "readonly" ||
      policyName === "read-only" ||
      policyName === "contentlocked" ||
      policyName === "sdtcontentlocked"
    ) {
      control.SetLock("sdtContentLocked");
    } else if (
      policyName === "editable" ||
      policyName === "unlocked" ||
      policyName === "write"
    ) {
      control.SetLock("unlocked");
    }
  }

  for (const control of controls) {
    const tag = normalizedFieldTag(control);

    if (!tag) {
      continue;
    }

    if (isPictureControl(control)) {
      skipped.push(tag);
      continue;
    }

    if (!hasOwn.call(values, tag)) {
      skipped.push(tag);
      continue;
    }

    try {
      if (!setControlValue(control, values[tag])) {
        throw new Error("The content control does not support text updates");
      }

      setControlPolicy(control, policies[tag]);
      applied.push(tag);
    } catch {
      failed.push(tag);
    }
  }

  return JSON.stringify({
    applied,
    failed,
    skipped,
  });
}

/**
 * Read the current content control in ONLYOFFICE document context.
 *
 * The returned identifier is kept inside the plugin only. It is never sent
 * across the form bridge.
 */
function getCurrentContentControlCommand() {
  const doc = Api.GetDocument();
  const control =
    typeof doc.GetCurrentContentControl === "function"
      ? doc.GetCurrentContentControl()
      : null;

  if (!control) {
    return JSON.stringify(null);
  }

  const properties =
    typeof doc.GetCurrentContentControlPr === "function"
      ? doc.GetCurrentContentControlPr("none")
      : {};
  const readString = (value) => {
    if (typeof value === "string" || typeof value === "number") {
      return String(value);
    }

    return "";
  };
  const identifier = readString(
    properties?.InternalId ??
      properties?.internalId ??
      properties?.Id ??
      properties?.id ??
      (typeof control.GetInternalId === "function"
        ? control.GetInternalId()
        : "")
  );
  let controlType = "unsupported";
  const formType =
    typeof control.GetFormType === "function"
      ? readString(control.GetFormType())
          .toLowerCase()
          .replaceAll(/[\s_-]+/g, "")
      : "";
  let knownFormType = true;

  switch (formType) {
    case "text":
    case "textform": {
      controlType = "text";
      break;
    }
    case "combobox":
    case "comboboxform": {
      controlType = "combo";
      break;
    }
    case "dropdown":
    case "dropdownform": {
      controlType = "dropdown";
      break;
    }
    case "checkbox":
    case "checkboxform": {
      controlType = "checkbox";
      break;
    }
    case "picture":
    case "pictureform": {
      controlType = "picture";
      break;
    }
    case "date":
    case "dateform": {
      controlType = "date";
      break;
    }
    case "radio":
    case "radioform":
    case "complex":
    case "complexform":
    case "signature":
    case "signatureform": {
      break;
    }
    default: {
      knownFormType = false;
    }
  }

  if (!knownFormType) {
    if (typeof control.IsCheckBox === "function" && control.IsCheckBox()) {
      controlType = "checkbox";
    } else if (
      typeof control.IsDatePicker === "function" &&
      control.IsDatePicker()
    ) {
      controlType = "date";
    } else if (
      typeof control.IsDropDownList === "function" &&
      control.IsDropDownList()
    ) {
      controlType = "dropdown";
    } else if (
      typeof control.IsComboBox === "function" &&
      control.IsComboBox()
    ) {
      controlType = "combo";
    } else if (typeof control.IsPicture === "function" && control.IsPicture()) {
      controlType = "picture";
    } else if (
      typeof control.GetClassType === "function" &&
      (control.GetClassType() === "inlineLvlSdt" ||
        control.GetClassType() === "blockLvlSdt")
    ) {
      controlType = "text";
    }
  }

  const tag =
    typeof control.GetTag === "function" ? readString(control.GetTag()) : "";

  return JSON.stringify({
    controlType,
    internalId: identifier || undefined,
    selected: true,
    tag: tag || null,
  });
}

/**
 * Runs inside ONLYOFFICE document context and applies one exact schema
 * pointer as the selected control tag.
 */
function setCurrentContentControlTagCommand() {
  const scope = typeof Asc !== "undefined" && Asc.scope ? Asc.scope : {};
  const tag =
    typeof scope.formBridgeSelectionTag === "string"
      ? scope.formBridgeSelectionTag
      : "";
  const targetId =
    typeof scope.formBridgeSelectionId === "string"
      ? scope.formBridgeSelectionId
      : "";

  if (!tag) {
    return JSON.stringify({ ok: false });
  }

  const doc = Api.GetDocument();
  const controls =
    typeof doc.GetAllContentControls === "function"
      ? doc.GetAllContentControls()
      : [];
  let selectedControl = null;

  for (const control of controls) {
    if (!control || typeof control.GetInternalId !== "function") {
      continue;
    }

    if (targetId && String(control.GetInternalId()) === targetId) {
      selectedControl = control;
      break;
    }
  }

  if (!selectedControl && typeof doc.GetCurrentContentControl === "function") {
    selectedControl = doc.GetCurrentContentControl();
  }

  if (!selectedControl || typeof selectedControl.SetTag !== "function") {
    return JSON.stringify({ ok: false });
  }

  const result = selectedControl.SetTag(tag);
  return JSON.stringify({
    ok: result !== false,
    tag,
  });
}

/**
 * Call an Office command and expose a Promise for the iframe-side flow.
 */
function callCommandResult(command) {
  return new Promise((resolve, reject) => {
    try {
      window.Asc.plugin.callCommand(command, false, false, resolve);
    } catch (error) {
      reject(error);
    }
  });
}
function executeMethodResult(method, args) {
  return new Promise((resolve, reject) => {
    try {
      window.Asc.plugin.executeMethod(method, args, resolve);
    } catch (error) {
      reject(error);
    }
  });
}
function restrictEditorToForms() {
  return executeMethodResult("SetEditingRestrictions", ["forms"]);
}
function freezeEditor() {
  return executeMethodResult("SetEditingRestrictions", ["readOnly"]);
}

function saveDocumentCommand() {
  return typeof Api.Save === "function" && Api.Save();
}

function parseCommandResult(result) {
  if (typeof result === "string") {
    return JSON.parse(result || "{}");
  }

  if (result && typeof result === "object") {
    return result;
  }

  return {};
}

/**
 * Extract the current document data.
 *
 * This callback API remains available for existing plugin consumers while
 * action handlers use extractFormDataPromise() below.
 */
function extractFormData(callback) {
  const done = typeof callback === "function" ? callback : () => {};

  try {
    const scope = window.Asc.scope || (window.Asc.scope = {});
    scope.formBridgeTagAliases = runtimeOptions.tagAliases;
    window.Asc.plugin.callCommand(
      extractFormDataCommand,
      false,
      false,
      (result) => {
        try {
          done(null, parseCommandResult(result));
        } catch (error) {
          done(error);
        }
      }
    );
  } catch (error) {
    done(error);
  }
}

function extractFormDataPromise() {
  return new Promise((resolve, reject) => {
    extractFormData((error, data) => {
      if (error) {
        reject(error);
        return;
      }

      resolve(data);
    });
  });
}

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function firstString(...values) {
  for (const value of values) {
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }

  return;
}

function copyPolicies(target, source) {
  if (Array.isArray(source)) {
    for (const field of source) {
      if (typeof field === "string" && field) {
        target[field] = "locked";
      }
    }

    return;
  }

  if (!isRecord(source)) {
    return;
  }

  for (const [field, policy] of Object.entries(source)) {
    target[field] = policy;
  }
}
function policiesForEditableFields(source) {
  if (Array.isArray(source)) {
    return Object.fromEntries(
      source
        .filter((field) => typeof field === "string" && field)
        .map((field) => [field, "editable"])
    );
  }

  if (!isRecord(source)) {
    return;
  }

  return Object.fromEntries(
    Object.entries(source).map(([field, editable]) => [
      field,
      editable === false ? "locked" : "editable",
    ])
  );
}

function normalizePrefill(options) {
  const rawPrefill = isRecord(options.prefill) ? options.prefill : undefined;
  const fieldValues = {};
  const policies = {};
  let valuesCandidate =
    (isRecord(options.prefillValues) && options.prefillValues) ||
    (isRecord(options.prefillData) && options.prefillData) ||
    undefined;

  if (rawPrefill) {
    if (isRecord(rawPrefill.values)) {
      valuesCandidate = rawPrefill.values;
    } else if (isRecord(rawPrefill.data)) {
      valuesCandidate = rawPrefill.data;
    } else if (isRecord(rawPrefill.fields)) {
      valuesCandidate = rawPrefill.fields;
    } else if (!valuesCandidate) {
      valuesCandidate = rawPrefill;
    }

    copyPolicies(policies, rawPrefill.policies);
    copyPolicies(policies, rawPrefill.fieldPolicies);
    copyPolicies(policies, rawPrefill.lockedFields);
    copyPolicies(policies, rawPrefill.locked);
    copyPolicies(
      policies,
      policiesForEditableFields(rawPrefill.editableFields)
    );
    copyPolicies(policies, policiesForEditableFields(rawPrefill.editable));
  }

  if (isRecord(valuesCandidate)) {
    const hasFieldDescriptors =
      (valuesCandidate === rawPrefill ||
        rawPrefill?.fields === valuesCandidate) &&
      !rawPrefill?.values &&
      !rawPrefill?.data;

    if (hasFieldDescriptors) {
      for (const [field, entry] of Object.entries(valuesCandidate)) {
        if (isRecord(entry) && Object.hasOwn(entry, "value")) {
          fieldValues[field] = entry.value;

          if (entry.policy !== undefined) {
            policies[field] = entry.policy;
          } else if (
            entry.locked !== undefined ||
            entry.editable !== undefined
          ) {
            policies[field] = entry;
          }
        } else {
          fieldValues[field] = entry;
        }
      }
    } else {
      Object.assign(fieldValues, valuesCandidate);
    }
  }

  copyPolicies(policies, options.policies);
  copyPolicies(policies, options.fieldPolicies);
  copyPolicies(policies, options.prefillPolicies);
  copyPolicies(policies, options.lockedFields);
  copyPolicies(policies, options.locked);

  copyPolicies(policies, policiesForEditableFields(options.editableFields));

  let defaultPolicy =
    options.prefillPolicy ?? rawPrefill?.policy ?? rawPrefill?.defaultPolicy;

  if (options.prefillLocked === true) {
    defaultPolicy = "locked";
  }

  if (typeof defaultPolicy === "string" || typeof defaultPolicy === "boolean") {
    for (const field of Object.keys(fieldValues)) {
      if (!Object.hasOwn(policies, field)) {
        policies[field] = defaultPolicy;
      }
    }
  } else if (isRecord(defaultPolicy)) {
    copyPolicies(policies, defaultPolicy.fields);
  }

  return {
    policies,
    values: fieldValues,
  };
}

function hasPrefillValues(prefill) {
  return Object.keys(prefill.values).length > 0;
}

function applyPrefill(prefill) {
  const normalizedPrefill = normalizePrefill({
    prefill: isRecord(prefill) ? prefill : {},
  });
  const scope = window.Asc.scope || (window.Asc.scope = {});
  scope.formBridgePrefill = normalizedPrefill;
  scope.formBridgeTagAliases = runtimeOptions.tagAliases;

  return callCommandResult(applyPrefillCommand)
    .then((result) => parseCommandResult(result))
    .finally(() => {
      if (scope.formBridgePrefill === normalizedPrefill) {
        delete scope.formBridgePrefill;
      }
    });
}
function wait(milliseconds) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

async function applyPrefillWhenReady(prefill) {
  for (let attempt = 0; attempt < PREFILL_MAX_ATTEMPTS; attempt += 1) {
    const result = await applyPrefill(prefill);
    if (
      result.applied?.length ||
      result.failed?.length ||
      result.skipped?.length
    ) {
      return result;
    }
    await wait(500);
  }
  return applyPrefill(prefill);
}
function ensurePrefill() {
  if (prefillApplied) {
    return Promise.resolve({ applied: [], failed: [], skipped: [] });
  }

  if (!prefillPromise) {
    prefillPromise = applyPrefillWhenReady(runtimeOptions.prefill).then(
      async (result) => {
        if (result.failed?.length) {
          throw new Error(
            `Could not prefill ${result.failed.length} field(s): ${result.failed.join(", ")}`
          );
        }
        prefillApplied = result.applied?.length > 0;
        await restrictEditorToForms();
        return result;
      }
    );
  }

  return prefillPromise;
}

function panelElement(id) {
  if (typeof document === "undefined") {
    return null;
  }

  if (typeof document.getElementById === "function") {
    const element = document.querySelector(`#${id}`);
    if (element) {
      return element;
    }
  }

  if (typeof document.querySelector === "function") {
    return document.querySelector(`#${id}`);
  }

  return null;
}

function setPanelText(element, value) {
  if (element) {
    element.textContent = String(value);
  }
}

function setPanelDisabled(element, disabled) {
  if (element) {
    element.disabled = disabled;
  }
}

function setPanelStatus(message, state) {
  const status = panelElements.status || panelElement(PANEL_IDS.status);
  if (!status) {
    return;
  }

  status.textContent = message;
  if (status.dataset) {
    status.dataset.state = state || "info";
  }
}

function appendPanelChild(parent, child) {
  if (!parent || !child) {
    return;
  }

  if (typeof parent.append === "function") {
    parent.append(child);
  } else if (typeof parent.appendChild === "function") {
    parent.appendChild(child);
  }
}

function clearPanelChildren(element) {
  if (!element) {
    return;
  }

  if (typeof element.replaceChildren === "function") {
    element.replaceChildren();
    return;
  }

  if (typeof element.removeChild !== "function") {
    return;
  }

  while (element.firstChild) {
    element.removeChild(element.firstChild);
  }
}

function typeLabel(controlType) {
  switch (controlType) {
    case "checkbox": {
      return "ช่องทำเครื่องหมาย";
    }
    case "date": {
      return "วันที่";
    }
    case "dropdown": {
      return "รายการเลือก";
    }
    case "combo": {
      return "รายการเลือกแบบพิมพ์ได้";
    }
    case "picture": {
      return "รูปภาพ";
    }
    case "text": {
      return "ข้อความ";
    }
    default: {
      return "ไม่รองรับ";
    }
  }
}
function schemaTypeLabel(schemaType) {
  switch (schemaType) {
    case "boolean": {
      return "จริง/เท็จ";
    }
    case "number": {
      return "ตัวเลข";
    }
    case "null": {
      return "ค่าว่าง";
    }
    case "string": {
      return "ข้อความ";
    }
    default: {
      return "ไม่รองรับ";
    }
  }
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim() ? value : null;
}

function primitiveString(value) {
  if (typeof value === "string" || typeof value === "number") {
    return String(value);
  }

  return null;
}

function valueFromRecord(record, names) {
  if (!isRecord(record)) {
    return null;
  }

  for (const name of names) {
    const value = primitiveString(record[name]);
    if (value) {
      return value;
    }
  }

  return null;
}

function normalizeControlType(...values) {
  for (const value of values) {
    if (typeof value === "string") {
      const normalized = value.toLowerCase().replaceAll(/[\s_-]+/g, "");

      if (
        normalized === "checkbox" ||
        normalized === "check" ||
        normalized === "checkboxcontentcontrol"
      ) {
        return "checkbox";
      }
      if (normalized === "date" || normalized === "datepicker") {
        return "date";
      }
      if (
        normalized === "dropdown" ||
        normalized === "dropdownlist" ||
        normalized === "select"
      ) {
        return "dropdown";
      }
      if (normalized === "combo" || normalized === "combobox") {
        return "combo";
      }
      if (
        normalized === "picture" ||
        normalized === "image" ||
        normalized === "picturecontentcontrol"
      ) {
        return "picture";
      }
      if (
        normalized === "text" ||
        normalized === "plaintext" ||
        normalized === "richtext" ||
        normalized === "inlinelevel" ||
        normalized === "inlinelevelcontentcontrol" ||
        normalized === "blocklevel" ||
        normalized === "blocklevelcontentcontrol"
      ) {
        return "text";
      }
    }

    if (isRecord(value)) {
      if (value.CheckBox === true || value.checkbox === true) {
        return "checkbox";
      }
      if (value.DatePicker === true || value.datePicker === true) {
        return "date";
      }
      if (value.DropDownList === true || value.dropdownList === true) {
        return "dropdown";
      }
      if (value.ComboBox === true || value.comboBox === true) {
        return "combo";
      }
      if (value.Picture === true || value.picture === true) {
        return "picture";
      }

      const nestedType =
        value.Type ?? value.type ?? value.ControlType ?? value.controlType;
      if (nestedType !== value) {
        const result = normalizeControlType(nestedType);
        if (result !== "unsupported") {
          return result;
        }
      }
    }
  }

  return "unsupported";
}

function normalizeSelectionSnapshot(currentControl, properties, hint) {
  const current =
    isRecord(currentControl) || typeof currentControl === "string"
      ? currentControl
      : hint;
  const props = isRecord(properties) ? properties : {};
  const currentObject = isRecord(current) ? current : {};
  const internalId =
    primitiveString(current) ||
    valueFromRecord(currentObject, ["InternalId", "internalId", "Id", "id"]) ||
    valueFromRecord(props, ["InternalId", "internalId", "Id", "id"]);
  const tag =
    nonEmptyString(currentObject.Tag) ||
    nonEmptyString(currentObject.tag) ||
    nonEmptyString(props.Tag) ||
    nonEmptyString(props.tag);
  const controlType = normalizeControlType(
    currentObject.controlType,
    currentObject.ControlType,
    currentObject.Type,
    props.controlType,
    props.ControlType,
    props.Type,
    props,
    current
  );
  const hasKnownType = controlType !== "unsupported";
  const explicitlySelected =
    currentObject.selected === true ||
    props.selected === true ||
    current === true;

  if (!internalId && !tag && !hasKnownType && !explicitlySelected) {
    return null;
  }

  const controlKey = internalId || `tag:${tag || ""}|type:${controlType}`;
  return {
    controlKey,
    controlType,
    documentTag: tag || null,
    internalId: internalId || null,
    tag: tag || null,
  };
}

async function readCurrentSelection(hint) {
  const plugin = window.Asc?.plugin;
  let currentControl;
  let properties;

  if (plugin && typeof plugin.executeMethod === "function") {
    try {
      currentControl = await executeMethodResult(
        "GetCurrentContentControl",
        []
      );
    } catch {
      currentControl = undefined;
    }

    try {
      properties = await executeMethodResult("GetCurrentContentControlPr", [
        "none",
      ]);
    } catch {
      properties = undefined;
    }

    const selection = normalizeSelectionSnapshot(
      currentControl,
      properties,
      hint
    );
    if (selection?.controlType === "unsupported" && currentControl !== null) {
      try {
        const fallback = normalizeSelectionSnapshot(
          parseCommandResult(
            await callCommandResult(getCurrentContentControlCommand)
          )
        );
        if (fallback) {
          return fallback;
        }
      } catch {
        // Keep the executeMethod snapshot when the Office command is unavailable.
      }
    }
    if (
      selection ||
      currentControl === null ||
      properties === null ||
      (!hint && currentControl !== undefined && properties !== undefined)
    ) {
      return selection;
    }
  }
  if (hint) {
    const selection = normalizeSelectionSnapshot(undefined, undefined, hint);
    if (selection) {
      return selection;
    }
  }

  try {
    const fallback = await callCommandResult(getCurrentContentControlCommand);
    return normalizeSelectionSnapshot(parseCommandResult(fallback));
  } catch {
    return null;
  }
}

function safeFieldSelection(selection) {
  if (!selection || !selection.controlKey) {
    return {
      controlType: "unsupported",
      selected: false,
      tag: null,
    };
  }

  return {
    controlType: FIELD_CONTROL_TYPES.includes(selection.controlType)
      ? selection.controlType
      : "unsupported",
    selected: true,
    tag: selection.tag || null,
  };
}

function parentWindow() {
  const target = window.top;
  if (!target || typeof target.postMessage !== "function") {
    throw new Error("The editor parent window is unavailable");
  }
  return target;
}

function postBridgeMessage(message) {
  const bridgeId = requireOption(runtimeOptions.bridgeId, "bridgeId");
  const parentOrigin = requireOption(
    runtimeOptions.parentOrigin,
    "parentOrigin"
  );
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
  if ((await callCommandResult(saveDocumentCommand)) !== true) {
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

function publishFieldSelection(selection) {
  const safe = safeFieldSelection(selection);
  const current = {
    ...safe,
    selectionId:
      fieldPanelState.selection?.selectionId ||
      `selection-${++selectionSequence}`,
  };
  const message = {
    controlType: current.controlType,
    selected: current.selected,
    selectionId: current.selectionId,
    source: BRIDGE_MESSAGE_SOURCE,
    tag: current.tag,
    type: "field-selection",
  };

  if (!bridgeAcknowledged) {
    pendingFieldSelection = message;
    return message;
  }

  postFieldSelection(message);
  return message;
}
function sameSelectionControl(left, right) {
  return Boolean(left && right && left.controlKey === right.controlKey);
}
function isPictureSelection(selection = fieldPanelState.selection) {
  return selection?.controlType === "picture";
}

function picturePrefillStatus() {
  return "ฟิลด์รูปภาพไม่รองรับการเติมข้อมูลล่วงหน้า";
}

function reportPicturePrefillError() {
  const error = picturePrefillStatus();
  setPanelStatus(error, "info");
  return error;
}

function clearPicturePrefillState() {
  if (!isPictureSelection()) {
    return;
  }

  fieldPanelState.currentPointer = null;
  fieldPanelState.prefillPolicy = "editable";
  fieldPanelState.selectedPointer = null;
}

function applyDefaultPolicyState() {
  fieldPanelState.currentPointer = null;
  fieldPanelState.prefillPolicy = "editable";
  fieldPanelState.required = false;
  fieldPanelState.rules = [];
  fieldPanelState.selectedPointer = null;
}

function hasNextSchemaPage() {
  return (
    Boolean(fieldPanelState.schemaCursor) &&
    fieldPanelState.schemaPageCount < MAX_SCHEMA_PAGES
  );
}

function updateFieldPanel() {
  const { selection } = fieldPanelState;
  const hasSelection = Boolean(selection?.controlKey);
  const hasTag = Boolean(selection?.tag);
  const pictureSelection = isPictureSelection();
  const { query } = panelElements;
  const policy = panelElements.policySelect;
  const { required } = panelElements;

  clearPicturePrefillState();
  setPanelText(
    panelElements.selectionTag,
    hasSelection ? selection.tag || "ยังไม่มีแท็ก" : "ยังไม่ได้เลือก"
  );
  setPanelText(
    panelElements.selectionType,
    hasSelection ? typeLabel(selection.controlType) : "ยังไม่ได้เลือก"
  );
  if (panelElements.pictureHelp) {
    panelElements.pictureHelp.hidden = !pictureSelection;
  }
  if (required) {
    required.checked = fieldPanelState.required;
  }
  if (policy) {
    policy.value = fieldPanelState.prefillPolicy;
  }

  setPanelDisabled(required, !hasSelection);
  setPanelDisabled(policy, !hasSelection || pictureSelection);
  setPanelDisabled(panelElements.save, !hasSelection || !hasTag);
  setPanelDisabled(query, !hasSelection || pictureSelection);
  setPanelDisabled(panelElements.search, !hasSelection || pictureSelection);
  setPanelDisabled(
    panelElements.nextPage,
    !hasSelection || pictureSelection || !hasNextSchemaPage()
  );
  setPanelDisabled(
    panelElements.applyPointer,
    !hasSelection || pictureSelection || !fieldPanelState.selectedPointer
  );
  // oxlint-disable-next-line no-use-before-define -- Hoisted render and pointer handlers form a cycle: pointer actions update the panel, which rebuilds their controls.
  renderSchemaItems();
}

function setSelectionState(snapshot) {
  const previous = fieldPanelState.selection;

  if (!snapshot) {
    const selectionId =
      previous?.selectionId || `selection-${++selectionSequence}`;
    fieldPanelState.selection = {
      controlKey: "",
      controlType: "unsupported",
      documentTag: null,
      internalId: null,
      previousTag: null,
      rulesLoaded: false,
      selectionId,
      tag: null,
    };
    applyDefaultPolicyState();
    updateFieldPanel();
    return false;
  }

  const sameControl = sameSelectionControl(previous, snapshot);
  const selectionId = sameControl
    ? previous.selectionId
    : `selection-${++selectionSequence}`;
  fieldPanelState.selection = {
    controlKey: snapshot.controlKey,
    controlType: snapshot.controlType,
    documentTag: sameControl ? previous.documentTag : snapshot.documentTag,
    internalId: snapshot.internalId,
    previousTag: sameControl ? previous.previousTag : snapshot.tag,
    rulesLoaded: sameControl ? previous.rulesLoaded === true : false,
    selectionId,
    tag: snapshot.tag,
  };
  if (!sameControl) {
    applyDefaultPolicyState();
  }
  clearPicturePrefillState();
  updateFieldPanel();
  return !sameControl;
}

function selectionChanged(previous, next) {
  if (!previous || !next) {
    return Boolean(previous) !== Boolean(next);
  }

  return (
    previous.controlKey !== next.controlKey ||
    previous.tag !== next.tag ||
    previous.controlType !== next.controlType
  );
}

function normalizeFieldRule(value) {
  if (!isRecord(value)) {
    return null;
  }

  const tag = nonEmptyString(value.tag);
  const prefillPointer =
    value.prefillPointer === null ? null : nonEmptyString(value.prefillPointer);
  const prefillPolicy = FIELD_RULE_POLICIES.includes(value.prefillPolicy)
    ? value.prefillPolicy
    : null;
  if (!tag || typeof value.required !== "boolean" || !prefillPolicy) {
    return null;
  }

  return {
    prefillPointer,
    prefillPolicy,
    required: value.required,
    tag,
  };
}

function fieldRulesFromResponse(payload) {
  if (!isRecord(payload) || !Array.isArray(payload.rules)) {
    return [];
  }

  return payload.rules
    .map((rule) => normalizeFieldRule(rule))
    .filter((rule) => rule !== null);
}

function panelErrorStatus(kind) {
  switch (kind) {
    case "capability": {
      return "ไม่สามารถยืนยันสิทธิ์การตั้งค่าฟิลด์ได้";
    }
    case "schema": {
      return "ไม่สามารถโหลดตัวชี้ข้อมูลได้";
    }
    case "rules": {
      return "ไม่สามารถโหลดนโยบายฟิลด์ได้";
    }
    case "save": {
      return "ไม่สามารถบันทึกนโยบายฟิลด์ได้";
    }
    case "selection": {
      return "ไม่สามารถอ่านฟิลด์ที่เลือกได้";
    }
    case "tag": {
      return "ไม่สามารถใช้ตัวชี้เป็นแท็กได้";
    }
    case "clipboard": {
      return "ไม่สามารถคัดลอกตัวชี้ได้";
    }
    default: {
      return "เกิดข้อผิดพลาด กรุณาลองใหม่";
    }
  }
}

function apiUrl(path) {
  const base = runtimeOptions.apiBase;

  if (!base) {
    return path;
  }

  return `${base}${path}`;
}

async function requestJson(path, init, capability) {
  const headers = new Headers();
  headers.set(
    "X-Editor-Capability",
    requireOption(capability, "editor capability")
  );

  if (init?.body !== undefined) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(apiUrl(path), {
    ...init,
    credentials: "omit",
    headers,
  });
  const text = await response.text();
  let payload = {};

  if (text) {
    try {
      payload = JSON.parse(text);
    } catch {
      payload = {
        message: text,
      };
    }
  }

  if (!response.ok) {
    const detail =
      (isRecord(payload) && (payload.error || payload.message)) ||
      `HTTP ${response.status}`;

    throw new Error(String(detail));
  }

  if (isRecord(payload) && payload.ok === false) {
    throw new Error(
      String(payload.error || payload.message || "Request failed")
    );
  }

  return payload;
}

function adminFormPath(suffix) {
  const publicId = encodeURIComponent(
    requireOption(runtimeOptions.publicId, "publicId")
  );
  return `${API_ROUTES.ADMIN_FORMS}/${publicId}/${suffix}`;
}

function isActionAllowedInMode(action) {
  switch (runtimeOptions.action) {
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
  if (!CAPABILITY_ACTIONS.includes(action) || !isActionAllowedInMode(action)) {
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

async function requestFieldApi(path, init) {
  let capability;
  try {
    capability = await requestActionCapability(ACTIONS.CONFIGURE_FIELDS);
  } catch (error) {
    const wrapped = new Error(panelErrorStatus("capability"));
    wrapped.cause = error;
    throw wrapped;
  }

  return requestJson(path, init, capability);
}

async function loadFieldRules() {
  const { selection } = fieldPanelState;
  if (!selection?.controlKey) {
    return { ok: false, rules: [] };
  }
  const pictureSelection = isPictureSelection();

  const { selectionId } = selection;
  try {
    const payload = await requestFieldApi(adminFormPath("field-rules"), {
      method: "GET",
    });
    if (fieldPanelState.selection?.selectionId !== selectionId) {
      return { ignored: true, ok: false, rules: [] };
    }

    const rules = fieldRulesFromResponse(payload);
    const baselineTag = selection.rulesLoaded
      ? selection.previousTag
      : selection.tag;
    const rule = baselineTag
      ? rules.find((candidate) => candidate.tag === baselineTag)
      : undefined;
    fieldPanelState.selection.rulesLoaded = true;
    fieldPanelState.selection.previousTag = rule?.tag || null;
    fieldPanelState.rules = rules;
    fieldPanelState.currentPointer = pictureSelection
      ? null
      : rule?.prefillPointer || null;
    fieldPanelState.prefillPolicy = pictureSelection
      ? "editable"
      : rule?.prefillPolicy || "editable";
    fieldPanelState.required = rule?.required === true;
    fieldPanelState.selectedPointer = null;
    updateFieldPanel();
    setPanelStatus("โหลดนโยบายฟิลด์แล้ว", "success");
    return { ok: true, rule: rule || null, rules };
  } catch {
    if (fieldPanelState.selection?.selectionId === selectionId) {
      setPanelStatus(panelErrorStatus("rules"), "error");
    }
    return { error: panelErrorStatus("rules"), ok: false, rules: [] };
  }
}

function normalizeSchemaItems(payload) {
  if (!isRecord(payload) || !Array.isArray(payload.items)) {
    return {
      items: [],
      nextCursor: null,
    };
  }
  const items = [];
  const seen = new Set();

  for (const item of payload.items) {
    if (
      !isRecord(item) ||
      typeof item.pointer !== "string" ||
      !item.pointer ||
      !["string", "number", "boolean", "null"].includes(item.type)
    ) {
      continue;
    }
    if (seen.has(item.pointer)) {
      continue;
    }
    seen.add(item.pointer);
    items.push({
      pointer: item.pointer,
      type: item.type,
    });
    if (items.length >= MAX_SCHEMA_PAGE_ITEMS) {
      break;
    }
  }

  return {
    items,
    nextCursor:
      typeof payload.nextCursor === "string" && payload.nextCursor
        ? payload.nextCursor
        : null,
  };
}

function schemaPath(query, cursor) {
  const params = [];
  if (query) {
    params.push(`q=${encodeURIComponent(query)}`);
  }
  if (cursor) {
    params.push(`cursor=${encodeURIComponent(cursor)}`);
  }
  const path = adminFormPath("schema");
  return params.length ? `${path}?${params.join("&")}` : path;
}

async function loadSchemaPage(reset = true) {
  const selectionId = fieldPanelState.selection?.selectionId;
  const requestSequence = ++schemaRequestSequence;
  if (!selectionId || !fieldPanelState.selection?.controlKey) {
    return { error: "ยังไม่ได้เลือกฟิลด์", items: [], ok: false };
  }
  if (isPictureSelection()) {
    return { error: reportPicturePrefillError(), items: [], ok: false };
  }
  if (!reset && !hasNextSchemaPage()) {
    return {
      error: "ไม่มีหน้าถัดไป",
      items: fieldPanelState.schemaItems,
      ok: false,
    };
  }

  const query = fieldPanelState.schemaQuery.slice(0, MAX_SCHEMA_QUERY_LENGTH);
  const cursor = reset ? null : fieldPanelState.schemaCursor;
  try {
    const payload = await requestFieldApi(schemaPath(query, cursor), {
      method: "GET",
    });
    if (
      fieldPanelState.selection?.selectionId !== selectionId ||
      requestSequence !== schemaRequestSequence
    ) {
      return { ignored: true, items: [], ok: false };
    }

    const page = normalizeSchemaItems(payload);
    if (reset) {
      fieldPanelState.schemaItems = page.items;
      fieldPanelState.schemaPageCount = 1;
    } else {
      const existing = new Set(
        fieldPanelState.schemaItems.map((item) => item.pointer)
      );
      fieldPanelState.schemaItems = fieldPanelState.schemaItems.concat(
        page.items.filter((item) => !existing.has(item.pointer))
      );
      fieldPanelState.schemaPageCount += 1;
    }
    fieldPanelState.schemaCursor = page.nextCursor;
    updateFieldPanel();
    setPanelStatus(
      page.items.length ? "โหลดตัวชี้ข้อมูลแล้ว" : "ไม่พบตัวชี้ข้อมูล",
      "success"
    );
    return {
      items: page.items,
      nextCursor: page.nextCursor,
      ok: true,
    };
  } catch {
    if (fieldPanelState.selection?.selectionId === selectionId) {
      setPanelStatus(panelErrorStatus("schema"), "error");
    }
    return { error: panelErrorStatus("schema"), items: [], ok: false };
  }
}

function setSchemaQuery(value) {
  schemaRequestSequence += 1;
  fieldPanelState.schemaQuery =
    typeof value === "string" ? value.slice(0, MAX_SCHEMA_QUERY_LENGTH) : "";
}

function selectSchemaPointer(pointer) {
  if (
    !fieldPanelState.selection?.controlKey ||
    typeof pointer !== "string" ||
    !pointer
  ) {
    return { ok: false };
  }
  if (isPictureSelection()) {
    return { error: reportPicturePrefillError(), ok: false };
  }

  fieldPanelState.selectedPointer = pointer;
  fieldPanelState.currentPointer = pointer;
  updateFieldPanel();
  setPanelStatus("เลือกตัวชี้ข้อมูลแล้ว", "info");
  return { ok: true, pointer };
}

async function copySchemaPointer(pointer) {
  if (!fieldPanelState.selection?.controlKey || typeof pointer !== "string") {
    return { error: "ยังไม่ได้เลือกฟิลด์", ok: false };
  }
  if (isPictureSelection()) {
    return { error: reportPicturePrefillError(), ok: false };
  }

  try {
    let copied = false;
    if (
      typeof navigator !== "undefined" &&
      navigator.clipboard &&
      typeof navigator.clipboard.writeText === "function"
    ) {
      try {
        await navigator.clipboard.writeText(pointer);
        copied = true;
      } catch {
        // Fall through to the synchronous browser copy path.
      }
    }
    if (
      !copied &&
      typeof document !== "undefined" &&
      typeof document.execCommand === "function" &&
      typeof document.createElement === "function"
    ) {
      const input = document.createElement("textarea");
      input.value = pointer;
      input.setAttribute("readonly", "true");
      input.style.position = "fixed";
      input.style.opacity = "0";
      appendPanelChild(document.body, input);
      input.select?.();
      copied = document.execCommand("copy");
      input.remove?.();
    }
    if (!copied) {
      throw new Error("Clipboard copy was rejected");
    }
    setPanelStatus("คัดลอกตัวชี้ข้อมูลแล้ว", "success");
    return { ok: true, pointer };
  } catch {
    setPanelStatus(panelErrorStatus("clipboard"), "error");
    return { error: panelErrorStatus("clipboard"), ok: false };
  }
}

async function applySchemaPointer(pointer) {
  if (!fieldPanelState.selection?.controlKey) {
    return { error: "ยังไม่ได้เลือกฟิลด์", ok: false };
  }
  if (isPictureSelection()) {
    return { error: reportPicturePrefillError(), ok: false };
  }
  if (typeof pointer !== "string" || !pointer) {
    return { error: panelErrorStatus("tag"), ok: false };
  }

  const { selection } = fieldPanelState;
  const scope = window.Asc.scope || (window.Asc.scope = {});
  scope.formBridgeSelectionId = selection.internalId || "";
  scope.formBridgeSelectionTag = pointer;
  try {
    const result = parseCommandResult(
      await callCommandResult(setCurrentContentControlTagCommand)
    );
    if (result.ok !== true) {
      throw new Error("Tag update was rejected");
    }
    if (fieldPanelState.selection?.selectionId !== selection.selectionId) {
      return { ignored: true, ok: false };
    }
    fieldPanelState.selectedPointer = pointer;
    fieldPanelState.currentPointer = pointer;
    fieldPanelState.selection.tag = pointer;
    updateFieldPanel();
    publishFieldSelection(fieldPanelState.selection);
    setPanelStatus("ใช้ตัวชี้เป็นแท็กแล้ว", "success");
    return { ok: true, tag: pointer };
  } catch {
    setPanelStatus(panelErrorStatus("tag"), "error");
    return { error: panelErrorStatus("tag"), ok: false };
  } finally {
    if (scope.formBridgeSelectionTag === pointer) {
      delete scope.formBridgeSelectionTag;
    }
    if (scope.formBridgeSelectionId === (selection.internalId || "")) {
      delete scope.formBridgeSelectionId;
    }
  }
}

async function saveFieldRule(overrides) {
  const { selection } = fieldPanelState;
  if (!selection?.controlKey || !selection.tag) {
    return { error: "ยังไม่ได้เลือกฟิลด์", ok: false };
  }
  const pictureSelection = isPictureSelection();
  clearPicturePrefillState();
  const { prefillPointer: overridePrefillPointer } = overrides ?? {};
  let prefillPointer = fieldPanelState.currentPointer;
  if (overridePrefillPointer !== undefined) {
    prefillPointer = overridePrefillPointer;
  }
  let prefillPolicy = overrides?.prefillPolicy || fieldPanelState.prefillPolicy;
  if (pictureSelection) {
    prefillPointer = null;
    prefillPolicy = "editable";
  }

  const body = {
    documentKey: requireOption(runtimeOptions.documentKey, "documentKey"),
    previousTag: selection.previousTag || null,
    prefillPointer,
    prefillPolicy,
    required:
      overrides?.required === undefined
        ? fieldPanelState.required
        : Boolean(overrides.required),
    tag: selection.tag,
  };
  if (
    (body.prefillPointer !== null && typeof body.prefillPointer !== "string") ||
    !FIELD_RULE_POLICIES.includes(body.prefillPolicy)
  ) {
    setPanelStatus(panelErrorStatus("save"), "error");
    return { error: panelErrorStatus("save"), ok: false };
  }

  const { selectionId } = selection;
  const originalDocumentTag = selection.documentTag;
  const tagChanged = body.tag !== originalDocumentTag;

  fieldPanelState.saving = true;
  updateFieldPanel();
  try {
    const payload = await requestFieldApi(adminFormPath("field-rules"), {
      body: JSON.stringify(body),
      method: "PATCH",
    });
    const normalizedRule = normalizeFieldRule(payload?.rule);
    if (!normalizedRule) {
      throw new Error("Invalid field rule response");
    }
    const rule = pictureSelection
      ? {
          ...normalizedRule,
          prefillPolicy: "editable",
          prefillPointer: null,
        }
      : normalizedRule;
    if (fieldPanelState.selection?.selectionId !== selectionId) {
      return { ignored: true, ok: false };
    }
    fieldPanelState.rules = fieldPanelState.rules
      .filter((candidate) => candidate.tag !== body.previousTag)
      .filter((candidate) => candidate.tag !== rule.tag)
      .concat(rule);
    fieldPanelState.currentPointer = rule.prefillPointer;
    fieldPanelState.prefillPolicy = rule.prefillPolicy;
    fieldPanelState.required = rule.required;
    fieldPanelState.selection.previousTag = rule.tag;
    fieldPanelState.selection.documentTag = rule.tag;
    setPanelStatus("บันทึกนโยบายฟิลด์แล้ว", "success");
    return { ok: true, rule };
  } catch {
    if (
      tagChanged &&
      originalDocumentTag &&
      fieldPanelState.selection?.selectionId === selectionId
    ) {
      await applySchemaPointer(originalDocumentTag);
    }
    setPanelStatus(panelErrorStatus("save"), "error");
    return { error: panelErrorStatus("save"), ok: false };
  } finally {
    fieldPanelState.saving = false;
    updateFieldPanel();
  }
}

function renderSchemaItems() {
  const { list } = panelElements;
  clearPanelChildren(list);
  if (
    !list ||
    typeof document === "undefined" ||
    typeof document.createElement !== "function"
  ) {
    return;
  }
  const pictureSelection = isPictureSelection();

  for (const item of fieldPanelState.schemaItems) {
    const row = document.createElement("li");
    const pointerButton = document.createElement("button");
    const type = document.createElement("span");
    const actions = document.createElement("span");
    const copyButton = document.createElement("button");
    const applyButton = document.createElement("button");

    row.className = "schema-item";
    pointerButton.className = "schema-pointer";
    pointerButton.type = "button";
    pointerButton.textContent = item.pointer;
    pointerButton.title = "เลือกตัวชี้";
    pointerButton.disabled = pictureSelection;
    pointerButton.addEventListener?.("click", () => {
      selectSchemaPointer(item.pointer);
    });
    type.className = "schema-type";
    type.textContent = schemaTypeLabel(item.type);
    actions.className = "inline-actions";
    copyButton.className = "secondary-action";
    copyButton.type = "button";
    copyButton.textContent = "คัดลอก";
    copyButton.disabled = pictureSelection;
    copyButton.addEventListener?.("click", () => {
      void copySchemaPointer(item.pointer);
    });
    applyButton.type = "button";
    applyButton.textContent = "ใช้เป็นแท็ก";
    applyButton.disabled = pictureSelection;
    applyButton.addEventListener?.("click", () => {
      void applySchemaPointer(item.pointer);
    });
    appendPanelChild(row, pointerButton);
    appendPanelChild(row, type);
    appendPanelChild(actions, copyButton);
    appendPanelChild(actions, applyButton);
    appendPanelChild(row, actions);
    appendPanelChild(list, row);
  }
}

async function refreshSelection(hint) {
  if (runtimeOptions.action !== ACTIONS.TEMPLATE_EDIT) {
    return { ignored: true, selected: false };
  }

  const requestId = ++panelSelectionRequestSequence;
  const previous = fieldPanelState.selection;
  const snapshot = await readCurrentSelection(hint);
  if (requestId !== panelSelectionRequestSequence) {
    return { ignored: true, selected: false };
  }

  const needsRulesRefresh = selectionChanged(previous, snapshot) || !previous;
  setSelectionState(snapshot);
  if (needsRulesRefresh) {
    publishFieldSelection(fieldPanelState.selection);
  }

  if (!snapshot) {
    setPanelStatus("ยังไม่ได้เลือกฟิลด์", "info");
    return { ok: true, selected: false };
  }

  if (needsRulesRefresh) {
    await loadFieldRules();
  }
  return {
    controlType: snapshot.controlType,
    ok: true,
    selected: true,
    tag: snapshot.tag,
  };
}

function getPanelState() {
  const { selection } = fieldPanelState;
  return {
    currentPointer: fieldPanelState.currentPointer,
    prefillPolicy: fieldPanelState.prefillPolicy,
    required: fieldPanelState.required,
    rules: fieldPanelState.rules.map((rule) => ({ ...rule })),
    schemaCursor: fieldPanelState.schemaCursor,
    schemaItems: fieldPanelState.schemaItems.map((item) => ({ ...item })),
    schemaPageCount: fieldPanelState.schemaPageCount,
    schemaQuery: fieldPanelState.schemaQuery,
    selectedPointer: fieldPanelState.selectedPointer,
    selection: selection
      ? {
          controlType: selection.controlType,
          selected: Boolean(selection.controlKey),
          selectionId: selection.selectionId,
          tag: selection.tag,
        }
      : null,
  };
}

function bindPanelEvent(element, eventName, listener) {
  if (element && typeof element.addEventListener === "function") {
    element.addEventListener(eventName, listener);
    return true;
  }

  return false;
}

function bindFieldPanelEvents() {
  if (panelEventsAttached) {
    return;
  }
  panelEventsAttached = true;

  bindPanelEvent(panelElements.policyForm, "submit", (event) => {
    event.preventDefault?.();
    void saveFieldRule();
  });
  bindPanelEvent(panelElements.required, "change", (event) => {
    fieldPanelState.required = Boolean(event.target?.checked);
  });
  bindPanelEvent(panelElements.policySelect, "change", (event) => {
    const value = event.target?.value;
    if (FIELD_RULE_POLICIES.includes(value)) {
      fieldPanelState.prefillPolicy = value;
    }
  });
  bindPanelEvent(panelElements.query, "input", (event) => {
    setSchemaQuery(event.target?.value);
    fieldPanelState.schemaCursor = null;
    fieldPanelState.schemaPageCount = 0;
    fieldPanelState.schemaItems = [];
    updateFieldPanel();
  });
  bindPanelEvent(panelElements.query, "keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault?.();
      void loadSchemaPage(true);
    }
  });
  bindPanelEvent(panelElements.search, "click", () => {
    void loadSchemaPage(true);
  });
  bindPanelEvent(panelElements.nextPage, "click", () => {
    void loadSchemaPage(false);
  });
  bindPanelEvent(panelElements.applyPointer, "click", () => {
    void applySchemaPointer(fieldPanelState.selectedPointer);
  });
}

function attachFieldSelectionEvents() {
  const plugin = window.Asc?.plugin;
  if (
    panelSelectionEventAttached ||
    !plugin ||
    typeof plugin.attachEditorEvent !== "function"
  ) {
    return;
  }

  panelSelectionEventAttached = true;
  try {
    plugin.attachEditorEvent("onDocumentContentReady", () => {
      void refreshSelection();
    });
    plugin.attachEditorEvent("onTargetPositionChanged", () => {
      void refreshSelection();
    });
    plugin.attachEditorEvent("onFocusContentControl", (control) => {
      void refreshSelection(control);
    });
    plugin.attachEditorEvent("onBlurContentControl", () => {
      void refreshSelection();
    });
  } catch {
    setPanelStatus(panelErrorStatus("selection"), "error");
  }
}

function hideFieldPanel() {
  const panel = panelElements.panel || panelElement(PANEL_IDS.panel);
  if (!panel) {
    return;
  }
  panel.hidden = true;
  panel.setAttribute?.("aria-hidden", "true");
}

function setupFieldPanel() {
  panelElements = {
    applyPointer: panelElement(PANEL_IDS.applyPointer),
    list: panelElement(PANEL_IDS.list),
    nextPage: panelElement(PANEL_IDS.nextPage),
    panel: panelElement(PANEL_IDS.panel),
    pictureHelp: panelElement(PANEL_IDS.pictureHelp),
    policyForm: panelElement(PANEL_IDS.policyForm),
    policySelect: panelElement(PANEL_IDS.policySelect),
    query: panelElement(PANEL_IDS.query),
    required: panelElement(PANEL_IDS.required),
    save: panelElement(PANEL_IDS.save),
    search: panelElement(PANEL_IDS.search),
    selectionTag: panelElement(PANEL_IDS.selectionTag),
    selectionType: panelElement(PANEL_IDS.selectionType),
    status: panelElement(PANEL_IDS.status),
  };
  const { panel } = panelElements;
  if (!panel) {
    return;
  }

  panel.hidden = false;
  panel.setAttribute?.("aria-hidden", "false");
  bindFieldPanelEvents();
  updateFieldPanel();
  attachFieldSelectionEvents();
  void refreshSelection();
}

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

function errorMessage(error) {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error || "Unknown error");
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
    if (typeof message.capability !== "string" || !message.capability.trim()) {
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
    tagAliases: isRecord(options.tagAliases) ? options.tagAliases : {},
    publicId: firstString(options.publicId),
    responseId: firstString(options.responseId),
    targetId: firstString(options.targetId),
  };
}

function actionLabel(action) {
  switch (action) {
    case ACTIONS.SAVE_TEMPLATE: {
      return "บันทึก Template";
    }
    case ACTIONS.PUBLISH: {
      return "เผยแพร่";
    }
    case ACTIONS.SAVE_DRAFT: {
      return "บันทึกฉบับร่าง";
    }
    case ACTIONS.SAVE_CORRECTION: {
      return "บันทึก Correction";
    }
    case ACTIONS.SUBMIT: {
      return "ส่งคำตอบ";
    }
    default: {
      return "การดำเนินการ";
    }
  }
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

function actionRequest(action, data, reason) {
  const documentKey = requireOption(runtimeOptions.documentKey, "documentKey");

  if (action === ACTIONS.SAVE_TEMPLATE || action === ACTIONS.PUBLISH) {
    const path = adminFormPath(
      action === ACTIONS.SAVE_TEMPLATE ? "save" : "publish"
    );

    return {
      body: {
        documentKey,
      },
      path,
    };
  }

  if (action === ACTIONS.SAVE_DRAFT || action === ACTIONS.SUBMIT) {
    const publicId = encodeURIComponent(
      requireOption(runtimeOptions.publicId, "publicId")
    );
    const body = {
      canonicalDateFields: data.canonicalDateFields,
      data: data.data,
      documentKey,
    };

    if (runtimeOptions.responseId) {
      body.responseId = runtimeOptions.responseId;
    }

    return {
      body,
      path: `${API_ROUTES.FORMS}/${publicId}/${
        action === ACTIONS.SAVE_DRAFT ? "draft" : "submit"
      }`,
    };
  }

  if (action === ACTIONS.SAVE_CORRECTION) {
    const responseId = encodeURIComponent(
      requireOption(runtimeOptions.responseId, "responseId")
    );
    return {
      body: {
        data: data.data,
        documentKey,
        reason: requireOption(reason, "correction reason"),
      },
      path: `${API_ROUTES.ADMIN_RESULTS}/${responseId}/correction`,
    };
  }

  throw new Error(`Unsupported form action: ${action || "none"}`);
}

async function postAction(action, data, reason, capability) {
  const request = actionRequest(action, data, reason);
  const result = await requestJson(
    request.path,
    {
      body: JSON.stringify(request.body),
      method: "POST",
    },
    capability
  );

  return result;
}

function operationIdFromResponse(payload) {
  if (!isRecord(payload)) {
    return;
  }

  const operation = isRecord(payload.operation) ? payload.operation : {};
  const data = isRecord(payload.data) ? payload.data : {};

  return firstString(
    payload.operationId,
    operation.operationId,
    operation.id,
    data.operationId
  );
}
function operationCapabilityFromResponse(payload) {
  if (!isRecord(payload)) {
    return;
  }

  return firstString(payload.operationCapability);
}

function operationStatus(payload) {
  const operation = isRecord(payload?.operation) ? payload.operation : {};
  const status = payload?.status ?? operation.status;

  if (typeof status !== "string") {
    return "";
  }

  return status
    .trim()
    .toLowerCase()
    .replaceAll(/[\s-]+/g, "_");
}

function operationFailureMessage(payload, operationId) {
  const operation = isRecord(payload?.operation) ? payload.operation : {};
  const detail = payload?.error ?? payload?.message ?? operation.error;

  return detail ? String(detail) : `การดำเนินการ ${operationId} ไม่สำเร็จ`;
}

async function pollOperation(operationId, label, operationCapability) {
  const operationLabel = label === undefined ? "operation" : label;
  const id = requireOption(operationId, "operationId");
  const capability = requireOption(operationCapability, "operationCapability");
  let previousStatus = "";

  for (let attempt = 0; attempt < MAX_OPERATION_POLLS; attempt += 1) {
    const payload = await requestJson(
      `${API_ROUTES.OPERATIONS}/${encodeURIComponent(id)}`,
      { method: "GET" },
      capability
    );
    const status = operationStatus(payload);

    if (!status) {
      throw new Error(`การดำเนินการ ${id} ไม่มีสถานะ`);
    }

    if (
      status !== previousStatus &&
      (status === "pending" || status === "queued" || status === "processing")
    ) {
      previousStatus = status;
      setStatus(`${operationLabel} กำลังประมวลผล…`, "pending");
    }

    if (
      status === "completed" ||
      status === "complete" ||
      status === "succeeded" ||
      status === "success" ||
      status === "done"
    ) {
      return payload;
    }

    if (
      status === "failed" ||
      status === "failure" ||
      status === "error" ||
      status === "cancelled" ||
      status === "canceled"
    ) {
      throw new Error(operationFailureMessage(payload, id));
    }

    if (attempt < MAX_OPERATION_POLLS - 1) {
      await wait(OPERATION_POLL_INTERVAL_MS);
    }
  }

  throw new Error(`การดำเนินการ ${id} ใช้เวลานานเกินไป`);
}

function shouldApplyRuntimePrefill() {
  return (
    (runtimeOptions.action === ACTIONS.FILL ||
      runtimeOptions.action === ACTIONS.CORRECTION) &&
    hasPrefillValues(runtimeOptions.prefill)
  );
}

async function runAction(action, reason = "") {
  if (actionInFlight || initializationPending) {
    setStatus("มีการดำเนินการของฟอร์มกำลังทำงานอยู่", "pending");
    return { ignored: true, ok: false };
  }

  actionInFlight = true;
  let operationId;
  let completedPayload;
  let editorFrozen = false;
  const correctionReason = typeof reason === "string" ? reason.trim() : "";

  try {
    const needsData =
      action === ACTIONS.SAVE_DRAFT ||
      action === ACTIONS.SAVE_CORRECTION ||
      action === ACTIONS.SUBMIT;
    let data;

    if (needsData) {
      if (shouldApplyRuntimePrefill()) {
        await ensurePrefill();
      } else if (
        runtimeOptions.action === ACTIONS.DRAFT ||
        runtimeOptions.action === ACTIONS.CORRECTION
      ) {
        await restrictEditorToForms();
      }
      await freezeEditor();
      editorFrozen = true;
      setStatus(`กำลังอ่านข้อมูลของ ${actionLabel(action)}…`, "pending");
      data = await extractFormDataPromise();
    }
    if (!needsData) {
      await freezeEditor();
      editorFrozen = true;
    }
    await flushEditorChanges();

    setStatus(`${actionLabel(action)} กำลังรอดำเนินการ…`, "pending");
    const capability = await requestActionCapability(action);
    const response = await postAction(
      action,
      data,
      correctionReason,
      capability
    );
    operationId = operationIdFromResponse(response);
    notifyParent(action, "pending", operationId, response);

    if (operationId) {
      completedPayload = await pollOperation(
        operationId,
        actionLabel(action),
        operationCapabilityFromResponse(response)
      );
    }

    setStatus(`${actionLabel(action)} สำเร็จ`, "success");
    notifyParent(
      action,
      "completed",
      operationId,
      completedPayload || response
    );

    if (needsData) {
      setDirtyState(false);
    }

    return {
      ok: true,
      operationId: operationId || null,
      response,
    };
  } catch (error) {
    const message = errorMessage(error);
    setStatus(`${actionLabel(action)} ไม่สำเร็จ: ${message}`, "error");
    notifyParent(action, "failed", operationId, completedPayload, message);

    return {
      error: message,
      ok: false,
      operationId: operationId || null,
    };
  } finally {
    if (editorFrozen) {
      try {
        await executeMethodResult("SetEditingRestrictions", [
          runtimeOptions.action === ACTIONS.TEMPLATE_EDIT ? "none" : "forms",
        ]);
      } catch {
        setStatus("ไม่สามารถคืนค่าการแก้ไขเฉพาะช่องกรอกได้", "error");
      }
    }
    actionInFlight = false;
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
    event.origin !== runtimeOptions.parentOrigin ||
    !isRecord(message) ||
    message.source !== PARENT_MESSAGE_SOURCE ||
    message.bridgeId !== runtimeOptions.bridgeId ||
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
      runtimeOptions.action === ACTIONS.TEMPLATE_EDIT &&
      fieldPanelState.selection?.controlKey &&
      !fieldPanelState.selection.rulesLoaded
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

function attachToolbarHandlers(actions) {
  for (const action of actions) {
    const buttonId = buttonIdForAction(action);

    if (!buttonId) {
      continue;
    }

    try {
      window.Asc.plugin.attachToolbarMenuClickEvent(buttonId, () => {
        void runAction(action);
      });
    } catch (error) {
      setStatus(
        `Could not attach ${actionLabel(action)}: ${errorMessage(error)}`,
        "error"
      );
    }
  }
}

/**
 * Compatibility entry point for existing callers of the original plugin.
 */
function submitForm() {
  return runAction(ACTIONS.SUBMIT);
}
function startInitializationTasks() {
  if (initializationStarted) {
    return;
  }
  initializationStarted = true;
  const tasks = [];
  if (shouldApplyRuntimePrefill()) {
    window.setTimeout(() => {
      ensurePrefill()
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
      restrictEditorToForms().catch((error) => {
        setStatus(`จำกัดการแก้ไขเอกสารไม่สำเร็จ: ${errorMessage(error)}`, "error");
      });
    }, 3000);
  }

  if (runtimeOptions.operationId) {
    tasks.push(
      pollOperation(
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
        restrictEditorToForms().catch((error) => {
          setStatus(
            `จำกัดการแก้ไขเอกสารไม่สำเร็จ: ${errorMessage(error)}`,
            "error"
          );
        });
      }
      startInitializationTasks();
    });
    plugin.attachEditorEvent("onDocumentContentChanged", () => {
      setDirtyState(true);
    });
    plugin.attachEditorEvent("onChangeContentControl", () => {
      setDirtyState(true);
    });
  }

  window.setTimeout(startInitializationTasks, 3000);
}

function exposeFormBridge() {
  const formBridge = Object.assign(window.FormBridge || {}, {
    applyPrefill,
    extractFormData,
    getRuntimeOptions: () => runtimeOptions,
    pollOperation,
    runAction,
    submitForm,
  });
  if (runtimeOptions.action === ACTIONS.TEMPLATE_EDIT) {
    Object.assign(formBridge, {
      applySchemaPointer,
      copySchemaPointer,
      getFieldRules: loadFieldRules,
      getPanelState,
      loadSchemaPage,
      refreshSelection,
      saveFieldRule,
      selectSchemaPointer,
      setSchemaQuery,
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
}
function initializePlugin() {
  if (pluginInitialized) {
    return;
  }

  pluginInitialized = true;
  runtimeOptions = normalizeRuntimeOptions();
  exposeFormBridge();
  startBridge();

  if (runtimeOptions.action === ACTIONS.TEMPLATE_EDIT) {
    setupFieldPanel();
  } else {
    hideFieldPanel();
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
    restrictEditorToForms().catch((error) => {
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
  setDirtyState(true);
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
