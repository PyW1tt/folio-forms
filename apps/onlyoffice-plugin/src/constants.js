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

export {
  ACTIONS,
  BUTTON_IDS,
  API_ROUTES,
  CAPABILITY_ACTIONS,
  FIELD_CONTROL_TYPES,
  FIELD_RULE_POLICIES,
  PANEL_IDS,
  BRIDGE_MESSAGE_SOURCE,
  PARENT_MESSAGE_SOURCE,
  BRIDGE_READY_TYPE,
  BRIDGE_ACK_TYPE,
  CAPABILITY_REQUEST_TYPE,
  CAPABILITY_RESPONSE_TYPE,
  DIRTY_STATE_TYPE,
  RUN_ACTION_TYPE,
  CLEAR_DIRTY_TYPE,
  OPERATION_MESSAGE_TYPE,
  EDITOR_SAVE_REQUEST_TYPE,
  EDITOR_SAVE_RESPONSE_TYPE,
  CAPABILITY_REQUEST_TIMEOUT_MS,
  OPERATION_POLL_INTERVAL_MS,
  MAX_OPERATION_POLLS,
  PREFILL_MAX_ATTEMPTS,
  MAX_SCHEMA_QUERY_LENGTH,
  MAX_SCHEMA_PAGE_ITEMS,
  MAX_SCHEMA_PAGES,
};
