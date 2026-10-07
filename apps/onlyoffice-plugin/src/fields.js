// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import {
  ACTIONS,
  BRIDGE_MESSAGE_SOURCE,
  FIELD_RULE_POLICIES,
  MAX_SCHEMA_PAGES,
  MAX_SCHEMA_PAGE_ITEMS,
  MAX_SCHEMA_QUERY_LENGTH,
} from "./constants.js";
import { setCurrentContentControlTagCommand } from "./document-commands.js";
import { appendPanelChild } from "./field-panel.js";
import {
  isRecord,
  nonEmptyString,
  parseCommandResult,
  requireOption,
  safeFieldSelection,
} from "./values.js";

export function createFields(options, office, bridge, panel, api) {
  let selectionSequence = 0;
  let schemaRequestSequence = 0;
  const fieldPanelState = {
    currentPointer: null,
    prefillPolicy: "editable",
    required: false,
    rules: [],
    saving: false,
    schemaCursor: null,
    schemaItems: [],
    schemaPageCount: 0,
    schemaQuery: "",
    selectedPointer: null,
    selection: null,
  };
  let panelSelectionRequestSequence = 0;

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

    return bridge.sendFieldSelection(message);
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
    panel.setPanelStatus(error, "info");
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
      panel.updateFieldPanel();
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
    panel.updateFieldPanel();
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
      value.prefillPointer === null
        ? null
        : nonEmptyString(value.prefillPointer);
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

  async function requestFieldApi(path, init) {
    let capability;
    try {
      capability = await bridge.requestActionCapability(
        ACTIONS.CONFIGURE_FIELDS
      );
    } catch (error) {
      const wrapped = new Error(panel.panelErrorStatus("capability"));
      wrapped.cause = error;
      throw wrapped;
    }

    return api.requestJson(path, init, capability);
  }

  async function loadFieldRules() {
    const { selection } = fieldPanelState;
    if (!selection?.controlKey) {
      return { ok: false, rules: [] };
    }
    const pictureSelection = isPictureSelection();

    const { selectionId } = selection;
    try {
      const payload = await requestFieldApi(api.adminFormPath("field-rules"), {
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
      panel.updateFieldPanel();
      panel.setPanelStatus("โหลดนโยบายฟิลด์แล้ว", "success");
      return { ok: true, rule: rule || null, rules };
    } catch {
      if (fieldPanelState.selection?.selectionId === selectionId) {
        panel.setPanelStatus(panel.panelErrorStatus("rules"), "error");
      }
      return { error: panel.panelErrorStatus("rules"), ok: false, rules: [] };
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
    const path = api.adminFormPath("schema");
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
      panel.updateFieldPanel();
      panel.setPanelStatus(
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
        panel.setPanelStatus(panel.panelErrorStatus("schema"), "error");
      }
      return { error: panel.panelErrorStatus("schema"), items: [], ok: false };
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
    panel.updateFieldPanel();
    panel.setPanelStatus("เลือกตัวชี้ข้อมูลแล้ว", "info");
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
      panel.setPanelStatus("คัดลอกตัวชี้ข้อมูลแล้ว", "success");
      return { ok: true, pointer };
    } catch {
      panel.setPanelStatus(panel.panelErrorStatus("clipboard"), "error");
      return { error: panel.panelErrorStatus("clipboard"), ok: false };
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
      return { error: panel.panelErrorStatus("tag"), ok: false };
    }

    const { selection } = fieldPanelState;
    const scope = window.Asc.scope || (window.Asc.scope = {});
    scope.formBridgeSelectionId = selection.internalId || "";
    scope.formBridgeSelectionTag = pointer;
    try {
      const result = parseCommandResult(
        await office.callCommandResult(setCurrentContentControlTagCommand)
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
      panel.updateFieldPanel();
      publishFieldSelection(fieldPanelState.selection);
      panel.setPanelStatus("ใช้ตัวชี้เป็นแท็กแล้ว", "success");
      return { ok: true, tag: pointer };
    } catch {
      panel.setPanelStatus(panel.panelErrorStatus("tag"), "error");
      return { error: panel.panelErrorStatus("tag"), ok: false };
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
    let prefillPolicy =
      overrides?.prefillPolicy || fieldPanelState.prefillPolicy;
    if (pictureSelection) {
      prefillPointer = null;
      prefillPolicy = "editable";
    }

    const body = {
      documentKey: requireOption(options.documentKey, "documentKey"),
      prefillPointer,
      prefillPolicy,
      previousTag: selection.previousTag || null,
      required:
        overrides?.required === undefined
          ? fieldPanelState.required
          : Boolean(overrides.required),
      tag: selection.tag,
    };
    if (
      (body.prefillPointer !== null &&
        typeof body.prefillPointer !== "string") ||
      !FIELD_RULE_POLICIES.includes(body.prefillPolicy)
    ) {
      panel.setPanelStatus(panel.panelErrorStatus("save"), "error");
      return { error: panel.panelErrorStatus("save"), ok: false };
    }

    const { selectionId } = selection;
    const originalDocumentTag = selection.documentTag;
    const tagChanged = body.tag !== originalDocumentTag;

    fieldPanelState.saving = true;
    panel.updateFieldPanel();
    try {
      const payload = await requestFieldApi(api.adminFormPath("field-rules"), {
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
            prefillPointer: null,
            prefillPolicy: "editable",
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
      panel.setPanelStatus("บันทึกนโยบายฟิลด์แล้ว", "success");
      return { ok: true, rule };
    } catch {
      if (
        tagChanged &&
        originalDocumentTag &&
        fieldPanelState.selection?.selectionId === selectionId
      ) {
        await applySchemaPointer(originalDocumentTag);
      }
      panel.setPanelStatus(panel.panelErrorStatus("save"), "error");
      return { error: panel.panelErrorStatus("save"), ok: false };
    } finally {
      fieldPanelState.saving = false;
      panel.updateFieldPanel();
    }
  }

  async function refreshSelection(hint) {
    if (options.action !== ACTIONS.TEMPLATE_EDIT) {
      return { ignored: true, selected: false };
    }

    const requestId = ++panelSelectionRequestSequence;
    const previous = fieldPanelState.selection;
    const snapshot = await office.readCurrentSelection(hint);
    if (requestId !== panelSelectionRequestSequence) {
      return { ignored: true, selected: false };
    }

    const needsRulesRefresh = selectionChanged(previous, snapshot) || !previous;
    setSelectionState(snapshot);
    if (needsRulesRefresh) {
      publishFieldSelection(fieldPanelState.selection);
    }

    if (!snapshot) {
      panel.setPanelStatus("ยังไม่ได้เลือกฟิลด์", "info");
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

  return {
    applySchemaPointer,
    clearPicturePrefillState,
    copySchemaPointer,
    getPanelState,
    hasNextSchemaPage,
    isPictureSelection,
    loadFieldRules,
    loadSchemaPage,
    publishFieldSelection,
    refreshSelection,
    saveFieldRule,
    selectSchemaPointer,
    setSchemaQuery,
    state: fieldPanelState,
  };
}
