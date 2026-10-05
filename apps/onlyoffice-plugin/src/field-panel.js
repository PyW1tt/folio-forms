// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import { FIELD_RULE_POLICIES, PANEL_IDS } from "./constants.js";

export function appendPanelChild(parent, child) {
  if (!parent || !child) {
    return;
  }

  if (typeof parent.append === "function") {
    parent.append(child);
  } else if (typeof parent.appendChild === "function") {
    parent.append(child);
  }
}

export function createFieldPanel(getFields) {
  let panelElements = {};
  let panelEventsAttached = false;
  let panelSelectionEventAttached = false;

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

  function updateFieldPanel() {
    const fields = getFields();
    const fieldPanelState = fields.state;
    const { selection } = fieldPanelState;
    const hasSelection = Boolean(selection?.controlKey);
    const hasTag = Boolean(selection?.tag);
    const pictureSelection = fields.isPictureSelection();
    const { query } = panelElements;
    const policy = panelElements.policySelect;
    const { required } = panelElements;

    fields.clearPicturePrefillState();
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
      !hasSelection || pictureSelection || !fields.hasNextSchemaPage()
    );
    setPanelDisabled(
      panelElements.applyPointer,
      !hasSelection || pictureSelection || !fieldPanelState.selectedPointer
    );
    // oxlint-disable-next-line no-use-before-define -- Hoisted render and pointer handlers form a cycle: pointer actions update the panel, which rebuilds their controls.
    renderSchemaItems(fields);
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

  function renderSchemaItems(fields) {
    const fieldPanelState = fields.state;
    const { list } = panelElements;
    clearPanelChildren(list);
    if (
      !list ||
      typeof document === "undefined" ||
      typeof document.createElement !== "function"
    ) {
      return;
    }
    const pictureSelection = fields.isPictureSelection();

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
        fields.selectSchemaPointer(item.pointer);
      });
      type.className = "schema-type";
      type.textContent = schemaTypeLabel(item.type);
      actions.className = "inline-actions";
      copyButton.className = "secondary-action";
      copyButton.type = "button";
      copyButton.textContent = "คัดลอก";
      copyButton.disabled = pictureSelection;
      copyButton.addEventListener?.("click", () => {
        void fields.copySchemaPointer(item.pointer);
      });
      applyButton.type = "button";
      applyButton.textContent = "ใช้เป็นแท็ก";
      applyButton.disabled = pictureSelection;
      applyButton.addEventListener?.("click", () => {
        void fields.applySchemaPointer(item.pointer);
      });
      appendPanelChild(row, pointerButton);
      appendPanelChild(row, type);
      appendPanelChild(actions, copyButton);
      appendPanelChild(actions, applyButton);
      appendPanelChild(row, actions);
      appendPanelChild(list, row);
    }
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
      void getFields().saveFieldRule();
    });
    bindPanelEvent(panelElements.required, "change", (event) => {
      const fieldPanelState = getFields().state;
      fieldPanelState.required = Boolean(event.target?.checked);
    });
    bindPanelEvent(panelElements.policySelect, "change", (event) => {
      const fieldPanelState = getFields().state;
      const value = event.target?.value;
      if (FIELD_RULE_POLICIES.includes(value)) {
        fieldPanelState.prefillPolicy = value;
      }
    });
    bindPanelEvent(panelElements.query, "input", (event) => {
      const fields = getFields();
      const fieldPanelState = fields.state;
      fields.setSchemaQuery(event.target?.value);
      fieldPanelState.schemaCursor = null;
      fieldPanelState.schemaPageCount = 0;
      fieldPanelState.schemaItems = [];
      updateFieldPanel();
    });
    bindPanelEvent(panelElements.query, "keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault?.();
        void getFields().loadSchemaPage(true);
      }
    });
    bindPanelEvent(panelElements.search, "click", () => {
      void getFields().loadSchemaPage(true);
    });
    bindPanelEvent(panelElements.nextPage, "click", () => {
      void getFields().loadSchemaPage(false);
    });
    bindPanelEvent(panelElements.applyPointer, "click", () => {
      const fields = getFields();
      void fields.applySchemaPointer(fields.state.selectedPointer);
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
        void getFields().refreshSelection();
      });
      plugin.attachEditorEvent("onTargetPositionChanged", () => {
        void getFields().refreshSelection();
      });
      plugin.attachEditorEvent("onFocusContentControl", (control) => {
        void getFields().refreshSelection(control);
      });
      plugin.attachEditorEvent("onBlurContentControl", () => {
        void getFields().refreshSelection();
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
    void getFields().refreshSelection();
  }

  return {
    hideFieldPanel,
    panelErrorStatus,
    setPanelStatus,
    setupFieldPanel,
    updateFieldPanel,
  };
}
