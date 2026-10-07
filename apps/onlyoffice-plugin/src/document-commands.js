// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.

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

function saveDocumentCommand() {
  return typeof Api.Save === "function" && Api.Save();
}

export {
  extractFormDataCommand,
  applyPrefillCommand,
  getCurrentContentControlCommand,
  setCurrentContentControlTagCommand,
  saveDocumentCommand,
};
