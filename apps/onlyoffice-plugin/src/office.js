// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import {
  extractFormDataCommand,
  getCurrentContentControlCommand,
} from "./document-commands.js";
import {
  isRecord,
  nonEmptyString,
  primitiveString,
  valueFromRecord,
  parseCommandResult,
} from "./values.js";

export function createOffice(options) {
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
      scope.formBridgeTagAliases = options.tagAliases;
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
      valueFromRecord(currentObject, [
        "InternalId",
        "internalId",
        "Id",
        "id",
      ]) ||
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

  return {
    callCommandResult,
    executeMethodResult,
    restrictEditorToForms,
    freezeEditor,
    extractFormData,
    extractFormDataPromise,
    readCurrentSelection,
  };
}
