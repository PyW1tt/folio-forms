// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import { ACTIONS, FIELD_CONTROL_TYPES } from "./constants.js";

function requireOption(value, name) {
  if (!value) {
    throw new Error(`Missing editor option: ${name}`);
  }

  return value;
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

function parseCommandResult(result) {
  if (typeof result === "string") {
    return JSON.parse(result || "{}");
  }

  if (result && typeof result === "object") {
    return result;
  }

  return {};
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

function wait(milliseconds) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

function errorMessage(error) {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error || "Unknown error");
}

export {
  requireOption,
  isRecord,
  firstString,
  nonEmptyString,
  primitiveString,
  valueFromRecord,
  parseCommandResult,
  safeFieldSelection,
  actionLabel,
  wait,
  errorMessage,
};
