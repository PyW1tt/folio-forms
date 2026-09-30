import { expect, test, vi } from "bun:test";

import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";

import {
  comboDisplayValue,
  comboStoredValue,
  NativeForm,
} from "../src/components/native-form";
import type { NativeFormField } from "../src/components/native-form";
import {
  getToken,
  legacySsoReturnPath,
  safeReturnPath,
  setToken,
  signOut,
} from "../src/lib/api";
import {
  afterEditorSave,
  SESSION_WARNING_WINDOW_MS,
  isSaveFlowBusy,
  saveThenDownload,
  shouldBlockDirtyNavigation,
  shouldWarnBeforeSessionExpiry,
} from "../src/lib/form-lifecycle";

test("dirty navigation blocks until an explicit bypass", () => {
  expect(shouldBlockDirtyNavigation(false, false)).toBe(false);
  expect(shouldBlockDirtyNavigation(true, false)).toBe(true);
  expect(shouldBlockDirtyNavigation(true, true)).toBe(false);
});
test("native combo displays option labels and distinguishes blank from empty option", () => {
  const field: NativeFormField = {
    label: "Color",
    options: [
      { displayText: "No color", value: "" },
      { displayText: "Blue", value: "blue" },
      { displayText: "blue", value: "label-match" },
    ],
    pictureMaxBytes: null,
    pictureMaxHeight: null,
    pictureMaxWidth: null,
    placeholder: null,
    position: 0,
    required: false,
    tag: "color",
    type: "combo",
  };
  const markup = renderToStaticMarkup(
    createElement(NativeForm, {
      fields: [field],
      lockedFields: {},
      onChange: () => {},
      onExportDocx: () => {},
      onExportPdf: () => {},
      onPictureChange: () => {},
      onSave: () => {},
      onSubmit: () => {},
      operationBusy: false,
      pictureFiles: {},
      pictureInputKey: 0,
      pictures: {},
      values: { color: "" },
    })
  );

  expect(markup).toContain('value="No color"');
  expect(markup).toContain('<option value="No color">No color</option>');
  expect(comboDisplayValue(field, null)).toBe("");
  expect(comboDisplayValue(field, "")).toBe("No color");
  expect(comboStoredValue(field, "")).toBeNull();
  expect(comboStoredValue(field, "No color")).toBe("");
  expect(comboStoredValue(field, "blue")).toBe("label-match");
  expect(comboStoredValue(field, "Custom color")).toBe("Custom color");
});
const renderNativeTextField = (operationBusy: boolean): string => {
  const markup = renderToStaticMarkup(
    createElement(NativeForm, {
      fields: [
        {
          label: "Full name",
          options: [],
          pictureMaxBytes: null,
          pictureMaxHeight: null,
          pictureMaxWidth: null,
          placeholder: null,
          position: 0,
          required: false,
          tag: "full_name",
          type: "text",
        },
      ],
      lockedFields: {},
      onChange: () => {},
      onExportDocx: () => {},
      onExportPdf: () => {},
      onPictureChange: () => {},
      onSave: () => {},
      onSubmit: () => {},
      operationBusy,
      pictureFiles: {},
      pictureInputKey: 0,
      pictures: {},
      values: { full_name: "" },
    })
  );
  return markup.match(/<textarea\b[^>]*>/u)?.[0] ?? "";
};
test("native fields cannot change while draft save is pending", () => {
  expect(renderNativeTextField(true)).toContain('disabled=""');
  expect(renderNativeTextField(false)).not.toContain('disabled=""');
});

test("save and export never downloads after a failed save", async () => {
  const events: string[] = [];
  await expect(
    saveThenDownload(
      () => {
        events.push("save");
        throw new Error("save failed");
      },
      () => {
        events.push("download");
      }
    )
  ).rejects.toThrow("save failed");
  expect(events).toEqual(["save"]);

  await saveThenDownload(
    () => {
      events.push("save-success");
    },
    () => {
      events.push("download-success");
    }
  );
  expect(events).toEqual(["save", "save-success", "download-success"]);
});

test("save and exit stays blocked while save and export is pending", () => {
  expect(isSaveFlowBusy(false, "docx", false)).toBe(true);
});

test("session warning has an exact five-minute window before one-hour expiry", () => {
  const now = Date.parse("2026-09-15T00:00:00.000Z");
  const expiresAt = new Date(now + 3_600_000).toISOString();
  expect(shouldWarnBeforeSessionExpiry(expiresAt, now)).toBe(false);
  expect(shouldWarnBeforeSessionExpiry(expiresAt, now + 3_300_000)).toBe(true);
  expect(shouldWarnBeforeSessionExpiry(expiresAt, Date.parse(expiresAt))).toBe(
    false
  );
  expect(
    shouldWarnBeforeSessionExpiry(
      new Date(now + SESSION_WARNING_WINDOW_MS).toISOString(),
      now
    )
  ).toBe(true);
  expect(
    shouldWarnBeforeSessionExpiry(
      new Date(now + SESSION_WARNING_WINDOW_MS + 1).toISOString(),
      now
    )
  ).toBe(false);
});

test("reauthentication waits for draft save and stops when save fails", async () => {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, "window");
  const events = new EventTarget();
  const order: string[] = [];
  let resolveSave: ((allowed: boolean) => void) | undefined;
  Object.defineProperty(globalThis, "window", {
    configurable: true,
    value: events,
  });
  events.addEventListener("folio:before-reauth", (event) => {
    const { detail } = event as CustomEvent<{
      handled: boolean;
      resolve: (allowed: boolean) => void;
    }>;
    detail.handled = true;
    order.push("save-start");
    resolveSave = (allowed) => {
      order.push(allowed ? "save-complete" : "save-failed");
      detail.resolve(allowed);
    };
  });
  try {
    const reauthentication = afterEditorSave(() => {
      order.push("sso-start");
    });
    expect(order).toEqual(["save-start"]);
    if (!resolveSave) {
      throw new Error("Draft save request was not observed");
    }
    resolveSave(true);
    await expect(reauthentication).resolves.toBe(true);
    expect(order).toEqual(["save-start", "save-complete", "sso-start"]);

    order.length = 0;
    const failedReauthentication = afterEditorSave(() => {
      order.push("failed-sso-start");
    });
    expect(order).toEqual(["save-start"]);
    if (!resolveSave) {
      throw new Error("Draft save request was not observed");
    }
    resolveSave(false);
    await expect(failedReauthentication).resolves.toBe(false);
    expect(order).toEqual(["save-start", "save-failed"]);
  } finally {
    if (originalWindow) {
      Object.defineProperty(globalThis, "window", originalWindow);
    } else {
      Reflect.deleteProperty(globalThis, "window");
    }
  }
});

test("safe return paths keep only the opaque response id", () => {
  expect(safeReturnPath("/forms/public/fill?responseId=response-1")).toBe(
    "/forms/public/fill?responseId=response-1"
  );
  expect(safeReturnPath("/forms/public/fill?prefill=secret")).toBeNull();
  expect(safeReturnPath("https://attacker.example/claim")).toBeNull();
  expect(safeReturnPath("//attacker.example/claim")).toBeNull();
});
test("legacy SSO return paths match server path and query rules", () => {
  const fillPath = `/forms/${"a".repeat(32)}/fill`;
  const responseId = "123e4567-e89b-12d3-a456-426614174000";
  expect(legacySsoReturnPath(fillPath, `?responseId=${responseId}`)).toBe(
    `${fillPath}?responseId=${responseId}`
  );
  expect(legacySsoReturnPath(fillPath, "")).toBe(fillPath);
  expect(legacySsoReturnPath(fillPath, "?responseId=invalid")).toBe(fillPath);
  expect(
    legacySsoReturnPath(
      fillPath,
      `?responseId=${responseId}&responseId=${responseId}`
    )
  ).toBe(fillPath);
  expect(legacySsoReturnPath(fillPath, `?ResponseId=${responseId}`)).toBe(
    fillPath
  );
  expect(legacySsoReturnPath(fillPath, "?prefill=secret")).toBe(fillPath);
  expect(legacySsoReturnPath("/dashboard", `?responseId=${responseId}`)).toBe(
    "/dashboard"
  );
  expect(legacySsoReturnPath("/receipt/submission-1", "")).toBe("/dashboard");
});

class MemoryStorage implements Pick<
  Storage,
  "getItem" | "removeItem" | "setItem"
> {
  private readonly values = new Map<string, string>();

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }

  setItem(key: string, value: string): void {
    this.values.set(key, String(value));
  }
}

test("sign-out clears token only when server confirms revocation", async () => {
  const localStorageDescriptor = Object.getOwnPropertyDescriptor(
    globalThis,
    "localStorage"
  );
  const storage = new MemoryStorage();
  Object.defineProperty(globalThis, "localStorage", {
    configurable: true,
    value: storage,
  });
  const fetchMock = vi
    .spyOn(globalThis, "fetch")
    .mockResolvedValueOnce(
      Response.json(
        {
          error: "sign_out_cleanup_failed",
          message: "AI authoring cleanup failed after session revocation",
          sessionRevoked: true,
        },
        { status: 500 }
      )
    )
    .mockResolvedValueOnce(
      Response.json({ error: "internal_error" }, { status: 500 })
    );
  try {
    setToken("revoked-session");
    await expect(signOut()).rejects.toMatchObject({
      code: "sign_out_cleanup_failed",
      sessionRevoked: true,
    });
    expect(getToken()).toBeNull();

    setToken("active-session");
    await expect(signOut()).rejects.toMatchObject({
      code: "internal_error",
      sessionRevoked: false,
    });
    expect(getToken()).toBe("active-session");
  } finally {
    fetchMock.mockRestore();
    if (localStorageDescriptor) {
      Object.defineProperty(globalThis, "localStorage", localStorageDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, "localStorage");
    }
  }
});
