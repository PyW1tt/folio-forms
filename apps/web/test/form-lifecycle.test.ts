import { expect, test, vi } from "bun:test";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";

import { NativeTextForm } from "../src/components/native-text-form";
import {
  getToken,
  safeReturnPath,
  setToken,
  signOut,
} from "../src/lib/api";
import {
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
test("native fields cannot change while draft save is pending", () => {
  const renderField = (operationBusy: boolean) => {
    const markup = renderToStaticMarkup(
      createElement(NativeTextForm, {
        fields: [
          {
            label: "Full name",
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
        onSave: () => {},
        onSubmit: () => {},
        operationBusy,
        values: { full_name: "" },
      })
    );
    return markup.match(/<textarea\b[^>]*>/u)?.[0] ?? "";
  };
  expect(renderField(true)).toContain('disabled=""');
  expect(renderField(false)).not.toContain('disabled=""');
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

test("session warning has an exact five-minute window", () => {
  const now = Date.parse("2026-09-15T00:00:00.000Z");
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
  expect(shouldWarnBeforeSessionExpiry(new Date(now).toISOString(), now)).toBe(
    false
  );
});

test("safe return paths keep only the opaque response id", () => {
  expect(safeReturnPath("/forms/public/fill?responseId=response-1")).toBe(
    "/forms/public/fill?responseId=response-1"
  );
  expect(safeReturnPath("/forms/public/fill?prefill=secret")).toBeNull();
  expect(safeReturnPath("https://attacker.example/claim")).toBeNull();
  expect(safeReturnPath("//attacker.example/claim")).toBeNull();
});

class MemoryStorage implements
  Pick<Storage, "getItem" | "removeItem" | "setItem">
{
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
      Object.defineProperty(
        globalThis,
        "localStorage",
        localStorageDescriptor
      );
    } else {
      Reflect.deleteProperty(globalThis, "localStorage");
    }
  }
});
