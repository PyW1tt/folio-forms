// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import { API_ROUTES } from "./constants.js";
import { isRecord, requireOption } from "./values.js";

export function createApi(options) {
  function apiUrl(path) {
    const base = options.apiBase;

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
      requireOption(options.publicId, "publicId")
    );
    return `${API_ROUTES.ADMIN_FORMS}/${publicId}/${suffix}`;
  }

  return { adminFormPath, requestJson };
}
