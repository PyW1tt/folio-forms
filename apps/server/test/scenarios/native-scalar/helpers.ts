import { expect } from "bun:test";

import type { createApp } from "../../../src/app";
import { jsonHeaders, onlyOfficeSaveCapabilityFor } from "../../fixtures/http";

export type ScenarioApp = ReturnType<typeof createApp>;

export interface NativeResponseConfig {
  capabilities: Record<"save-draft" | "submit", string>;
  data: Record<string, unknown>;
  documentKey: string;
  fields: {
    label: string;
    options: { displayText: string; value: string }[];
    placeholder: string | null;
    position: number;
    required: boolean;
    tag: string;
    type: string;
  }[];
  fillMethod: string;
  lockedFields: Record<string, boolean>;
  responseId: string;
}

export interface OnlyOfficeResponseConfig {
  config?: {
    document?: { key?: string };
    editorConfig?: {
      plugins?: {
        options?: Record<
          string,
          {
            prefill?: {
              data: Record<string, unknown>;
              editableFields: Record<string, unknown>;
            };
            tagAliases?: Record<string, string>;
          }
        >;
      };
    };
  };
  bridge?: {
    capabilities?: Record<"save-draft" | "submit", string>;
  };
  fillMethod?: string;
}

export const capabilityHeaders = (
  capability: string
): Record<string, string> => ({
  ...jsonHeaders,
  "X-Editor-Capability": capability,
});

export const patchFillMethod = (
  app: ScenarioApp,
  publicId: string,
  adminBearer: string,
  fillMethod: "native" | "onlyoffice"
) =>
  app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}`, {
      body: JSON.stringify({ fillMethod }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );

export const getNativeConfig = async (
  app: ScenarioApp,
  editorConfigUrl: string,
  userBearer: string
): Promise<NativeResponseConfig> => {
  const response = await app.handle(
    new Request(new URL(editorConfigUrl, "http://test.local").toString(), {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(response.status).toBe(200);
  return (await response.json()) as NativeResponseConfig;
};

export const nativeDraftRequest = (
  app: ScenarioApp,
  publicId: string,
  responseId: string,
  config: NativeResponseConfig,
  data: Record<string, unknown>,
  targetApp: ScenarioApp = app
): Promise<Response> =>
  targetApp.handle(
    new Request(`http://test.local/api/forms/${publicId}/draft`, {
      body: JSON.stringify({
        data,
        documentKey: config.documentKey,
        fillMethod: "native",
        responseId,
      }),
      headers: capabilityHeaders(config.capabilities["save-draft"]),
      method: "POST",
    })
  );

export const nativeSubmitRequest = (
  app: ScenarioApp,
  publicId: string,
  responseId: string,
  config: NativeResponseConfig,
  data: Record<string, unknown>
) =>
  app.handle(
    new Request(`http://test.local/api/forms/${publicId}/submit`, {
      body: JSON.stringify({
        data,
        documentKey: config.documentKey,
        fillMethod: "native",
        responseId,
      }),
      headers: capabilityHeaders(config.capabilities.submit),
      method: "POST",
    })
  );

export const onlyOfficeDraftRequest = (
  app: ScenarioApp,
  publicId: string,
  responseId: string,
  onlyOfficeDocumentKey: string,
  onlyOfficeSaveCapability: string,
  data: Record<string, unknown>
) =>
  app.handle(
    new Request(`http://test.local/api/forms/${publicId}/draft`, {
      body: JSON.stringify({
        data,
        documentKey: onlyOfficeDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(onlyOfficeSaveCapability),
      method: "POST",
    })
  );

export const refreshOnlyOfficeEditorConfig = async (
  app: ScenarioApp,
  publicId: string,
  userBearer: string
) => {
  const refreshedStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  const start = (await refreshedStartResponse.json()) as {
    editorConfigUrl?: string;
  };
  if (!start.editorConfigUrl) {
    throw new Error("The saved ONLYOFFICE response did not reopen");
  }
  const configResponse = await app.handle(
    new Request(
      new URL(start.editorConfigUrl, "http://test.local").toString(),
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  const config = (await configResponse.json()) as OnlyOfficeResponseConfig;
  const documentKey = config.config?.document?.key;
  if (!documentKey) {
    throw new Error("The refreshed ONLYOFFICE document key is missing");
  }
  const onlyOfficeSaveCapability = onlyOfficeSaveCapabilityFor(config);
  const onlyOfficeDocumentKey = documentKey;
  return { onlyOfficeDocumentKey, onlyOfficeSaveCapability };
};
