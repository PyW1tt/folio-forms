import { expect } from "bun:test";

import type { createApp } from "../../../src/app";
import {
  jsonHeaders,
  formCreationRequest,
  waitForOperation,
} from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";

export const capabilityHeaders = (
  capability: string
): Record<string, string> => ({
  ...jsonHeaders,
  "X-Editor-Capability": capability,
});

export const publishFixture = async (
  app: ReturnType<typeof createApp>,
  adminBearer: string,
  label: string,
  bytes: Uint8Array
): Promise<{ operation: Record<string, unknown>; publicId: string }> => {
  const createFixtureResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes, name: `${label}.docx` },
      title: `Ticket 10 ${label}`,
    })
  );
  expect(createFixtureResponse.status).toBe(200);
  const fixtureBody = (await createFixtureResponse.json()) as {
    form?: { publicId?: string };
  };
  const fixturePublicId = fixtureBody.form?.publicId;
  if (!fixturePublicId) {
    throw new Error(`The ${label} fixture did not receive a public ID`);
  }
  const fixtureEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${fixturePublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(fixtureEditorResponse.status).toBe(200);
  const fixtureEditor =
    (await fixtureEditorResponse.json()) as EditorConfigBody;
  const fixturePublishCapability = fixtureEditor.bridge.capabilities.publish;
  if (!fixturePublishCapability) {
    throw new Error(`The ${label} fixture publish capability was not returned`);
  }
  const response = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${fixturePublicId}/publish`,
      {
        body: JSON.stringify({
          documentKey: fixtureEditor.config.document.key,
        }),
        headers: capabilityHeaders(fixturePublishCapability),
        method: "POST",
      }
    )
  );
  expect(response.status).toBe(202);
  const body = (await response.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!body.operationCapability || !body.operationId) {
    throw new Error(`The ${label} fixture operation was not created`);
  }
  const operation = await waitForOperation(app, body.operationId, {
    "X-Editor-Capability": body.operationCapability,
  });
  return { operation, publicId: fixturePublicId };
};

export const draftRequest = (
  app: ReturnType<typeof createApp>,
  publicId: string,
  responseDocumentKey: string,
  responseId: string,
  saveDraftCapability: string,
  data: Record<string, unknown>
): Promise<Response> =>
  app.handle(
    new Request(`http://test.local/api/forms/${publicId}/draft`, {
      body: JSON.stringify({
        data,
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );

export const submitRequest = (
  app: ReturnType<typeof createApp>,
  publicId: string,
  responseDocumentKey: string,
  responseId: string,
  submitCapability: string,
  data: Record<string, unknown>,
  targetApp: ReturnType<typeof createApp> = app
): Promise<Response> =>
  targetApp.handle(
    new Request(`http://test.local/api/forms/${publicId}/submit`, {
      body: JSON.stringify({
        data,
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(submitCapability),
      method: "POST",
    })
  );

export const refreshResponseEditor = async (
  app: ReturnType<typeof createApp>,
  publicId: string,
  responseId: string,
  userBearer: string
): Promise<RefreshResponseEditorOutput> => {
  const refreshedResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${publicId}/editor-config?responseId=${responseId}&action=draft`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(refreshedResponse.status).toBe(200);
  const refreshedConfig = (await refreshedResponse.json()) as EditorConfigBody;
  const refreshedSaveDraftCapability =
    refreshedConfig.bridge.capabilities["save-draft"];
  const refreshedSubmitCapability = refreshedConfig.bridge.capabilities.submit;
  if (!refreshedSaveDraftCapability || !refreshedSubmitCapability) {
    throw new Error("The refreshed User editor capabilities were not returned");
  }
  const responseDocumentKey = refreshedConfig.config.document.key;
  const saveDraftCapability = refreshedSaveDraftCapability;
  const submitCapability = refreshedSubmitCapability;
  const userLease = refreshedConfig.bridge.lease;
  return {
    responseDocumentKey,
    saveDraftCapability,
    submitCapability,
    userLease,
  };
};

export const accountRequest = (
  app: ReturnType<typeof createApp>,
  method: string,
  pathname: string,
  token: string,
  body?: Record<string, unknown>
): Promise<Response> => {
  const headers =
    body === undefined
      ? { Authorization: `Bearer ${token}` }
      : { ...jsonHeaders, Authorization: `Bearer ${token}` };
  return app.handle(
    new Request(`http://test.local${pathname}`, {
      body: body === undefined ? undefined : JSON.stringify(body),
      headers,
      method,
    })
  );
};

export const replacePassword = (
  app: ReturnType<typeof createApp>,
  token: string,
  oldPassword: string,
  newPassword: string
): Promise<Response> =>
  app.handle(
    new Request("http://test.local/api/account/password", {
      body: JSON.stringify({
        currentPassword: oldPassword,
        newPassword,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${token}` },
      method: "POST",
    })
  );

export const sessionStatus = async (
  app: ReturnType<typeof createApp>,
  token: string
): Promise<number> => {
  const response = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${token}` },
    })
  );
  return response.status;
};

export interface RefreshResponseEditorOutput {
  responseDocumentKey: string;
  saveDraftCapability: string;
  submitCapability: string;
  userLease: EditorConfigBody["bridge"]["lease"];
}
