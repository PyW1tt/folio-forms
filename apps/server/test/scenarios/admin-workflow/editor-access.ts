// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";

import type { createApp } from "../../../src/app";
import {
  pluginGuid,
  verifyEditorCapability,
  createEditorCapability,
  createOnlyOfficeAuthorization,
  createDocumentAccessToken,
} from "../../../src/onlyoffice";
import { readObject, DOCX_CONTENT_TYPE } from "../../../src/storage";
import {
  jsonHeaders,
  formCreationRequest,
  onlyOfficeBaseUrl,
  documentBaseUrl,
  apiBaseUrl,
} from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import { capabilityHeaders } from "./helpers";
import type {
  BootstrapAndCreationOutput,
  TemplateUploadsOutput,
} from "./setup";

export interface FieldConfigurationInput {
  app: ReturnType<typeof createApp>;
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  formId: BootstrapAndCreationOutput["formId"];
  adminId: BootstrapAndCreationOutput["adminId"];
  templateDocumentKey: BootstrapAndCreationOutput["templateDocumentKey"];
  createdFormRecord: BootstrapAndCreationOutput["createdFormRecord"];
}

export interface FieldConfigurationOutput {
  adminEditor: EditorConfigBody;
  adminLease: {
    expiresAt: string;
    id: string;
    releaseUrl: string;
    renewUrl: string;
  };
  publishCapability: string;
  saveTemplateCapability: string;
  selectedPointer: string;
}

export const runFieldConfiguration = async (
  input: FieldConfigurationInput
): Promise<FieldConfigurationOutput> => {
  const {
    app,
    formPublicId,
    adminBearer,
    formId,
    adminId,
    templateDocumentKey,
    createdFormRecord,
  } = input;

  const adminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(adminEditorResponse.status).toBe(200);
  const adminEditor = (await adminEditorResponse.json()) as EditorConfigBody;
  const adminPluginOptions =
    adminEditor.config.editorConfig.plugins.options[pluginGuid];
  const adminLease = adminEditor.bridge.lease;
  const publishCapability = adminEditor.bridge.capabilities.publish;
  const configureFieldsCapability =
    adminEditor.bridge.capabilities["configure-fields"];
  const saveTemplateCapability =
    adminEditor.bridge.capabilities["save-template"];
  if (
    !adminPluginOptions ||
    !publishCapability ||
    !configureFieldsCapability ||
    !saveTemplateCapability
  ) {
    throw new Error("The Admin editor capabilities were not returned");
  }
  const adminEditorSerialized = JSON.stringify(adminEditor);
  expect(adminEditorSerialized).not.toContain(adminBearer);
  expect(adminEditorSerialized).not.toContain('"authToken"');
  expect(JSON.stringify(adminEditor.config)).not.toContain(publishCapability);
  expect(adminPluginOptions.publicId).toBe(formPublicId);
  expect(adminPluginOptions).not.toHaveProperty("formId");
  expect(adminEditorSerialized).not.toContain(formId);
  expect(adminEditor.bridge.id).toBe(adminPluginOptions.bridgeId);
  expect(adminEditor.bridge.pluginOrigin).toBe(
    new URL(process.env.API_BASE ?? "http://localhost:3000").origin
  );
  expect(adminPluginOptions.parentOrigin).toBe(
    new URL(process.env.CORS_ORIGIN ?? "http://localhost:5173").origin
  );
  expect(adminEditor.apiUrl).toBe(onlyOfficeBaseUrl);
  expect(adminEditor.config.editorConfig.callbackUrl).toBe(
    `${documentBaseUrl}/onlyoffice/callback`
  );
  expect(adminEditor.config.editorConfig.plugins.pluginsData).toEqual([
    `${apiBaseUrl}/onlyoffice-plugin/config.json`,
  ]);
  const pluginRootResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/")
  );
  expect(pluginRootResponse.status).toBe(200);
  expect(pluginRootResponse.headers.get("content-type")).toBe(
    "text/html; charset=utf-8"
  );
  const pluginHtmlResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/index.html")
  );
  expect(pluginHtmlResponse.status).toBe(200);
  expect(pluginHtmlResponse.headers.get("content-type")).toBe(
    "text/html; charset=utf-8"
  );
  const pluginHtml = await pluginHtmlResponse.text();
  const expectedPluginSdkUrl = `${onlyOfficeBaseUrl}/sdkjs-plugins/v1/plugins.js`;
  const escapedPluginSdkUrl = expectedPluginSdkUrl
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
  expect(pluginHtml).toContain(
    `<script src="${escapedPluginSdkUrl}"></script>`
  );

  expect(adminEditor.config.token.length).toBeGreaterThan(20);
  expect(adminLease).toEqual({
    expiresAt: expect.any(String),
    id: expect.any(String),
    releaseUrl: expect.any(String),
    renewUrl: expect.any(String),
  });
  for (const leaseValue of Object.values(adminLease)) {
    expect(JSON.stringify(adminEditor.config)).not.toContain(leaseValue);
  }
  expect(adminPluginOptions).not.toHaveProperty("lease");
  const adminLeaseRow = await prisma.editorLease.findUnique({
    where: { id: adminLease.id },
  });
  if (!adminLeaseRow) {
    throw new Error("The Admin editor lease was not persisted");
  }
  const configureClaims = verifyEditorCapability(configureFieldsCapability);
  expect(configureClaims).toMatchObject({
    action: "configure-fields",
    actorId: adminId,
    documentKey: templateDocumentKey,
    formId,
    leaseId: adminLease.id,
    targetId: createdFormRecord.templateDraft.id,
    targetType: "template-draft",
  });
  const schemaWithoutCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/schema`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(schemaWithoutCapabilityResponse.status).toBe(401);
  expect(await schemaWithoutCapabilityResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const configureHeaders = {
    ...jsonHeaders,
    "X-Editor-Capability": configureFieldsCapability,
  };
  const schemaFirstResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/schema`, {
      headers: configureHeaders,
    })
  );
  expect(schemaFirstResponse.status).toBe(200);
  const schemaFirst = (await schemaFirstResponse.json()) as {
    items: { pointer: string; type: string }[];
    nextCursor: string | null;
  };
  expect(schemaFirst.items).toHaveLength(5);
  expect(schemaFirst.nextCursor).toEqual(expect.any(String));
  expect(
    schemaFirst.items.every(
      (item) =>
        item.type === "string" ||
        item.type === "number" ||
        item.type === "boolean" ||
        item.type === "null"
    )
  ).toBe(true);
  expect(
    schemaFirst.items.some((item) => item.pointer.includes("/contacts/"))
  ).toBe(false);
  const schemaSecondResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?cursor=${encodeURIComponent(schemaFirst.nextCursor ?? "")}`,
      { headers: configureHeaders }
    )
  );
  expect(schemaSecondResponse.status).toBe(200);
  const schemaSecond = (await schemaSecondResponse.json()) as {
    items: { pointer: string; type: string }[];
    nextCursor: string | null;
  };
  expect(schemaSecond.items.length).toBeGreaterThan(0);
  expect(new Set(schemaSecond.items.map((item) => item.pointer)).size).toBe(
    schemaSecond.items.length
  );
  const schemaFilteredResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?q=${encodeURIComponent("ADDRESS")}`,
      { headers: configureHeaders }
    )
  );
  expect(schemaFilteredResponse.status).toBe(200);
  const schemaFiltered = (await schemaFilteredResponse.json()) as {
    items: { pointer: string; type: string }[];
    nextCursor: string | null;
  };
  expect(schemaFiltered.items.length).toBeGreaterThan(0);
  expect(
    schemaFiltered.items.every((item) =>
      item.pointer.toLowerCase().includes("address")
    )
  ).toBe(true);
  const escapedSchemaResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?q=${encodeURIComponent("display")}`,
      { headers: configureHeaders }
    )
  );
  expect(escapedSchemaResponse.status).toBe(200);
  expect(await escapedSchemaResponse.json()).toMatchObject({
    items: [{ pointer: "/account/display~1name", type: "string" }],
  });
  const invalidSchemaCursorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?cursor=invalid`,
      { headers: configureHeaders }
    )
  );
  expect(invalidSchemaCursorResponse.status).toBe(400);
  expect(await invalidSchemaCursorResponse.json()).toMatchObject({
    error: "invalid_schema_cursor",
  });
  const emptyRulesResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        headers: configureHeaders,
      }
    )
  );
  expect(emptyRulesResponse.status).toBe(200);
  expect(await emptyRulesResponse.json()).toEqual({ rules: [] });
  const selectedPointer = schemaFirst.items.find(
    (item) => item.type === "string"
  )?.pointer;
  if (!selectedPointer) {
    throw new Error("Schema did not return a selectable pointer");
  }
  const createRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "lock-when-available",
          previousTag: null,
          required: true,
          tag: "ticket-09-field",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(createRuleResponse.status).toBe(200);
  expect(await createRuleResponse.json()).toEqual({
    rule: {
      prefillPointer: selectedPointer,
      prefillPolicy: "lock-when-available",
      required: true,
      tag: "ticket-09-field",
    },
  });
  const conflictingPointerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "editable",
          previousTag: null,
          required: false,
          tag: "ticket-09-conflict",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(conflictingPointerResponse.status).toBe(409);
  expect(await conflictingPointerResponse.json()).toMatchObject({
    error: "field_rule_conflict",
  });
  const renamedRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "editable",
          previousTag: "ticket-09-field",
          required: false,
          tag: "ticket-09-renamed",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(renamedRuleResponse.status).toBe(200);
  expect(await renamedRuleResponse.json()).toEqual({
    rule: {
      prefillPointer: selectedPointer,
      prefillPolicy: "editable",
      required: false,
      tag: "ticket-09-renamed",
    },
  });
  const reopenedRulesResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        headers: configureHeaders,
      }
    )
  );
  expect(await reopenedRulesResponse.json()).toEqual({
    rules: [
      {
        prefillPointer: selectedPointer,
        prefillPolicy: "editable",
        required: false,
        tag: "ticket-09-renamed",
      },
    ],
  });
  const invalidPolicyResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: null,
          prefillPolicy: "lock-when-available",
          previousTag: "ticket-09-renamed",
          required: false,
          tag: "ticket-09-renamed",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(invalidPolicyResponse.status).toBe(400);
  expect(await invalidPolicyResponse.json()).toMatchObject({
    error: "invalid_field_config",
  });
  const staleDocumentResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: "template-stale",
          prefillPointer: selectedPointer,
          prefillPolicy: "editable",
          previousTag: "ticket-09-renamed",
          required: false,
          tag: "ticket-09-renamed",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(staleDocumentResponse.status).toBe(409);
  expect(await staleDocumentResponse.json()).toMatchObject({
    error: "stale_document",
  });
  const restoredRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "lock-when-available",
          previousTag: "ticket-09-renamed",
          required: true,
          tag: "full_name",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(restoredRuleResponse.status).toBe(200);
  expect(await restoredRuleResponse.json()).toEqual({
    rule: {
      prefillPointer: selectedPointer,
      prefillPolicy: "lock-when-available",
      required: true,
      tag: "full_name",
    },
  });
  for (const requiredTag of ["accept_terms", "department", "start_date"]) {
    const requiredRuleResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
        {
          body: JSON.stringify({
            documentKey: templateDocumentKey,
            prefillPointer: null,
            prefillPolicy: "editable",
            previousTag: null,
            required: true,
            tag: requiredTag,
          }),
          headers: configureHeaders,
          method: "PATCH",
        }
      )
    );
    expect(requiredRuleResponse.status).toBe(200);
  }

  expect(
    Math.abs(
      adminLeaseRow.expiresAt.getTime() -
        adminLeaseRow.createdAt.getTime() -
        90 * 1000
    )
  ).toBeLessThanOrEqual(1000);
  return {
    adminEditor,
    adminLease,
    publishCapability,
    saveTemplateCapability,
    selectedPointer,
  };
};

export interface EditorAccessInput {
  app: ReturnType<typeof createApp>;
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  adminLease: FieldConfigurationOutput["adminLease"];
  competingAdminBearer: TemplateUploadsOutput["competingAdminBearer"];
  publishCapability: FieldConfigurationOutput["publishCapability"];
  adminId: BootstrapAndCreationOutput["adminId"];
  templateDocumentKey: BootstrapAndCreationOutput["templateDocumentKey"];
  formId: BootstrapAndCreationOutput["formId"];
  adminEditor: FieldConfigurationOutput["adminEditor"];
  saveTemplateCapability: FieldConfigurationOutput["saveTemplateCapability"];
  createdFormRecord: BootstrapAndCreationOutput["createdFormRecord"];
}

export interface EditorAccessOutput {
  secondFormPublicId: string;
  initialTemplateBytes: Uint8Array<ArrayBuffer>;
  saveTemplateCapability: string;
}

export const runEditorAccess = async (
  input: EditorAccessInput
): Promise<EditorAccessOutput> => {
  const {
    app,
    formPublicId,
    adminBearer,
    adminLease,
    competingAdminBearer,
    publishCapability,
    adminId,
    templateDocumentKey,
    formId,
    adminEditor,
    createdFormRecord,
  } = input;
  let { saveTemplateCapability } = input;

  const repeatedAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(repeatedAdminEditorResponse.status).toBe(200);
  const repeatedAdminEditor =
    (await repeatedAdminEditorResponse.json()) as EditorConfigBody;
  expect(repeatedAdminEditor.bridge.lease.id).toBe(adminLease.id);

  const competingAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${competingAdminBearer}` },
      }
    )
  );
  expect(competingAdminEditorResponse.status).toBe(409);
  expect(await competingAdminEditorResponse.json()).toMatchObject({
    error: "editor_in_use",
  });

  const renewResponse = await app.handle(
    new Request(new URL(adminLease.renewUrl, "http://test.local").toString(), {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(renewResponse.status).toBe(200);
  const renewedBody = (await renewResponse.json()) as {
    lease?: { expiresAt?: string; id?: string };
  };
  expect(renewedBody).toMatchObject({
    lease: {
      expiresAt: expect.any(String),
      id: adminLease.id,
    },
  });
  const renewedLeaseRow = await prisma.editorLease.findUnique({
    where: { id: adminLease.id },
  });
  if (!renewedLeaseRow || !renewedBody.lease?.expiresAt) {
    throw new Error("The Admin editor lease was not renewed");
  }
  expect(renewedLeaseRow.renewedAt.getTime()).toBeGreaterThanOrEqual(
    renewedLeaseRow.createdAt.getTime()
  );
  expect(
    Math.abs(
      renewedLeaseRow.expiresAt.getTime() -
        renewedLeaseRow.renewedAt.getTime() -
        90 * 1000
    )
  ).toBeLessThanOrEqual(1000);

  const publishClaims = verifyEditorCapability(publishCapability);
  expect(publishClaims).toMatchObject({
    action: "publish",
    actorId: adminId,
    documentKey: templateDocumentKey,
    formId,
    leaseId: adminLease.id,
    leaseProof: expect.any(String),
    role: "admin",
    targetType: "template-draft",
  });
  expect((publishClaims?.expiresAt ?? 0) - (publishClaims?.issuedAt ?? 0)).toBe(
    5 * 60
  );

  const sessionOnlyAdminSaveResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyAdminSaveResponse.status).toBe(401);
  expect(await sessionOnlyAdminSaveResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const sessionOnlyAdminPublishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyAdminPublishResponse.status).toBe(401);
  expect(await sessionOnlyAdminPublishResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const crossActionResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(publishCapability),
      method: "POST",
    })
  );
  expect(crossActionResponse.status).toBe(403);

  const [capabilityHeader, capabilityPayload, capabilitySignature] =
    publishCapability.split(".");
  if (!capabilityHeader || !capabilityPayload || !capabilitySignature) {
    throw new Error("The Admin editor capability was malformed");
  }
  const tamperedPublishCapability = [
    capabilityHeader,
    capabilityPayload,
    `${capabilitySignature[0] === "a" ? "b" : "a"}${capabilitySignature.slice(1)}`,
  ].join(".");
  const tamperedCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(tamperedPublishCapability),
      method: "POST",
    })
  );
  expect(tamperedCapabilityResponse.status).toBe(401);

  if (!publishClaims) {
    throw new Error("The Admin publish capability did not verify");
  }
  const expiredPublishCapability = createEditorCapability({
    action: publishClaims.action,
    actorId: publishClaims.actorId,
    documentKey: publishClaims.documentKey,
    expiresAt: Math.floor(Date.now() / 1000) - 1,
    formId: publishClaims.formId,
    leaseId: publishClaims.leaseId,
    leaseProof: publishClaims.leaseProof,
    role: publishClaims.role,
    targetId: publishClaims.targetId,
    targetType: publishClaims.targetType,
  });
  const expiredCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(expiredPublishCapability),
      method: "POST",
    })
  );
  expect(expiredCapabilityResponse.status).toBe(401);

  const secondCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "blank",
      title: "Capability scope target",
    })
  );
  expect(secondCreateResponse.status).toBe(200);
  const secondCreatedForm = (await secondCreateResponse.json()) as {
    form?: { publicId?: string };
  };
  const secondFormPublicId = secondCreatedForm.form?.publicId;
  const secondFormRecord = secondFormPublicId
    ? await prisma.form.findUnique({
        include: { templateDraft: true },
        where: { publicId: secondFormPublicId },
      })
    : null;
  if (!secondFormPublicId || !secondFormRecord?.templateDraft) {
    throw new Error("The cross-target Form was not created");
  }
  const crossTargetResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${secondFormPublicId}/publish`,
      {
        body: JSON.stringify({
          documentKey: secondFormRecord.templateDraft.documentKey,
        }),
        headers: capabilityHeaders(publishCapability),
        method: "POST",
      }
    )
  );
  expect(crossTargetResponse.status).toBe(403);

  const documentUrl = adminEditor.config.document.url;
  const parsedDocumentUrl = new URL(documentUrl);
  expect(parsedDocumentUrl.origin).toBe(new URL(documentBaseUrl).origin);
  expect(parsedDocumentUrl.pathname).toBe(
    `/onlyoffice/document/${encodeURIComponent(templateDocumentKey)}`
  );
  const unsignedDocumentResponse = await app.handle(new Request(documentUrl));
  expect(unsignedDocumentResponse.status).toBe(401);
  const alteredDocumentUrl = new URL(documentUrl);
  alteredDocumentUrl.searchParams.set(
    "token",
    `${alteredDocumentUrl.searchParams.get("token") ?? ""}x`
  );
  const alteredDocumentResponse = await app.handle(
    new Request(alteredDocumentUrl.toString(), {
      headers: {
        Authorization: createOnlyOfficeAuthorization({
          url: alteredDocumentUrl.toString(),
        }),
      },
    })
  );
  expect(alteredDocumentResponse.status).toBe(401);
  const expiredDocumentUrl = new URL(documentUrl);
  expiredDocumentUrl.searchParams.set(
    "token",
    createDocumentAccessToken(
      templateDocumentKey,
      Math.floor(Date.now() / 1000) - 1
    )
  );
  const expiredDocumentResponse = await app.handle(
    new Request(expiredDocumentUrl.toString(), {
      headers: {
        Authorization: createOnlyOfficeAuthorization({
          url: expiredDocumentUrl.toString(),
        }),
      },
    })
  );
  expect(expiredDocumentResponse.status).toBe(401);
  const signedDocumentResponse = await app.handle(
    new Request(documentUrl, {
      headers: {
        Authorization: createOnlyOfficeAuthorization({ url: documentUrl }),
      },
    })
  );
  expect(signedDocumentResponse.status).toBe(200);
  expect(signedDocumentResponse.headers.get("content-type")).toBe(
    DOCX_CONTENT_TYPE
  );
  const initialTemplateBytes = new Uint8Array(
    await signedDocumentResponse.arrayBuffer()
  );
  expect(initialTemplateBytes.byteLength).toBeGreaterThan(0);

  const forbiddenPluginOriginResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/config.json", {
      headers: { Origin: "https://attacker.example" },
    })
  );
  expect(forbiddenPluginOriginResponse.status).toBe(403);
  const sameOriginPluginConfigResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/config.json")
  );
  expect(sameOriginPluginConfigResponse.status).toBe(200);
  expect(
    sameOriginPluginConfigResponse.headers.get("access-control-allow-origin")
  ).toBeNull();
  const allowedPluginOrigin = new URL(
    process.env.ONLYOFFICE_URL ?? "http://localhost:8080"
  ).origin;
  const pluginConfigResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/config.json", {
      headers: { Origin: allowedPluginOrigin },
    })
  );
  expect(pluginConfigResponse.status).toBe(200);
  expect(pluginConfigResponse.headers.get("access-control-allow-origin")).toBe(
    allowedPluginOrigin
  );
  const releaseResponse = await app.handle(
    new Request(
      new URL(adminLease.releaseUrl, "http://test.local").toString(),
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(releaseResponse.status).toBe(200);
  expect(await releaseResponse.json()).toEqual({ ok: true });
  expect(
    await prisma.editorLease.findUnique({ where: { id: adminLease.id } })
  ).toBeNull();

  const competingAdminEditorAfterReleaseResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${competingAdminBearer}` },
      }
    )
  );
  expect(competingAdminEditorAfterReleaseResponse.status).toBe(200);
  const competingAdminEditor =
    (await competingAdminEditorAfterReleaseResponse.json()) as EditorConfigBody;
  const competingLease = competingAdminEditor.bridge.lease;
  const competingSaveTemplateCapability =
    competingAdminEditor.bridge.capabilities["save-template"];
  if (!competingSaveTemplateCapability) {
    throw new Error("The competing Admin editor capability was not returned");
  }
  expect(competingLease.id).toBeTruthy();
  const releasedCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(releasedCapabilityResponse.status).toBe(409);
  expect(await releasedCapabilityResponse.json()).toMatchObject({
    error: "editor_lease_inactive",
  });

  await prisma.editorLease.update({
    data: { createdAt: new Date(0), expiresAt: new Date(1) },
    where: { id: competingLease.id },
  });
  const reclaimedAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(reclaimedAdminEditorResponse.status).toBe(200);
  const reclaimedAdminEditor =
    (await reclaimedAdminEditorResponse.json()) as EditorConfigBody;
  expect(reclaimedAdminEditor.bridge.lease.id).not.toBe(competingLease.id);
  const reclaimedSaveTemplateCapability =
    reclaimedAdminEditor.bridge.capabilities["save-template"];
  if (!reclaimedSaveTemplateCapability) {
    throw new Error("The reclaimed Admin editor capability was not returned");
  }
  const expiredCompetingCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(competingSaveTemplateCapability),
      method: "POST",
    })
  );
  expect(expiredCompetingCapabilityResponse.status).toBe(409);
  expect(await expiredCompetingCapabilityResponse.json()).toMatchObject({
    error: "editor_lease_inactive",
  });
  saveTemplateCapability = reclaimedSaveTemplateCapability;
  const operationCountBeforeInvalidSaveBodies = await prisma.operation.count({
    where: { formId },
  });
  const extraSaveFieldResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({
        documentKey: templateDocumentKey,
        unexpected: true,
      }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(extraSaveFieldResponse.status).toBe(400);
  expect(await extraSaveFieldResponse.json()).toMatchObject({
    error: "invalid_request",
  });
  const oversizedSaveBodyResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({
        documentKey: templateDocumentKey,
        padding: "x".repeat(8192),
      }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(oversizedSaveBodyResponse.status).toBe(413);
  expect(await oversizedSaveBodyResponse.json()).toMatchObject({
    error: "payload_too_large",
  });
  expect(await prisma.operation.count({ where: { formId } })).toBe(
    operationCountBeforeInvalidSaveBodies
  );
  expect(await readObject(createdFormRecord.templateDraft.objectKey)).toEqual(
    initialTemplateBytes
  );
  return {
    initialTemplateBytes,
    saveTemplateCapability,
    secondFormPublicId,
  };
};
