// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { unzipSync, strToU8, zipSync } from "fflate";

import { readObject, putObject, DOCX_CONTENT_TYPE } from "../../../src/storage";
import {
  contentControl,
  docxXmlFixture,
  contentControlDocument,
} from "../../fixtures/documents";
import {
  createCredentialFixture,
  bearerFor,
  formCreationRequest,
  waitForOperation,
} from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import { capabilityHeaders, patchFillMethod } from "./helpers";
import type { ScenarioApp } from "./helpers";

export interface NativeSetupInput {
  app: ScenarioApp;
}

export interface NativeSetupOutput {
  adminBearer: string;
  userBearer: string;
  publicId: string;
  responseId: string;
  editorConfigUrl: string;
  prefillValues: {
    full_name: string;
    state_checkbox: boolean;
  };
  prefillLocks: {
    full_name: boolean;
    state_checkbox: boolean;
  };
}

export const runNativeSetup = async (
  input: NativeSetupInput
): Promise<NativeSetupOutput> => {
  const { app } = input;

  const adminPassword = "Ticket06-admin-password";
  const userPassword = "Ticket06-user-password";
  const adminEmail = `ticket-06-admin-${crypto.randomUUID()}@example.com`;
  const userEmail = `ticket-06-user-${crypto.randomUUID()}@example.com`;
  await createCredentialFixture({
    email: adminEmail,
    name: "Ticket 06 Admin",
    password: adminPassword,
    role: "admin",
  });
  await createCredentialFixture({
    email: userEmail,
    name: "Ticket 06 User",
    password: userPassword,
  });
  const adminBearer = await bearerFor(app, adminEmail, adminPassword);
  const userBearer = await bearerFor(app, userEmail, userPassword);
  const fullNameControl = contentControl({
    alias: "Full name",
    placeholderText: "Enter full name",
    tag: " full_name ",
    type: "<w:text/>",
  })
    .replace(
      '<w:tag w:val=" full_name "/>',
      '<w:tag xmlns:ext="urn:fixture" ext:val="metadata" w:val=" \nfull_name "/>'
    )
    .replace(
      "<w:sdtContent><w:r><w:t>Enter full name</w:t></w:r></w:sdtContent>",
      '<w:sdtContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Enter</w:t></w:r></w:p><w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> full name</w:t></w:r></w:p></w:sdtContent>'
    )
    .replace("<w:showingPlcHdr/>", "<w:showingPlcHdr></w:showingPlcHdr>");
  const stateCheckboxControl = contentControl({
    alias: "Checkbox with state children",
    tag: "state_checkbox",
    type: '<w14:checkbox><w14:checkedState w14:val="2713" w14:font="Ticket Symbols"/><w14:uncheckedState w14:val="25A1" w14:font="Ticket Symbols"/></w14:checkbox>',
  })
    .replace('<w:tag w:val="state_checkbox"/>', "")
    .replace(
      "</w14:checkbox>",
      '</w14:checkbox><w:tag w:val="state_checkbox"/>'
    );
  const branchCheckboxControl = stateCheckboxControl.replaceAll(
    "w14:",
    "branch:"
  );
  const commentsControl = contentControl({
    alias: "Comments",
    placeholderText: "Add comments",
    tag: "comments",
    type: "<w:text/>",
  }).replace(
    "<w:sdtContent><w:r><w:t>Add comments</w:t></w:r></w:sdtContent>",
    '<w:sdtContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>Add comments</w:t></w:r></w:p></w:sdtContent>'
  );
  const valueCheckboxControl = contentControl({
    alias: "Checkbox with unchecked marker value",
    tag: "value_checkbox",
    type: '<w14:checkbox><w14:checked/><w14:checkedState w14:val="2713" w14:font="Ticket Symbols"/><w14:uncheckedState w14:val="25A1" w14:font="Ticket Symbols"/></w14:checkbox>',
  });
  const formTemplate = docxXmlFixture({
    document: contentControlDocument(
      `<w:p><w:r><w:t>Static layout</w:t></w:r></w:p>` +
        `<mc:AlternateContent><mc:Choice xmlns:branch="http://schemas.microsoft.com/office/word/2010/wordml" Requires="branch">${
          fullNameControl
        }${commentsControl}${contentControl({
          alias: "Enabled",
          tag: "enabled",
          type: "<w14:checkbox/>",
        })}${
          branchCheckboxControl
        }</mc:Choice><mc:Fallback><w:sdt><w:sdtPr/></w:sdt></mc:Fallback></mc:AlternateContent>${
          valueCheckboxControl
        }${contentControl({
          alias: "Start date",
          tag: "start_date",
          type: '<w:date w:fullDate="2020-01-01T00:00:00Z"><w:lid w:val="en-US"/><w:dateFormat w:val="dddd, MMMM d, yyyy &apos;d literal&apos;"/></w:date>',
        })}${contentControl({
          alias: "Cleared date",
          tag: "cleared_date",
          type: '<w:date w:fullDate="1999-12-31T00:00:00Z"><w:dateFormat w:val="dd/MM/yy"/></w:date>',
        })}${contentControl({
          alias: "Category",
          tag: "category",
          type: '<w:dropDownList w:lastValue="old_value"><w:listItem w:displayText="Friendly label" w:value="stored_value"/><w:listItem w:displayText="Empty option label" w:value=""/><w:listItem w:displayText="Choice B" w:value="Empty option label"/></w:dropDownList>',
        })}${contentControl({
          alias: "Empty category",
          tag: "empty_category",
          type: '<w:dropDownList><w:listItem w:displayText="Empty option label" w:value=""/><w:listItem w:displayText="Whitespace option label" w:value=" "/></w:dropDownList>',
        })}${contentControl({
          alias: "Nullable category",
          tag: "nullable_category",
          type: '<w:dropDownList><w:listItem w:displayText="Nullable empty option label" w:value=""/></w:dropDownList>',
        })}${contentControl({
          alias: "Cleared category",
          tag: "cleared_category",
          type: '<w:dropDownList w:lastValue="old_value"><w:listItem w:displayText="Old label" w:value="old_value"/></w:dropDownList>',
        })}${contentControl({
          alias: "Custom category",
          tag: "custom_category",
          type: '<w:comboBox><w:listItem w:displayText="Suggested label" w:value="suggested"/><w:listItem w:displayText="Empty combo label" w:value=""/><w:listItem w:displayText="Choice B" w:value="Empty combo label"/><w:listItem w:displayText="Whitespace combo label" w:value=" "/></w:comboBox>',
        })}`
    ),
  });
  const createResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes: formTemplate, name: "native-text.docx" },
      title: "Ticket 06 Native Text Form",
    })
  );
  expect(createResponse.status).toBe(200);
  const createdBody = (await createResponse.json()) as {
    form?: { publicId?: string };
  };
  const publicId = createdBody.form?.publicId;
  if (!publicId) {
    throw new Error("The Ticket 06 text form was not created");
  }
  const createdForm = await prisma.form.findUniqueOrThrow({
    select: { fillMethod: true, id: true },
    where: { publicId },
  });
  expect(createdForm.fillMethod).toBe("onlyoffice");
  const prematureNativeResponse = await patchFillMethod(
    app,
    publicId,
    adminBearer,
    "native"
  );
  expect(prematureNativeResponse.status).toBe(409);
  expect(await prematureNativeResponse.json()).toMatchObject({
    error: "native_fill_unsupported",
  });
  const adminEditorResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/editor-config`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminEditorResponse.status).toBe(200);
  const adminEditor = (await adminEditorResponse.json()) as EditorConfigBody;
  const configureFieldsCapability =
    adminEditor.bridge.capabilities["configure-fields"];
  if (!configureFieldsCapability) {
    throw new Error("The Ticket 06 field configuration capability is missing");
  }
  const requiredFieldResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/field-rules`, {
      body: JSON.stringify({
        documentKey: adminEditor.config.document.key,
        prefillPointer: null,
        prefillPolicy: "editable",
        previousTag: null,
        required: true,
        tag: "full_name",
      }),
      headers: capabilityHeaders(configureFieldsCapability),
      method: "PATCH",
    })
  );
  expect(requiredFieldResponse.status).toBe(200);
  const requiredCheckboxResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/field-rules`, {
      body: JSON.stringify({
        documentKey: adminEditor.config.document.key,
        prefillPointer: null,
        prefillPolicy: "editable",
        previousTag: null,
        required: true,
        tag: "enabled",
      }),
      headers: capabilityHeaders(configureFieldsCapability),
      method: "PATCH",
    })
  );
  expect(requiredCheckboxResponse.status).toBe(200);
  const requiredEmptyDropdownResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/field-rules`, {
      body: JSON.stringify({
        documentKey: adminEditor.config.document.key,
        prefillPointer: null,
        prefillPolicy: "editable",
        previousTag: null,
        required: true,
        tag: "empty_category",
      }),
      headers: capabilityHeaders(configureFieldsCapability),
      method: "PATCH",
    })
  );
  expect(requiredEmptyDropdownResponse.status).toBe(200);
  const requiredComboResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/field-rules`, {
      body: JSON.stringify({
        documentKey: adminEditor.config.document.key,
        prefillPointer: null,
        prefillPolicy: "editable",
        previousTag: null,
        required: true,
        tag: "custom_category",
      }),
      headers: capabilityHeaders(configureFieldsCapability),
      method: "PATCH",
    })
  );
  expect(requiredComboResponse.status).toBe(200);
  const publishEditorResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/editor-config`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  const publishEditor =
    (await publishEditorResponse.json()) as EditorConfigBody;
  const publishCapability = publishEditor.bridge.capabilities.publish;
  if (!publishCapability) {
    throw new Error("The Ticket 06 publish capability is missing");
  }
  const publishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}/publish`, {
      body: JSON.stringify({
        documentKey: publishEditor.config.document.key,
      }),
      headers: capabilityHeaders(publishCapability),
      method: "POST",
    })
  );
  expect(publishResponse.status).toBe(202);
  const publishBody = (await publishResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!publishBody.operationCapability || !publishBody.operationId) {
    throw new Error("The Ticket 06 publish operation was not created");
  }
  const publishOperation = await waitForOperation(
    app,
    publishBody.operationId,
    {
      "X-Editor-Capability": publishBody.operationCapability,
    }
  );
  expect(publishOperation.status).toBe("completed");
  const publishedTemplate = await prisma.publishedTemplate.findUniqueOrThrow({
    include: {
      manifest: {
        include: { fields: { orderBy: { position: "asc" } } },
      },
    },
    where: { formId: createdForm.id },
  });
  const publishedManifest = publishedTemplate.manifest;
  if (!publishedManifest) {
    throw new Error("The Ticket 06 manifest was not published");
  }
  const publishedTemplateBytes = await readObject(publishedTemplate.objectKey);
  const unsupportedDateArchive = unzipSync(publishedTemplateBytes);
  const publishedDocumentXml = new TextDecoder().decode(
    unsupportedDateArchive["word/document.xml"]
  );
  unsupportedDateArchive["word/document.xml"] = strToU8(
    publishedDocumentXml.replace(
      'w:dateFormat w:val="dddd, MMMM d, yyyy &apos;d literal&apos;"',
      'w:dateFormat w:val="yyyy-MM-dd HH:mm"'
    )
  );
  await putObject(
    publishedTemplate.objectKey,
    zipSync(unsupportedDateArchive),
    DOCX_CONTENT_TYPE
  );
  const unsupportedDateMethod = await patchFillMethod(
    app,
    publicId,
    adminBearer,
    "native"
  );
  await putObject(
    publishedTemplate.objectKey,
    publishedTemplateBytes,
    DOCX_CONTENT_TYPE
  );
  expect(unsupportedDateMethod.status).toBe(409);
  expect(await unsupportedDateMethod.json()).toMatchObject({
    error: "native_fill_unsupported",
  });
  const assertNativeFillRejected = async (documentXml: string) => {
    const modifiedArchive = unzipSync(publishedTemplateBytes);
    modifiedArchive["word/document.xml"] = strToU8(documentXml);
    await putObject(
      publishedTemplate.objectKey,
      zipSync(modifiedArchive),
      DOCX_CONTENT_TYPE
    );
    try {
      const response = await patchFillMethod(
        app,
        publicId,
        adminBearer,
        "native"
      );
      expect(response.status).toBe(409);
      expect(await response.json()).toMatchObject({
        error: "native_fill_unsupported",
      });
    } finally {
      await putObject(
        publishedTemplate.objectKey,
        publishedTemplateBytes,
        DOCX_CONTENT_TYPE
      );
      await patchFillMethod(app, publicId, adminBearer, "onlyoffice");
    }
  };
  const duplicateChoiceDocumentXml = publishedDocumentXml.replace(
    "</mc:Choice>",
    `</mc:Choice><mc:Choice xmlns:branch="http://schemas.microsoft.com/office/word/2010/wordml" Requires="branch">${fullNameControl}</mc:Choice>`
  );
  if (duplicateChoiceDocumentXml === publishedDocumentXml) {
    throw new Error("The published Choice fixture is missing");
  }
  await assertNativeFillRejected(duplicateChoiceDocumentXml);
  const supportedChoiceOpen =
    '<mc:Choice xmlns:branch="http://schemas.microsoft.com/office/word/2010/wordml" Requires="branch">';
  const earlierSupportedChoiceDocumentXml = publishedDocumentXml.replace(
    supportedChoiceOpen,
    `${supportedChoiceOpen}<w:p><w:r><w:t>Preferred branch without native fields</w:t></w:r></w:p></mc:Choice>${supportedChoiceOpen}`
  );
  if (earlierSupportedChoiceDocumentXml === publishedDocumentXml) {
    throw new Error("The published Choice fixture is missing");
  }
  await assertNativeFillRejected(earlierSupportedChoiceDocumentXml);
  const fallbackFieldDocumentXml = publishedDocumentXml.replace(
    "</mc:Fallback>",
    `${contentControl({ tag: "full_name", type: "<w:text/>" })}</mc:Fallback>`
  );
  await assertNativeFillRejected(fallbackFieldDocumentXml);
  const unsupportedChoiceDocumentXml = publishedDocumentXml.replace(
    'xmlns:branch="http://schemas.microsoft.com/office/word/2010/wordml" Requires="branch"',
    'xmlns:branch="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:unsupported="urn:unsupported" Requires="branch unsupported"'
  );
  await assertNativeFillRejected(unsupportedChoiceDocumentXml);
  await prisma.$transaction(async (tx) => {
    await tx.$executeRaw`ALTER TABLE "field_manifests" DISABLE TRIGGER "field_manifests_immutable"`;
    await tx.$executeRaw`ALTER TABLE "manifest_fields" DISABLE TRIGGER "manifest_fields_immutable"`;
    try {
      await tx.$executeRaw`
        UPDATE "manifest_fields"
        SET "position" = "position" + 100000
        WHERE "manifest_id" = ${publishedManifest.id}::uuid
      `;
      await tx.$executeRaw`
        WITH ranked_fields AS (
          SELECT
            "id",
            (ROW_NUMBER() OVER (ORDER BY "tag") - 1)::integer AS "position"
          FROM "manifest_fields"
          WHERE "manifest_id" = ${publishedManifest.id}::uuid
        )
        UPDATE "manifest_fields" AS field
        SET
          "label" = field."tag",
          "placeholder" = NULL,
          "position" = ranked_fields."position"
        FROM ranked_fields
        WHERE field."id" = ranked_fields."id"
      `;
      await tx.fieldManifest.update({
        data: { displayMetadataVersion: 0 },
        where: { id: publishedManifest.id },
      });
    } finally {
      await tx.$executeRaw`ALTER TABLE "manifest_fields" ENABLE TRIGGER "manifest_fields_immutable"`;
      await tx.$executeRaw`ALTER TABLE "field_manifests" ENABLE TRIGGER "field_manifests_immutable"`;
    }
  });
  const nativeMethodResponse = await patchFillMethod(
    app,
    publicId,
    adminBearer,
    "native"
  );
  expect(nativeMethodResponse.status).toBe(200);
  expect(await nativeMethodResponse.json()).toMatchObject({
    form: { fillMethod: "native" },
  });
  const adminDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${publicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(await adminDetailResponse.json()).toMatchObject({
    form: { fillMethod: "native", nativeFillAvailable: true },
  });
  const publicFormResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(await publicFormResponse.json()).toMatchObject({
    form: { fillMethod: "native" },
  });
  const nativeStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(nativeStartResponse.status).toBe(200);
  const startBody = (await nativeStartResponse.json()) as {
    editorConfigUrl?: string;
    fillMethod?: string;
    response?: { id?: string };
  };
  const responseId = startBody.response?.id;
  const { editorConfigUrl } = startBody;
  if (!responseId || !editorConfigUrl) {
    throw new Error("The Ticket 06 native response did not start");
  }
  expect(startBody.fillMethod).toBe("native");
  const prefillValues = {
    full_name: "Trusted\nPrefill",
    state_checkbox: true,
  };
  const prefillLocks = {
    full_name: true,
    state_checkbox: true,
  };
  await prisma.prefillSnapshot.update({
    data: { lockedFields: prefillLocks, values: prefillValues },
    where: { responseId },
  });
  return {
    adminBearer,
    editorConfigUrl,
    prefillLocks,
    prefillValues,
    publicId,
    responseId,
    userBearer,
  };
};
