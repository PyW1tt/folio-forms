// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";
import { createHash } from "node:crypto";

import type {
  PublishedTemplate,
  FieldManifest,
  ManifestField,
  PrefillConfiguration,
  PrefillField,
} from "@onlyoffice/db";
import { prisma } from "@onlyoffice/db";
import { strToU8 } from "fflate";

import type { createApp } from "../../../src/app";
import {
  verifyEditorCapability,
  createOnlyOfficeAuthorization,
} from "../../../src/onlyoffice";
import { readObject } from "../../../src/storage";
import {
  docxFixture,
  docxXmlFixture,
  contentControlDocument,
  contentControl,
} from "../../fixtures/documents";
import { waitForOperation } from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";
import type {
  EditorAccessOutput,
  FieldConfigurationOutput,
} from "./editor-access";
import { capabilityHeaders, publishFixture } from "./helpers";
import type { BootstrapAndCreationOutput } from "./setup";

export interface PrimaryPublicationInput {
  app: ReturnType<typeof createApp>;
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  templateDocumentKey: BootstrapAndCreationOutput["templateDocumentKey"];
  saveTemplateCapability: EditorAccessOutput["saveTemplateCapability"];
  formId: BootstrapAndCreationOutput["formId"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  initialTemplateBytes: EditorAccessOutput["initialTemplateBytes"];
  selectedPointer: FieldConfigurationOutput["selectedPointer"];
}

export interface PrimaryPublicationOutput {
  refreshedAdminEditor: EditorConfigBody;
  activeTemplateDocumentKey: string;
  activePublishCapability: string;
  publishedManifestRecord: PublishedTemplate & {
    manifest: FieldManifest & { fields: ManifestField[] };
    prefillConfiguration:
      | (PrefillConfiguration & { fields: PrefillField[] })
      | null;
  };
  publishedManifest: FieldManifest & { fields: ManifestField[] };
  publishedBytes: Uint8Array<ArrayBufferLike>;
}

export const runPrimaryPublication = async (
  input: PrimaryPublicationInput
): Promise<PrimaryPublicationOutput> => {
  const {
    app,
    formPublicId,
    templateDocumentKey,
    saveTemplateCapability,
    formId,
    adminBearer,
    initialTemplateBytes,
    selectedPointer,
  } = input;

  const saveTemplateResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(saveTemplateResponse.status).toBe(202);
  const saveTemplateBody = (await saveTemplateResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!saveTemplateBody.operationCapability || !saveTemplateBody.operationId) {
    throw new Error("The template save operation was not created");
  }
  const saveTemplatePollClaims = verifyEditorCapability(
    saveTemplateBody.operationCapability
  );
  expect(saveTemplatePollClaims).toMatchObject({
    action: "poll-operation",
    operationId: saveTemplateBody.operationId,
  });
  expect(
    (saveTemplatePollClaims?.expiresAt ?? 0) -
      (saveTemplatePollClaims?.issuedAt ?? 0)
  ).toBe(6 * 60);
  const saveTemplateOperation = await waitForOperation(
    app,
    saveTemplateBody.operationId,
    {
      "X-Editor-Capability": saveTemplateBody.operationCapability,
    }
  );
  expect(saveTemplateOperation.status).toBe("completed");
  expect(saveTemplateOperation).toMatchObject({
    result: { publicId: formPublicId },
  });
  expect(JSON.stringify(saveTemplateOperation)).not.toContain(formId);

  const refreshedAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(refreshedAdminEditorResponse.status).toBe(200);
  const refreshedAdminEditor =
    (await refreshedAdminEditorResponse.json()) as EditorConfigBody;
  const activeTemplateDocumentKey = refreshedAdminEditor.config.document.key;
  const activePublishCapability =
    refreshedAdminEditor.bridge.capabilities.publish;
  if (!activeTemplateDocumentKey || !activePublishCapability) {
    throw new Error("The refreshed Admin editor capability was not returned");
  }
  expect(activeTemplateDocumentKey).not.toBe(templateDocumentKey);
  const reopenedDocumentUrl = refreshedAdminEditor.config.document.url;
  const reopenedDocumentResponse = await app.handle(
    new Request(reopenedDocumentUrl, {
      headers: {
        Authorization: createOnlyOfficeAuthorization({
          url: reopenedDocumentUrl,
        }),
      },
    })
  );
  expect(reopenedDocumentResponse.status).toBe(200);
  expect(new Uint8Array(await reopenedDocumentResponse.arrayBuffer())).toEqual(
    initialTemplateBytes
  );
  const publishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: activeTemplateDocumentKey }),
      headers: capabilityHeaders(activePublishCapability),
      method: "POST",
    })
  );
  expect(publishResponse.status).toBe(202);
  const publishBody = (await publishResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!publishBody.operationCapability || !publishBody.operationId) {
    throw new Error("The publish operation was not created");
  }
  const publishPollClaims = verifyEditorCapability(
    publishBody.operationCapability
  );
  expect(publishPollClaims).toMatchObject({
    action: "poll-operation",
    operationId: publishBody.operationId,
  });
  expect(
    (publishPollClaims?.expiresAt ?? 0) - (publishPollClaims?.issuedAt ?? 0)
  ).toBe(6 * 60);
  const publishOperation = await waitForOperation(
    app,
    publishBody.operationId,
    {
      "X-Editor-Capability": publishBody.operationCapability,
    }
  );
  expect(publishOperation.status).toBe("completed");
  expect(publishOperation).toMatchObject({
    result: { publicId: formPublicId, version: 1 },
  });
  expect(JSON.stringify(publishOperation)).not.toContain(formId);
  const publishedListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(publishedListResponse.status).toBe(200);
  const publishedList = (await publishedListResponse.json()) as {
    forms?: { publicId: string; status: string; version: number }[];
  };
  expect(
    publishedList.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({ status: "published", version: 1 });
  const publishedManifestRecord = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: { include: { fields: true } },
    },
    where: { formId },
  });
  if (!publishedManifestRecord?.manifest) {
    throw new Error("The published manifest was not persisted");
  }
  const publishedManifest = publishedManifestRecord.manifest;
  const publishedBytes = await readObject(publishedManifestRecord.objectKey);
  const expectedPublishedHash = createHash("sha256")
    .update(publishedBytes)
    .digest("hex");
  expect(publishedManifestRecord).toMatchObject({
    contentHash: expectedPublishedHash,
    version: 1,
  });
  expect(publishedManifestRecord.manifest).toMatchObject({
    configurationHash: expectedPublishedHash,
    displayMetadataVersion: 1,
  });
  expect(publishedManifestRecord.manifest.fields).toMatchObject([
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: true,
      tag: "accept_terms",
      type: "checkbox",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: [
        { displayText: "Choose an item", value: "" },
        { displayText: "Engineering", value: "engineering" },
        { displayText: "Human Resources", value: "hr" },
        { displayText: "Finance", value: "finance" },
      ],
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: true,
      tag: "department",
      type: "dropdown",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: false,
      tag: "description_1",
      type: "text",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: false,
      tag: "description_2",
      type: "text",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "lock_when_available",
      required: true,
      tag: "full_name",
      type: "text",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: true,
      tag: "start_date",
      type: "date",
    },
  ]);
  expect(publishedManifestRecord.prefillConfiguration?.fields).toEqual([
    {
      configurationId: expect.any(String),
      id: expect.any(String),
      pointer: selectedPointer,
      policy: "lock_when_available",
      tag: "full_name",
    },
  ]);
  expect(publishedManifestRecord.prefillConfiguration).toMatchObject({
    configurationHash: expectedPublishedHash,
    publishedTemplateId: publishedManifestRecord.id,
  });
  return {
    activePublishCapability,
    activeTemplateDocumentKey,
    publishedBytes,
    publishedManifest,
    publishedManifestRecord: {
      ...publishedManifestRecord,
      manifest: publishedManifestRecord.manifest,
    },
    refreshedAdminEditor,
  };
};

export interface ValidTemplateContractsInput {
  app: ReturnType<typeof createApp>;
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
}

export const runValidTemplateContracts = async (
  input: ValidTemplateContractsInput
): Promise<void> => {
  const { app, adminBearer } = input;
  const validExtendedFixture = docxXmlFixture({
    document: contentControlDocument(
      contentControl({
        alias: "Applicant photo",
        tag: "photo",
        type: "<w:picture/>",
      }) +
        contentControl({
          alias: "Applicant name",
          placeholderText: "Enter applicant name",
          tag: "/person/name",
          type: "<w:text/>",
        }) +
        contentControl({
          alias: "   ",
          placeholderText: "Choose a department",
          tag: "department_choice",
          type: `<w:comboBox><w:listItem w:displayText="Engineering" w:value="engineering"/><w:listItem w:displayText="Finance" w:value="finance"/></w:comboBox>`,
        })
    ),
  });
  const validExtendedResult = await publishFixture(
    app,
    adminBearer,
    "combo-and-picture",
    validExtendedFixture
  );
  expect(validExtendedResult.operation.status).toBe("completed");
  const validExtendedForm = await prisma.form.findUniqueOrThrow({
    select: { id: true },
    where: { publicId: validExtendedResult.publicId },
  });
  const validExtendedPublished =
    await prisma.publishedTemplate.findUniqueOrThrow({
      include: {
        manifest: { include: { fields: { orderBy: { position: "asc" } } } },
      },
      where: { formId: validExtendedForm.id },
    });
  if (!validExtendedPublished.manifest) {
    throw new Error("The extended fixture manifest was not persisted");
  }
  const validExtendedFields = validExtendedPublished.manifest.fields;
  expect(validExtendedFields).toHaveLength(3);
  expect(
    validExtendedFields.map(({ label, placeholder, position, tag }) => ({
      label,
      placeholder,
      position,
      tag,
    }))
  ).toEqual([
    {
      label: "Applicant photo",
      placeholder: null,
      position: 0,
      tag: "photo",
    },
    {
      label: "Applicant name",
      placeholder: "Enter applicant name",
      position: 1,
      tag: "/person/name",
    },
    {
      label: "department_choice",
      placeholder: "Choose a department",
      position: 2,
      tag: "department_choice",
    },
  ]);
  expect(
    validExtendedFields.find((field) => field.tag === "/person/name")
  ).toMatchObject({
    options: null,
    pictureMaxBytes: null,
    pictureMaxHeight: null,
    pictureMaxWidth: null,
    tag: "/person/name",
    type: "text",
  });
  expect(
    validExtendedFields.find((field) => field.tag === "department_choice")
  ).toMatchObject({
    options: [
      { displayText: "Engineering", value: "engineering" },
      { displayText: "Finance", value: "finance" },
    ],
    type: "combo",
  });
  expect(
    validExtendedFields.find((field) => field.tag === "photo")
  ).toMatchObject({
    options: null,
    pictureMaxBytes: 10 * 1024 * 1024,
    pictureMaxHeight: 4096,
    pictureMaxWidth: 4096,
    type: "picture",
  });
  const orderedPlaceholderFixture = docxXmlFixture({
    additionalParts: {
      "word/_rels/document.xml.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdGlossary" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>`
      ),
      "word/glossary/document.xml": strToU8(
        `<?xml version="1.0"?><w:glossaryDocument xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:docParts><w:docPart><w:docPartPr><w:docPartName w:val="ReceiptPrompt"/></w:docPartPr><w:docPartBody><w:p><w:r><w:t>Configured glossary placeholder</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`
      ),
    },
    document: contentControlDocument(
      `<w:sdt><w:sdtPr><w:alias w:val="Parent label"/><w:tag w:val="parent"/><w:text/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:alias w:val="Child label"/><w:tag w:val="child"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Child answer</w:t></w:r></w:sdtContent></w:sdt></w:sdtContent></w:sdt>` +
        `<w:sdt><w:sdtPr><w:alias w:val="Prompt label"/><w:tag w:val="prompt"/><w:placeholder><w:docPart w:val="ReceiptPrompt"/></w:placeholder><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Entered answer</w:t></w:r></w:sdtContent></w:sdt>`
    ).replaceAll(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "http://purl.oclc.org/ooxml/wordprocessingml/main"
    ),
  });
  const orderedPlaceholderResult = await publishFixture(
    app,
    adminBearer,
    "nested-and-glossary-placeholder",
    orderedPlaceholderFixture
  );
  const orderedPlaceholderForm = await prisma.form.findUniqueOrThrow({
    select: { id: true },
    where: { publicId: orderedPlaceholderResult.publicId },
  });
  const orderedPlaceholderPublished =
    await prisma.publishedTemplate.findUniqueOrThrow({
      include: { manifest: { include: { fields: true } } },
      where: { formId: orderedPlaceholderForm.id },
    });
  expect(
    orderedPlaceholderPublished.manifest?.fields
      .toSorted((left, right) => left.position - right.position)
      .map(({ label, placeholder, position, tag }) => ({
        label,
        placeholder,
        position,
        tag,
      }))
  ).toEqual([
    { label: "Parent label", placeholder: null, position: 0, tag: "parent" },
    { label: "Child label", placeholder: null, position: 1, tag: "child" },
    {
      label: "Prompt label",
      placeholder: "Configured glossary placeholder",
      position: 2,
      tag: "prompt",
    },
  ]);
};

export interface InvalidTemplateContractsInput {
  app: ReturnType<typeof createApp>;
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  refreshedAdminEditor: PrimaryPublicationOutput["refreshedAdminEditor"];
  activeTemplateDocumentKey: PrimaryPublicationOutput["activeTemplateDocumentKey"];
  activePublishCapability: PrimaryPublicationOutput["activePublishCapability"];
}

export const runInvalidTemplateContracts = async (
  input: InvalidTemplateContractsInput
): Promise<void> => {
  const {
    app,
    adminBearer,
    formPublicId,
    refreshedAdminEditor,
    activeTemplateDocumentKey,
    activePublishCapability,
  } = input;
  const invalidFixtureCases = [
    {
      bytes: docxFixture("no-controls"),
      label: "no-controls",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "duplicate", type: "<w:text/>" }) +
            contentControl({ tag: "duplicate", type: "<w:text/>" })
        ),
      }),
      label: "duplicate-tags",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({
            tag: "malformed-options",
            type: `<w:comboBox><w:listItem w:displayText="Missing value"/></w:comboBox>`,
          })
        ),
      }),
      label: "malformed-options",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "", type: "<w:text/>" })
        ),
      }),
      label: "blank-tag",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({
            tag: "duplicate-options",
            type: `<w:comboBox><w:listItem w:displayText="One" w:value="same"/><w:listItem w:displayText="Two" w:value="same"/></w:comboBox>`,
          })
        ),
      }),
      label: "duplicate-options",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "unsupported", type: "<w:group/>" })
        ),
      }),
      label: "unsupported-group",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag w:val="nested-marker"/><w:placeholder><w:text/></w:placeholder></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "nested-marker",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag w:val="nested-option"/><w:comboBox/><w:placeholder><w:listItem w:displayText="Wrong parent" w:value="wrong"/></w:placeholder></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "nested-option",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag w:val="nested-option-child"/><w:comboBox><w:listItem w:displayText="One" w:value="one"><w:bogus/></w:listItem></w:comboBox></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "nested-option-child",
    },
    {
      bytes: docxXmlFixture({
        additionalParts: {
          "word/header1.xml": strToU8(
            `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${contentControl({ tag: "orphan-header", type: "<w:text/>" })}</w:hdr>`
          ),
        },
        document:
          '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
      }),
      label: "orphan-header-control",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "unknown", type: "<w:unknown/>" })
        ),
      }),
      label: "unknown-control",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "wrong-namespace", type: "<w14:picture/>" })
        ),
      }),
      label: "wrong-namespace",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag word:val="forged"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "wrong-attribute-namespace",
    },
  ];
  for (const fixture of invalidFixtureCases) {
    const result = await publishFixture(
      app,
      adminBearer,
      fixture.label,
      fixture.bytes
    );
    expect(result.operation.status).toBe("failed");
    const fixtureForm = await prisma.form.findUniqueOrThrow({
      select: { id: true },
      where: { publicId: result.publicId },
    });
    expect(
      await prisma.publishedTemplate.findUnique({
        where: { formId: fixtureForm.id },
      })
    ).toBeNull();
  }
  const publishedDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(publishedDeleteResponse.status).toBe(409);
  expect(await publishedDeleteResponse.json()).toMatchObject({
    error: "form_not_draft",
  });
  expect(
    await prisma.form.findUnique({ where: { publicId: formPublicId } })
  ).not.toBeNull();

  const immutableSaveCapability =
    refreshedAdminEditor.bridge.capabilities["save-template"];
  if (!immutableSaveCapability) {
    throw new Error("The immutable save capability was not returned");
  }
  const immutableSaveResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: activeTemplateDocumentKey }),
      headers: capabilityHeaders(immutableSaveCapability),
      method: "POST",
    })
  );
  expect(immutableSaveResponse.status).toBe(409);
  expect(await immutableSaveResponse.json()).toMatchObject({
    error: "published_immutable",
  });
  const immutablePublishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: activeTemplateDocumentKey }),
      headers: capabilityHeaders(activePublishCapability),
      method: "POST",
    })
  );
  expect(immutablePublishResponse.status).toBe(409);
  expect(await immutablePublishResponse.json()).toMatchObject({
    error: "published_immutable",
  });
};
