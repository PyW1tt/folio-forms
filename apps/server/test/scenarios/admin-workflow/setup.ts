// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { ensureBootstrapAdmin } from "@onlyoffice/auth";
import type { Form, TemplateDraft } from "@onlyoffice/db";
import { prisma } from "@onlyoffice/db";
import { zipSync, strToU8 } from "fflate";

import type { createApp } from "../../../src/app";
import { resolveCallbackDocumentUrl } from "../../../src/operations/callback";
import { readObject, objectExists } from "../../../src/storage";
import {
  docxFixture,
  sizedDocxFixture,
  maxTemplateUploadBytes,
  docxXmlFixture,
  templateContentTypesXml,
  strictDocxFixture,
  utf16DocxFixture,
} from "../../fixtures/documents";
import {
  jsonHeaders,
  formCreationRequest,
  createCredentialFixture,
  bearerFor,
} from "../../fixtures/http";
import type { EditorConfigBody } from "../../fixtures/http";

export interface BootstrapAndCreationInput {
  app: ReturnType<typeof createApp>;
}

export interface BootstrapAndCreationOutput {
  adminEmail: string;
  userEmail: string;
  password: string;
  otherUserEmail: string;
  adminId: string;
  adminBearer: string;
  secretFormTitle: string;
  secretFormDescription: string;
  formPublicId: string;
  createdFormRecord: Form & { templateDraft: TemplateDraft };
  formId: string;
  templateDocumentKey: string;
}

export const runBootstrapAndCreation = async (
  input: BootstrapAndCreationInput
): Promise<BootstrapAndCreationOutput> => {
  const { app } = input;
  const adminEmail = `ticket-02-admin-${crypto.randomUUID()}@example.com`;
  const userEmail = `ticket-02-user-${crypto.randomUUID()}@example.com`;
  const password = "Ticket02-password-for-test";
  const otherUserEmail = `ticket-03-other-${crypto.randomUUID()}@example.com`;
  const bootstrapEmail =
    process.env.BOOTSTRAP_ADMIN_EMAIL?.trim().toLowerCase();
  const bootstrapName = process.env.BOOTSTRAP_ADMIN_NAME?.trim();
  if (
    !bootstrapEmail ||
    !bootstrapName ||
    !process.env.BOOTSTRAP_ADMIN_PASSWORD
  ) {
    throw new Error(
      "The HTTP application test requires all BOOTSTRAP_ADMIN_* variables"
    );
  }
  expect(
    resolveCallbackDocumentUrl(
      "https://docs.example//169.254.169.254/latest/meta-data?key=value",
      new Set(["https://docs.example"]),
      "https://docs.example",
      "http://onlyoffice"
    )
  ).toBe("http://onlyoffice//169.254.169.254/latest/meta-data?key=value");
  const callbackDocumentUrl = new URL(
    "/onlyoffice/document/template.docx",
    process.env.ONLYOFFICE_DOCUMENT_BASE_URL ??
      "http://host.docker.internal:3000"
  ).toString();
  expect(resolveCallbackDocumentUrl(callbackDocumentUrl)).toBe(
    callbackDocumentUrl
  );
  expect(
    resolveCallbackDocumentUrl(
      "https://user:password@docs.example/document.docx",
      new Set(["https://docs.example"]),
      "https://docs.example",
      "http://onlyoffice"
    )
  ).toBeNull();
  const bootstrapUserSelect = {
    createdAt: true,
    email: true,
    enabled: true,
    id: true,
    mustChangePassword: true,
    name: true,
    role: true,
    updatedAt: true,
  } as const;

  const existingBootstrap = await prisma.user.findUnique({
    select: bootstrapUserSelect,
    where: { email: bootstrapEmail },
  });
  const existingAdmin = await prisma.user.findFirst({
    select: { id: true },
    where: { role: "admin" },
  });
  const existingBootstrapAccount = existingBootstrap
    ? await prisma.account.findFirst({
        where: { providerId: "credential", userId: existingBootstrap.id },
      })
    : null;
  const firstBootstrapResult = await ensureBootstrapAdmin();
  expect(firstBootstrapResult).toBe(!existingAdmin);
  const bootstrapBefore = await prisma.user.findUnique({
    select: bootstrapUserSelect,
    where: { email: bootstrapEmail },
  });
  const bootstrapAccountBefore = bootstrapBefore
    ? await prisma.account.findFirst({
        where: { providerId: "credential", userId: bootstrapBefore.id },
      })
    : null;
  if (existingBootstrap) {
    expect(bootstrapBefore).toEqual(existingBootstrap);
    expect(bootstrapAccountBefore).toEqual(existingBootstrapAccount);
  } else if (existingAdmin) {
    expect(bootstrapBefore).toBeNull();
  } else {
    if (!bootstrapBefore) {
      throw new Error("Bootstrap Admin was not created");
    }
    expect(bootstrapBefore).toMatchObject({
      email: bootstrapEmail,
      enabled: true,
      mustChangePassword: true,
      name: bootstrapName,
      role: "admin",
    });
    expect(bootstrapAccountBefore).toMatchObject({
      accountId: bootstrapBefore.id,
      issuer: "local:credential",
      password: expect.any(String),
      providerId: "credential",
      userId: bootstrapBefore.id,
    });
  }
  const adminCountBeforeSecondEnsure = await prisma.user.count({
    where: { role: "admin" },
  });

  const secondBootstrapResult = await ensureBootstrapAdmin();
  expect(secondBootstrapResult).toBe(false);
  expect(
    await prisma.user.count({
      where: { role: "admin" },
    })
  ).toBe(adminCountBeforeSecondEnsure);
  const bootstrapAfter = bootstrapBefore
    ? await prisma.user.findUnique({
        select: bootstrapUserSelect,
        where: { id: bootstrapBefore.id },
      })
    : null;
  const bootstrapAccountAfter = bootstrapBefore
    ? await prisma.account.findFirst({
        where: { providerId: "credential", userId: bootstrapBefore.id },
      })
    : null;
  expect(bootstrapAfter).toEqual(bootstrapBefore);
  expect(bootstrapAccountAfter).toEqual(bootstrapAccountBefore);

  const signupEmail = `ticket-04-signup-${crypto.randomUUID()}@example.com`;
  const signupResponse = await app.handle(
    new Request("http://test.local/api/auth/sign-up/email", {
      body: JSON.stringify({
        email: signupEmail,
        name: "Ticket 04 Direct Signup",
        password,
      }),
      headers: jsonHeaders,
      method: "POST",
    })
  );
  expect(signupResponse.status).toBe(404);
  expect(
    await prisma.user.findUnique({
      where: { email: signupEmail },
    })
  ).toBeNull();

  const healthResponse = await app.handle(
    new Request("http://test.local/health")
  );
  expect(healthResponse.status).toBe(200);
  expect(await healthResponse.json()).toEqual({ ok: true });

  const unauthorizedResponse = await app.handle(
    formCreationRequest({ source: "blank", title: "Denied Form" })
  );
  expect(unauthorizedResponse.status).toBe(401);
  expect(await unauthorizedResponse.json()).toMatchObject({
    error: "unauthorized",
  });

  const admin = await createCredentialFixture({
    email: adminEmail,
    name: "Ticket 04 Workflow Admin",
    password,
    role: "admin",
  });
  const adminId = admin.id;
  const adminBearer = await bearerFor(app, adminEmail, password);
  const secretFormTitle = `Ticket 04 secret title ${crypto.randomUUID()}`;
  const secretFormDescription = `Ticket 04 secret description ${crypto.randomUUID()}`;

  const jsonCreateResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      body: JSON.stringify({ source: "blank", title: "JSON is not accepted" }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${adminBearer}`,
      },
      method: "POST",
    })
  );
  expect(jsonCreateResponse.status).toBe(415);
  expect(await jsonCreateResponse.json()).toMatchObject({
    error: "invalid_file_type",
  });
  for (const invalidRequest of [
    formCreationRequest({
      authorization: adminBearer,
      title: "Missing source",
    }),
    formCreationRequest({
      authorization: adminBearer,
      source: "blank",
    }),
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      title: "Missing uploaded template",
    }),
    formCreationRequest({
      authorization: adminBearer,
      source: "blank",
      template: {
        bytes: docxFixture("unexpected-blank-template"),
        name: "unexpected.docx",
      },
      title: "Unexpected blank template",
    }),
  ]) {
    const invalidResponse = await app.handle(invalidRequest);
    expect(invalidResponse.status).toBe(400);
    expect(await invalidResponse.json()).toMatchObject({
      error: "invalid_request",
    });
  }
  const createResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      description: secretFormDescription,
      source: "blank",
      title: secretFormTitle,
    })
  );
  expect(createResponse.status).toBe(200);
  const createdForm = (await createResponse.json()) as {
    form?: {
      activeDraftCount?: number;
      hasTemplateDraft?: boolean;
      publicId?: string;
      status?: string;
      submissionCount?: number;
      title?: string;
    };
  };
  const formPublicId = createdForm.form?.publicId;
  expect(createdForm.form).toMatchObject({
    activeDraftCount: 0,
    hasTemplateDraft: true,
    status: "draft",
    submissionCount: 0,
    title: secretFormTitle,
  });
  if (!formPublicId) {
    throw new Error("The test form did not receive a public identifier");
  }
  expect(formPublicId).toMatch(/^[0-9a-f]{32}$/u);
  const createdFormRecord = await prisma.form.findUnique({
    include: { templateDraft: true },
    where: { publicId: formPublicId },
  });
  if (!createdFormRecord?.templateDraft) {
    throw new Error("The test form did not receive a template draft");
  }
  const formId = createdFormRecord.id;
  const templateDocumentKey = createdFormRecord.templateDraft.documentKey;
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: createdFormRecord.templateDraft.objectKey },
    })
  ).toBeNull();
  const serializedCreate = JSON.stringify(createdForm);
  expect(serializedCreate).not.toContain(formId);
  expect(serializedCreate).not.toContain(adminId);
  expect(serializedCreate).not.toContain(
    createdFormRecord.templateDraft.objectKey
  );
  expect(serializedCreate).not.toContain(templateDocumentKey);

  const listResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(listResponse.status).toBe(200);
  const listBody = (await listResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      status: string;
      submissionCount: number;
      title: string;
    }[];
  };
  const listedForm = listBody.forms?.find(
    (form) => form.publicId === formPublicId
  );
  expect(listedForm).toMatchObject({
    activeDraftCount: 0,
    publicId: formPublicId,
    status: "draft",
    submissionCount: 0,
    title: secretFormTitle,
  });
  for (const privateField of [
    "createdBy",
    "id",
    "objectKey",
    "publishedDocumentKey",
    "templateDocumentKey",
  ]) {
    expect(listedForm).not.toHaveProperty(privateField);
  }
  const serializedList = JSON.stringify(listBody);
  expect(serializedList).not.toContain(formId);
  expect(serializedList).not.toContain(adminId);
  expect(serializedList).not.toContain(
    createdFormRecord.templateDraft.objectKey
  );
  expect(serializedList).not.toContain(templateDocumentKey);
  const detailResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(detailResponse.status).toBe(200);
  const detailBody = (await detailResponse.json()) as {
    editorConfigUrl?: string;
    form?: { publicId?: string };
  };
  expect(detailBody).toMatchObject({
    editorConfigUrl: `/api/admin/forms/${formPublicId}/editor-config`,
    form: { publicId: formPublicId },
  });
  for (const privateField of [
    "createdBy",
    "id",
    "objectKey",
    "publishedDocumentKey",
    "templateDocumentKey",
  ]) {
    expect(detailBody.form).not.toHaveProperty(privateField);
  }
  const serializedDetail = JSON.stringify(detailBody);
  expect(serializedDetail).not.toContain(formId);
  expect(serializedDetail).not.toContain(adminId);
  expect(serializedDetail).not.toContain(
    createdFormRecord.templateDraft.objectKey
  );
  expect(serializedDetail).not.toContain(templateDocumentKey);
  return {
    adminBearer,
    adminEmail,
    adminId,
    createdFormRecord: {
      ...createdFormRecord,
      templateDraft: createdFormRecord.templateDraft,
    },
    formId,
    formPublicId,
    otherUserEmail,
    password,
    secretFormDescription,
    secretFormTitle,
    templateDocumentKey,
    userEmail,
  };
};

export interface TemplateUploadsInput {
  app: ReturnType<typeof createApp>;
  password: BootstrapAndCreationOutput["password"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  adminEmail: BootstrapAndCreationOutput["adminEmail"];
}

export interface TemplateUploadsOutput {
  competingAdmin: { email: string; id: string };
  competingAdminBearer: string;
  uploadPublicId: string;
}

export const runTemplateUploads = async (
  input: TemplateUploadsInput
): Promise<TemplateUploadsOutput> => {
  const { app, password, adminBearer, adminEmail } = input;
  const competingAdminEmail = `ticket-08-competing-${crypto.randomUUID()}@example.com`;
  const competingAdmin = await createCredentialFixture({
    email: competingAdminEmail,
    name: "Ticket 08 Competing Admin",
    password,
    role: "admin",
  });
  const competingAdminBearer = await bearerFor(
    app,
    competingAdminEmail,
    password,
    `ticket-08-competing-admin-${crypto.randomUUID()}`
  );
  const formCountBeforeUploadChecks = await prisma.form.count();
  const uploadBytes = docxFixture(`ticket-08-upload-${crypto.randomUUID()}`);
  const uploadTitle = `Ticket 08 upload ${crypto.randomUUID()}`;
  const uploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes: uploadBytes, name: "template.docx" },
      title: uploadTitle,
    })
  );
  expect(uploadResponse.status).toBe(200);
  const uploadBody = (await uploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const uploadPublicId = uploadBody.form?.publicId;
  const uploadRecord = uploadPublicId
    ? await prisma.form.findUnique({
        include: { templateDraft: true },
        where: { publicId: uploadPublicId },
      })
    : null;
  if (!uploadPublicId || !uploadRecord?.templateDraft) {
    throw new Error("The uploaded Template Draft was not created");
  }
  const uploadObjectKey = uploadRecord.templateDraft.objectKey;
  const uploadEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${uploadPublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(uploadEditorResponse.status).toBe(200);
  const uploadEditor = (await uploadEditorResponse.json()) as EditorConfigBody;
  const uploadLeaseId = uploadEditor.bridge.lease.id;
  expect(await readObject(uploadObjectKey)).toEqual(uploadBytes);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: uploadObjectKey },
    })
  ).toBeNull();
  const sameAdminOtherBearer = await bearerFor(
    app,
    adminEmail,
    password,
    `ticket-08-same-admin-other-session-${crypto.randomUUID()}`
  );
  const sameAdminOtherSessionDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${uploadPublicId}`, {
      headers: { Authorization: `Bearer ${sameAdminOtherBearer}` },
      method: "DELETE",
    })
  );
  expect(sameAdminOtherSessionDeleteResponse.status).toBe(409);
  expect(await sameAdminOtherSessionDeleteResponse.json()).toMatchObject({
    error: "editor_in_use",
  });
  const competingDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${uploadPublicId}`, {
      headers: { Authorization: `Bearer ${competingAdminBearer}` },
      method: "DELETE",
    })
  );
  expect(competingDeleteResponse.status).toBe(409);
  expect(await competingDeleteResponse.json()).toMatchObject({
    error: "editor_in_use",
  });
  expect(
    await prisma.form.findUnique({ where: { publicId: uploadPublicId } })
  ).not.toBeNull();
  expect(await objectExists(uploadObjectKey)).toBe(true);

  const maximumUploadBytes = sizedDocxFixture(
    `ticket-08-limit-${crypto.randomUUID()}`,
    maxTemplateUploadBytes
  );
  const maximumUploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: maximumUploadBytes,
        name: "maximum-size.docx",
        type: "application/octet-stream",
      },
      title: `Ticket 08 maximum upload ${crypto.randomUUID()}`,
    })
  );
  expect(maximumUploadResponse.status).toBe(200);
  const maximumUploadBody = (await maximumUploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const maximumUploadPublicId = maximumUploadBody.form?.publicId;
  expect(maximumUploadPublicId).toBeTruthy();

  const oversizedUploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: new Uint8Array(maxTemplateUploadBytes + 1),
        name: "too-large.docx",
      },
      title: "Ticket 08 oversized upload",
    })
  );
  expect(oversizedUploadResponse.status).toBe(413);
  expect(await oversizedUploadResponse.json()).toMatchObject({
    error: "payload_too_large",
  });
  const nonDocxResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxFixture("wrong-extension"),
        name: "template.pdf",
      },
      title: "Ticket 08 wrong upload type",
    })
  );
  expect(nonDocxResponse.status).toBe(415);
  expect(await nonDocxResponse.json()).toMatchObject({
    error: "invalid_file_type",
  });
  const malformedDocxResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: new TextEncoder().encode("not a ZIP package"),
        name: "malformed.docx",
      },
      title: "Ticket 08 malformed upload",
    })
  );
  expect(malformedDocxResponse.status).toBe(422);
  expect(await malformedDocxResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const incompleteDocxResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: zipSync({
          "word/document.xml": strToU8(
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
          ),
        }),
        name: "incomplete.docx",
      },
      title: "Ticket 08 incomplete upload",
    })
  );
  expect(incompleteDocxResponse.status).toBe(422);
  expect(await incompleteDocxResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const externalRelationshipResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          additionalParts: {
            "word/_rels/document.xml.rels": strToU8(
              '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="external" Target="http://127.0.0.1:80/" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/></Relationships>'
            ),
          },
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
        }),
        name: "external-relationship.docx",
      },
      title: "Ticket 10 external relationship upload",
    })
  );
  expect(externalRelationshipResponse.status).toBe(422);
  expect(await externalRelationshipResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const malformedXmlResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document><',
        }),
        name: "malformed-xml.docx",
      },
      title: "Ticket 08 malformed XML upload",
    })
  );
  expect(malformedXmlResponse.status).toBe(422);
  expect(await malformedXmlResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const doctypeResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          document:
            '<?xml version="1.0"?><!DOCTYPE word:document [<!ENTITY injected "value">]><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body><word:p>&injected;</word:p></word:body></word:document>',
        }),
        name: "doctype.docx",
      },
      title: "Ticket 08 XML doctype upload",
    })
  );
  expect(doctypeResponse.status).toBe(422);
  expect(await doctypeResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const declaredSecondaryXmlResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          additionalParts: {
            "word/_rels/document.xml.rels": strToU8(
              '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Target="header.bin" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"/></Relationships>'
            ),
            "word/header.bin": strToU8(
              '<!DOCTYPE w:hdr [<!ENTITY injected "value">]><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">&injected;</w:hdr>'
            ),
          },
          contentTypes: templateContentTypesXml.replace(
            "</Types>",
            '<Override PartName="/word/header.bin" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
          ),
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
        }),
        name: "declared-secondary-xml.docx",
      },
      title: "Ticket 08 declared secondary XML upload",
    })
  );
  expect(declaredSecondaryXmlResponse.status).toBe(422);
  expect(await declaredSecondaryXmlResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const strictUploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: strictDocxFixture(`ticket-08-strict-${crypto.randomUUID()}`),
        name: "strict.docx",
      },
      title: "Ticket 08 Strict OOXML upload",
    })
  );
  const vmlDoctypeResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          additionalParts: {
            "word/_rels/document.xml.rels": strToU8(
              '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="vml" Target="drawings/vmlDrawing1.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/></Relationships>'
            ),
            "word/drawings/vmlDrawing1.vml": strToU8(
              '<!DOCTYPE xml [<!ENTITY injected "value">]><xml xmlns:v="urn:schemas-microsoft-com:vml">&injected;</xml>'
            ),
          },
          contentTypes: templateContentTypesXml.replace(
            "</Types>",
            '<Override PartName="/word/drawings/vmlDrawing1.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>'
          ),
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
        }),
        name: "vml-doctype.docx",
      },
      title: "Ticket 08 VML doctype upload",
    })
  );
  expect(vmlDoctypeResponse.status).toBe(422);
  expect(await vmlDoctypeResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  expect(strictUploadResponse.status).toBe(200);
  const strictUploadBody = (await strictUploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const strictUploadPublicId = strictUploadBody.form?.publicId;
  if (!strictUploadPublicId) {
    throw new Error("The Strict OOXML Template Draft was not created");
  }
  const strictDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${strictUploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(strictDeleteResponse.status).toBe(200);
  const utf16UploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: utf16DocxFixture(),
        name: "utf16.docx",
      },
      title: "Ticket 08 UTF-16 OOXML upload",
    })
  );
  expect(utf16UploadResponse.status).toBe(200);
  const utf16UploadBody = (await utf16UploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const utf16UploadPublicId = utf16UploadBody.form?.publicId;
  if (!utf16UploadPublicId) {
    throw new Error("The UTF-16 OOXML Template Draft was not created");
  }
  const utf16DeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${utf16UploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(utf16DeleteResponse.status).toBe(200);
  expect(await prisma.form.count()).toBe(formCountBeforeUploadChecks + 2);

  const deleteUploadResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${uploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(deleteUploadResponse.status).toBe(200);
  expect(await deleteUploadResponse.json()).toEqual({ deleted: true });
  expect(
    await prisma.form.findUnique({ where: { publicId: uploadPublicId } })
  ).toBeNull();
  expect(
    await prisma.editorLease.findUnique({ where: { id: uploadLeaseId } })
  ).toBeNull();
  expect(await objectExists(uploadObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: uploadObjectKey },
    })
  ).toBeNull();

  if (!maximumUploadPublicId) {
    throw new Error("The maximum-sized Template Draft was not created");
  }
  const maximumUploadRecord = await prisma.form.findUnique({
    include: { templateDraft: true },
    where: { publicId: maximumUploadPublicId },
  });
  if (!maximumUploadRecord?.templateDraft) {
    throw new Error("The maximum-sized Template Draft was not persisted");
  }
  const maximumUploadObjectKey = maximumUploadRecord.templateDraft.objectKey;
  const deleteMaximumUploadResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${maximumUploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(deleteMaximumUploadResponse.status).toBe(200);
  expect(await objectExists(maximumUploadObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: maximumUploadObjectKey },
    })
  ).toBeNull();
  return {
    competingAdmin,
    competingAdminBearer,
    uploadPublicId,
  };
};
