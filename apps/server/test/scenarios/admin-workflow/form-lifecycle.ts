// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";

import type { createApp } from "../../../src/app";
import { readObject, objectExists } from "../../../src/storage";
import { docxFixture } from "../../fixtures/documents";
import {
  jsonHeaders,
  formCreationRequest,
  createCredentialFixture,
  bearerFor,
} from "../../fixtures/http";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";
import type { PrimaryPublicationOutput } from "./template-contract";

export interface FormLifecycleInput {
  app: ReturnType<typeof createApp>;
  formRecord: PrefillEntryOutput["formRecord"];
  userBearer: PrefillEntryOutput["userBearer"];
  formId: BootstrapAndCreationOutput["formId"];
  user: PrefillEntryOutput["user"];
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  secretFormTitle: BootstrapAndCreationOutput["secretFormTitle"];
  secretFormDescription: BootstrapAndCreationOutput["secretFormDescription"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  publishedBytes: PrimaryPublicationOutput["publishedBytes"];
  password: BootstrapAndCreationOutput["password"];
  adminId: BootstrapAndCreationOutput["adminId"];
}

export interface FormLifecycleOutput {
  publishedContractBefore: PrimaryPublicationOutput["publishedManifestRecord"];
  archivedNoResponseEmail: string;
  archivedNoResponseBearer: string;
}

export const runFormLifecycle = async (
  input: FormLifecycleInput
): Promise<FormLifecycleOutput> => {
  const {
    app,
    formRecord,
    userBearer,
    formId,
    user,
    formPublicId,
    secretFormTitle,
    secretFormDescription,
    adminBearer,
    publishedBytes,
    password,
    adminId,
  } = input;
  const concurrentStarts = await Promise.all(
    [0, 1].map(() =>
      app.handle(
        new Request(
          `http://test.local/api/forms/${formRecord.publicId}/start`,
          {
            headers: { Authorization: `Bearer ${userBearer}` },
            method: "POST",
          }
        )
      )
    )
  );
  expect(concurrentStarts.every((response) => response.status === 200)).toBe(
    true
  );
  const concurrentStartBodies = (await Promise.all(
    concurrentStarts.map((response) => response.json())
  )) as { response?: { id?: string } }[];
  expect(concurrentStartBodies[0]).toMatchObject({
    response: { id: expect.any(String) },
  });
  expect(concurrentStartBodies[1]).toMatchObject({
    response: { id: concurrentStartBodies[0]?.response?.id },
  });
  expect(
    await prisma.response.count({ where: { formId, userId: user.id } })
  ).toBe(1);
  const existingResponseId = concurrentStartBodies[0]?.response?.id;
  if (!existingResponseId) {
    throw new Error("The concurrent response was not created");
  }
  const publishedContractBefore = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
    },
    where: { formId },
  });
  if (!publishedContractBefore?.manifest) {
    throw new Error("The Ticket 11 published contract was not found");
  }
  const unauthorizedMetadataUpdate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        description: "must not change",
        title: "must not change",
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "PATCH",
    })
  );
  expect(unauthorizedMetadataUpdate.status).toBe(403);
  const updatedTitle = `${secretFormTitle} metadata`;
  const updatedDescription = `${secretFormDescription} metadata`;
  const metadataUpdate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        description: updatedDescription,
        title: updatedTitle,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(metadataUpdate.status).toBe(200);
  expect(await metadataUpdate.json()).toMatchObject({
    form: {
      description: updatedDescription,
      publicId: formPublicId,
      status: "published",
      title: updatedTitle,
      version: 1,
    },
  });
  const publicMetadataAfterUpdate = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(await publicMetadataAfterUpdate.json()).toMatchObject({
    form: { description: updatedDescription, title: updatedTitle },
  });
  const publishedContractAfter = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
    },
    where: { formId },
  });
  if (!publishedContractAfter?.manifest) {
    throw new Error("The Ticket 11 published contract was removed");
  }
  expect(publishedContractAfter).toMatchObject({
    contentHash: publishedContractBefore.contentHash,
    documentKey: publishedContractBefore.documentKey,
    id: publishedContractBefore.id,
    objectKey: publishedContractBefore.objectKey,
    version: publishedContractBefore.version,
  });
  expect(publishedContractAfter.manifest).toMatchObject({
    configurationHash: publishedContractBefore.manifest.configurationHash,
    id: publishedContractBefore.manifest.id,
  });
  expect(publishedContractAfter.manifest.fields).toEqual(
    publishedContractBefore.manifest.fields
  );
  expect(publishedContractAfter.prefillConfiguration).toEqual(
    publishedContractBefore.prefillConfiguration
  );
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: { action: "update_form_metadata", targetId: formPublicId },
    })
  ).toMatchObject({ outcome: "success", targetId: formPublicId });
  const restoreMetadata = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        description: secretFormDescription,
        title: secretFormTitle,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(restoreMetadata.status).toBe(200);

  const unauthorizedDuplicate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/duplicate`, {
      body: "{}",
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(unauthorizedDuplicate.status).toBe(403);
  const duplicatePublishedResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/duplicate`, {
      body: "{}",
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(duplicatePublishedResponse.status).toBe(200);
  const duplicatePublishedBody = (await duplicatePublishedResponse.json()) as {
    form?: {
      description?: string;
      publicId?: string;
      status?: string;
      title?: string;
      version?: number;
    };
  };
  const duplicatePublishedPublicId = duplicatePublishedBody.form?.publicId;
  if (!duplicatePublishedPublicId) {
    throw new Error("The published Form duplicate was not created");
  }
  expect(duplicatePublishedBody.form).toMatchObject({
    description: secretFormDescription,
    status: "draft",
    title: secretFormTitle,
    version: 0,
  });
  expect(duplicatePublishedPublicId).not.toBe(formPublicId);
  expect(duplicatePublishedPublicId).toMatch(/^[0-9a-f]{32}$/u);
  const duplicatePublished = await prisma.form.findUnique({
    include: {
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
      templateDraft: { include: { fieldRules: { orderBy: { tag: "asc" } } } },
    },
    where: { publicId: duplicatePublishedPublicId },
  });
  if (!duplicatePublished?.templateDraft) {
    throw new Error("The published Form duplicate has no Template Draft");
  }
  expect(
    Buffer.from(await readObject(duplicatePublished.templateDraft.objectKey))
  ).toEqual(Buffer.from(publishedBytes));
  const sourcePrefillByTag = new Map(
    (publishedContractBefore.prefillConfiguration?.fields ?? []).map(
      (field) => [field.tag, field]
    )
  );
  expect(
    duplicatePublished.templateDraft.fieldRules.map((field) => ({
      prefillPointer: field.prefillPointer,
      prefillPolicy: field.prefillPolicy,
      required: field.required,
      tag: field.tag,
    }))
  ).toEqual(
    publishedContractBefore.manifest.fields
      .map((field) => ({
        prefillPointer: sourcePrefillByTag.get(field.tag)?.pointer ?? null,
        prefillPolicy: field.prefillPolicy,
        required: field.required,
        tag: field.tag,
      }))
      .toSorted((left, right) => left.tag.localeCompare(right.tag))
  );
  expect(duplicatePublished.prefillConfiguration).toBeNull();
  const [
    duplicateResponseCount,
    duplicateSubmissionCount,
    duplicateOperationCount,
    duplicateLeaseCount,
    duplicateAuditEvents,
  ] = await Promise.all([
    prisma.response.count({ where: { formId: duplicatePublished.id } }),
    prisma.submission.count({ where: { formId: duplicatePublished.id } }),
    prisma.operation.count({ where: { formId: duplicatePublished.id } }),
    prisma.editorLease.count({
      where: {
        targetId: duplicatePublished.templateDraft.id,
        targetType: "template_draft",
      },
    }),
    prisma.auditEvent.findMany({
      orderBy: { createdAt: "asc" },
      select: { action: true, outcome: true },
      where: { targetId: duplicatePublishedPublicId },
    }),
  ]);
  expect(duplicateResponseCount).toBe(0);
  expect(duplicateSubmissionCount).toBe(0);
  expect(duplicateOperationCount).toBe(0);
  expect(duplicateLeaseCount).toBe(0);
  expect(duplicateAuditEvents).toEqual([
    { action: "duplicate_form", outcome: "success" },
  ]);
  const duplicateTitle = `${secretFormTitle} duplicate`;
  const duplicateMetadataUpdate = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        body: JSON.stringify({
          description: "Independent duplicate",
          title: duplicateTitle,
        }),
        headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
        method: "PATCH",
      }
    )
  );
  expect(duplicateMetadataUpdate.status).toBe(200);
  const sourceAfterDuplicate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(await sourceAfterDuplicate.json()).toMatchObject({
    form: { description: secretFormDescription, title: secretFormTitle },
  });
  const duplicatePublishedAfterUpdate = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(await duplicatePublishedAfterUpdate.json()).toMatchObject({
    form: {
      description: "Independent duplicate",
      publicId: duplicatePublishedPublicId,
      status: "draft",
      title: duplicateTitle,
    },
  });
  const duplicatePublishedObjectKey =
    duplicatePublished.templateDraft.objectKey;
  const deletePublishedDuplicate = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(deletePublishedDuplicate.status).toBe(200);
  expect(await objectExists(duplicatePublishedObjectKey)).toBe(false);
  expect(await objectExists(publishedContractBefore.objectKey)).toBe(true);
  const duplicateDeleteLookupResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(duplicateDeleteLookupResponse.status).toBe(404);

  const draftDuplicateSourceResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      description: "Ticket 11 draft source",
      source: "blank",
      title: "Ticket 11 draft source",
    })
  );
  expect(draftDuplicateSourceResponse.status).toBe(200);
  const draftDuplicateSourceBody =
    (await draftDuplicateSourceResponse.json()) as {
      form?: { publicId?: string };
    };
  const draftDuplicateSourcePublicId = draftDuplicateSourceBody.form?.publicId;
  if (!draftDuplicateSourcePublicId) {
    throw new Error("The draft duplicate source was not created");
  }
  const draftDuplicateSource = await prisma.form.findUniqueOrThrow({
    include: { templateDraft: true },
    where: { publicId: draftDuplicateSourcePublicId },
  });
  if (!draftDuplicateSource.templateDraft) {
    throw new Error("The draft duplicate source has no Template Draft");
  }
  const draftDuplicateResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${draftDuplicateSourcePublicId}/duplicate`,
      {
        body: "{}",
        headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      }
    )
  );
  expect(draftDuplicateResponse.status).toBe(200);
  const draftDuplicateBody = (await draftDuplicateResponse.json()) as {
    form?: { publicId?: string; status?: string; title?: string };
  };
  const draftDuplicatePublicId = draftDuplicateBody.form?.publicId;
  if (!draftDuplicatePublicId) {
    throw new Error("The draft Form duplicate was not created");
  }
  expect(draftDuplicateBody.form).toMatchObject({
    status: "draft",
    title: "Ticket 11 draft source",
  });
  const draftDuplicate = await prisma.form.findUniqueOrThrow({
    include: { templateDraft: true },
    where: { publicId: draftDuplicatePublicId },
  });
  if (!draftDuplicate.templateDraft) {
    throw new Error("The draft Form duplicate has no Template Draft");
  }
  expect(
    Buffer.from(await readObject(draftDuplicate.templateDraft.objectKey))
  ).toEqual(
    Buffer.from(await readObject(draftDuplicateSource.templateDraft.objectKey))
  );
  const draftDuplicateObjectKey = draftDuplicate.templateDraft.objectKey;
  const draftDuplicateDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${draftDuplicatePublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(draftDuplicateDeleteResponse.status).toBe(200);
  expect(await objectExists(draftDuplicateObjectKey)).toBe(false);
  const draftDuplicateSourceDeleteResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${draftDuplicateSourcePublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(draftDuplicateSourceDeleteResponse.status).toBe(200);
  const draftMetadataCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      description: "Ticket 10 draft metadata secret",
      source: "upload",
      template: {
        bytes: docxFixture("draft-metadata"),
        name: "draft-metadata.docx",
      },
      title: "Ticket 10 draft metadata secret",
    })
  );
  expect(draftMetadataCreateResponse.status).toBe(200);
  const draftMetadataCreateBody =
    (await draftMetadataCreateResponse.json()) as {
      form?: { publicId?: string };
    };
  const draftMetadataPublicId = draftMetadataCreateBody.form?.publicId;
  if (!draftMetadataPublicId) {
    throw new Error("The draft metadata fixture was not created");
  }
  const draftMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${draftMetadataPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftMetadataResponse.status).toBe(404);
  const draftMetadataBody = await draftMetadataResponse.json();
  expect(JSON.stringify(draftMetadataBody)).not.toContain(
    "Ticket 10 draft metadata secret"
  );
  const draftMetadataDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${draftMetadataPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(draftMetadataDeleteResponse.status).toBe(200);
  const unauthenticatedMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}`)
  );
  expect(unauthenticatedMetadataResponse.status).toBe(401);
  const unauthenticatedMetadataBody =
    (await unauthenticatedMetadataResponse.json()) as Record<string, unknown>;
  expect(JSON.stringify(unauthenticatedMetadataBody)).not.toContain(
    secretFormTitle
  );
  expect(JSON.stringify(unauthenticatedMetadataBody)).not.toContain(
    secretFormDescription
  );
  expect(JSON.stringify(unauthenticatedMetadataBody)).not.toContain(
    "published"
  );
  const authenticatedMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(authenticatedMetadataResponse.status).toBe(200);
  expect(await authenticatedMetadataResponse.json()).toMatchObject({
    form: {
      description: secretFormDescription,
      title: secretFormTitle,
    },
  });
  const archivedNoResponseEmail = `ticket-14-archived-new-${crypto.randomUUID()}@example.com`;
  await createCredentialFixture({
    email: archivedNoResponseEmail,
    name: "Ticket 14 Archived New User",
    password,
  });
  const archivedNoResponseBearer = await bearerFor(
    app,
    archivedNoResponseEmail,
    password
  );
  const mixedLifecycleUpdate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        status: "archived",
        title: "must not mix lifecycle and metadata",
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(mixedLifecycleUpdate.status).toBe(400);
  expect(await mixedLifecycleUpdate.json()).toMatchObject({
    error: "invalid_request",
  });
  const unauthorizedArchive = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({ status: "archived" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "PATCH",
    })
  );
  expect(unauthorizedArchive.status).toBe(403);
  const [archiveResponse, archivedExistingStart] = await Promise.all([
    app.handle(
      new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
        body: JSON.stringify({ status: "archived" }),
        headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
        method: "PATCH",
      })
    ),
    app.handle(
      new Request(`http://test.local/api/forms/${formPublicId}/start`, {
        headers: { Authorization: `Bearer ${userBearer}` },
        method: "POST",
      })
    ),
  ]);
  expect(archiveResponse.status).toBe(200);
  expect(await archiveResponse.json()).toMatchObject({
    form: {
      publicId: formPublicId,
      status: "archived",
      version: 1,
    },
  });
  expect(archivedExistingStart.status).toBe(200);
  expect(await archivedExistingStart.json()).toMatchObject({
    response: { id: existingResponseId, status: "draft" },
  });
  const archivedNewMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
    })
  );
  expect(archivedNewMetadataResponse.status).toBe(404);
  expect(
    JSON.stringify(await archivedNewMetadataResponse.json())
  ).not.toContain(secretFormTitle);
  const archivedNewStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}/start`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "POST",
    })
  );
  expect(archivedNewStartResponse.status).toBe(409);
  expect(await archivedNewStartResponse.json()).toMatchObject({
    error: "form_unavailable",
  });
  const archivedExistingMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(archivedExistingMetadataResponse.status).toBe(200);
  expect(await archivedExistingMetadataResponse.json()).toMatchObject({
    form: {
      publicId: formPublicId,
      title: secretFormTitle,
    },
  });
  const archivedFormListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(archivedFormListResponse.status).toBe(200);
  const archivedFormListBody = (await archivedFormListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      status: string;
      submissionCount: number;
    }[];
  };
  expect(
    archivedFormListBody.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({
    activeDraftCount: 1,
    publicId: formPublicId,
    status: "archived",
    submissionCount: 0,
  });
  const archivedContract = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
    },
    where: { formId },
  });
  expect(archivedContract).toEqual(publishedContractBefore);
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "archive_form",
        outcome: "success",
        targetId: formPublicId,
      },
    })
  ).toMatchObject({ actorId: adminId, targetId: formPublicId });
  const userDeleteFormResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(userDeleteFormResponse.status).toBe(403);
  return {
    archivedNoResponseBearer,
    archivedNoResponseEmail,
    publishedContractBefore: {
      ...publishedContractBefore,
      manifest: publishedContractBefore.manifest,
    },
  };
};
