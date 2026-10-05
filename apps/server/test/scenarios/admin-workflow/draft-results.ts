import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";

import type { createApp } from "../../../src/app";
import { readObject, DOCX_CONTENT_TYPE } from "../../../src/storage";
import type { DraftLifecycleOutput } from "./draft";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";
import type { PrimaryPublicationOutput } from "./template-contract";

export interface DraftResultsAndExportsInput {
  app: ReturnType<typeof createApp>;
  formId: BootstrapAndCreationOutput["formId"];
  publishedManifestRecord: PrimaryPublicationOutput["publishedManifestRecord"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  formRecord: PrefillEntryOutput["formRecord"];
  userEmail: BootstrapAndCreationOutput["userEmail"];
  userBearer: PrefillEntryOutput["userBearer"];
  responseId: DraftLifecycleOutput["responseId"];
  savedDraftData: DraftLifecycleOutput["savedDraftData"];
  savedResponseDocumentKey: DraftLifecycleOutput["savedResponseDocumentKey"];
  savedResponseObjectKey: DraftLifecycleOutput["savedResponseObjectKey"];
  savedResponseDocumentBytes: DraftLifecycleOutput["savedResponseDocumentBytes"];
  adminId: BootstrapAndCreationOutput["adminId"];
}

export const runDraftResultsAndExports = async (
  input: DraftResultsAndExportsInput
): Promise<void> => {
  const {
    app,
    formId,
    publishedManifestRecord,
    adminBearer,
    formRecord,
    userEmail,
    userBearer,
    responseId,
    savedDraftData,
    savedResponseDocumentKey,
    savedResponseObjectKey,
    savedResponseDocumentBytes,
    adminId,
  } = input;
  const paginationUsers = Array.from({ length: 26 }, (_, index) => ({
    email: `ticket20-pagination-${String(index).padStart(2, "0")}@example.com`,
    id: `ticket20-pagination-${String(index).padStart(2, "0")}`,
    name: `Ticket 20 pagination ${index}`,
  }));
  const [firstPaginationUser] = paginationUsers;
  if (!firstPaginationUser) {
    throw new Error("The pagination fixture user was not created");
  }
  await prisma.user.createMany({ data: paginationUsers });
  await prisma.response.createMany({
    data: paginationUsers.map((paginationUser, index) => ({
      createdAt: new Date(2020, 0, index + 1),
      draftDocumentKey: `responses/pagination-${index}/draft/document.docx`,
      draftObjectKey: `responses/pagination-${index}/draft/object.docx`,
      formId,
      publishedTemplateId: publishedManifestRecord.id,
      publishedVersion: publishedManifestRecord.version,
      updatedAt: new Date(2020, 0, index + 1),
      userId: paginationUser.id,
    })),
  });
  const paginatedResultsFirstResponse = await app.handle(
    new Request("http://test.local/api/admin/results", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(paginatedResultsFirstResponse.status).toBe(200);
  const paginatedResultsFirstBody =
    (await paginatedResultsFirstResponse.json()) as {
      nextCursor: string | null;
      results: { id: string }[];
    };
  expect(paginatedResultsFirstBody.results).toHaveLength(25);
  expect(paginatedResultsFirstBody.nextCursor).toEqual(expect.any(String));
  const paginatedResultsSecondResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?cursor=${encodeURIComponent(paginatedResultsFirstBody.nextCursor as string)}`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(paginatedResultsSecondResponse.status).toBe(200);
  const paginatedResultsSecondBody =
    (await paginatedResultsSecondResponse.json()) as {
      nextCursor: string | null;
      results: { id: string }[];
    };
  expect(paginatedResultsSecondBody.results.length).toBeGreaterThan(0);
  expect(
    new Set([
      ...paginatedResultsFirstBody.results.map((result) => result.id),
      ...paginatedResultsSecondBody.results.map((result) => result.id),
    ]).size
  ).toBe(
    paginatedResultsFirstBody.results.length +
      paginatedResultsSecondBody.results.length
  );
  const filteredPaginationResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&user=${encodeURIComponent(firstPaginationUser.email)}&state=draft&correction=0`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(filteredPaginationResponse.status).toBe(200);
  expect(await filteredPaginationResponse.json()).toMatchObject({
    nextCursor: null,
    results: [
      {
        formPublicId: formRecord.publicId,
        state: "draft",
        userEmail: firstPaginationUser.email,
      },
    ],
  });
  const excludedPaginationResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&user=${encodeURIComponent(firstPaginationUser.email)}&state=submitted`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(excludedPaginationResponse.status).toBe(200);
  expect(await excludedPaginationResponse.json()).toEqual({
    nextCursor: null,
    results: [],
  });
  const adminDraftResultsResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&user=${encodeURIComponent(userEmail)}&state=draft&correction=0&from=2020-01-01T00:00:00.000Z&to=2030-01-01T00:00:00.000Z`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminDraftResultsResponse.status).toBe(200);
  const adminDraftResultsBody = (await adminDraftResultsResponse.json()) as {
    nextCursor: string | null;
    results: Record<string, unknown>[];
  };
  const unsavedDraftResponse = await prisma.response.findFirstOrThrow({
    select: { id: true },
    where: { userId: firstPaginationUser.id },
  });
  const unsavedDraftDetailResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${unsavedDraftResponse.id}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(unsavedDraftDetailResponse.status).toBe(200);
  expect(await unsavedDraftDetailResponse.json()).toMatchObject({
    result: { document: { available: false, state: "draft" }, state: "draft" },
  });
  const forbiddenUnsavedDraftDetailResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${unsavedDraftResponse.id}`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(forbiddenUnsavedDraftDetailResponse.status).toBe(403);
  const unavailableDraftViewerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${unsavedDraftResponse.id}/viewer-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(unavailableDraftViewerResponse.status).toBe(409);
  expect(await unavailableDraftViewerResponse.json()).toMatchObject({
    error: "document_unavailable",
  });
  const forbiddenUnsavedDraftViewerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${unsavedDraftResponse.id}/viewer-config`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(forbiddenUnsavedDraftViewerResponse.status).toBe(403);
  await prisma.response.deleteMany({
    where: {
      userId: {
        in: paginationUsers.map((paginationUser) => paginationUser.id),
      },
    },
  });
  await prisma.user.deleteMany({
    where: {
      id: { in: paginationUsers.map((paginationUser) => paginationUser.id) },
    },
  });
  expect(adminDraftResultsBody).toMatchObject({
    nextCursor: null,
    results: [
      {
        formPublicId: formRecord.publicId,
        id: responseId,
        latestCorrectionNumber: null,
        state: "draft",
        userEmail,
      },
    ],
  });
  const adminDraftDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/results/${responseId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminDraftDetailResponse.status).toBe(200);
  const forbiddenSavedDraftDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/results/${responseId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(forbiddenSavedDraftDetailResponse.status).toBe(403);
  const adminDraftDetailBody = (await adminDraftDetailResponse.json()) as {
    result: {
      fields: {
        label: string;
        options: { displayText: string; value: string }[];
        placeholder: string | null;
        position: number;
        tag: string;
        type: string;
      }[];
    } & Record<string, unknown>;
  };
  const expectedAdminDraftFields = [...publishedManifestRecord.manifest.fields]
    .toSorted((left, right) => left.position - right.position)
    .map(({ label, options, placeholder, position, tag, type }) => {
      const receiptOptions = Array.isArray(options)
        ? options.flatMap((option) => {
            if (
              !option ||
              typeof option !== "object" ||
              Array.isArray(option)
            ) {
              return [];
            }
            const optionRecord = option as Record<string, unknown>;
            const { displayText } = optionRecord;
            const { value } = optionRecord;
            return typeof displayText === "string" && typeof value === "string"
              ? [{ displayText, value }]
              : [];
          })
        : [];
      return {
        label,
        options: receiptOptions,
        placeholder,
        position,
        tag,
        type,
      };
    });
  expect(adminDraftDetailBody.result).toMatchObject({
    data: savedDraftData,
    document: { available: true, state: "draft" },
    id: responseId,
    state: "draft",
  });
  expect(adminDraftDetailBody.result.fields).toEqual(expectedAdminDraftFields);
  const adminDraftViewerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminDraftViewerResponse.status).toBe(200);
  expect(await adminDraftViewerResponse.json()).toMatchObject({
    config: {
      document: {
        fileType: "docx",
        key: savedResponseDocumentKey,
        permissions: {
          comment: false,
          download: false,
          edit: false,
          fillForms: false,
          review: false,
        },
      },
      editorConfig: { mode: "view" },
    },
  });
  const forbiddenDraftViewerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/viewer-config`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(forbiddenDraftViewerResponse.status).toBe(403);
  expect(
    await prisma.response.findUniqueOrThrow({
      select: { draftData: true, draftObjectKey: true },
      where: { id: responseId },
    })
  ).toMatchObject({
    draftData: savedDraftData,
    draftObjectKey: savedResponseObjectKey,
  });
  expect(await readObject(savedResponseObjectKey)).toEqual(
    savedResponseDocumentBytes
  );
  expect(JSON.stringify(adminDraftDetailBody)).not.toContain("draftObjectKey");
  const adminDraftExportResponse = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/json`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminDraftExportResponse.status).toBe(403);
  const userResultsResponse = await app.handle(
    new Request("http://test.local/api/admin/results", {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(userResultsResponse.status).toBe(403);
  const invalidAdminResultsCursor = await app.handle(
    new Request("http://test.local/api/admin/results?cursor=not-a-cursor", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(invalidAdminResultsCursor.status).toBe(400);
  const draftViewAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "view_response",
      actorId: adminId,
      targetId: responseId,
      targetType: "response",
    },
  });
  expect(draftViewAudit).toMatchObject({
    outcome: "success",
    safeMetadata: { state: "draft" },
  });
  expect(draftViewAudit.safeMetadata).toEqual({
    revision: 0,
    state: "draft",
  });
  const draftResponseBeforeExport = await prisma.response.findUnique({
    select: { draftData: true, draftObjectKey: true },
    where: { id: responseId },
  });
  if (!draftResponseBeforeExport?.draftObjectKey) {
    throw new Error("The Draft export fixture was not created");
  }
  const draftJsonExport = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/json`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftJsonExport.status).toBe(200);
  expect(draftJsonExport.headers.get("content-type")).toBe(
    "application/json; charset=utf-8"
  );
  expect(draftJsonExport.headers.get("content-disposition")).toBe(
    `attachment; filename="response-${responseId}.json"`
  );
  expect(JSON.parse(await draftJsonExport.text())).toEqual(savedDraftData);
  const draftDocxExport = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/docx`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftDocxExport.status).toBe(200);
  expect(draftDocxExport.headers.get("content-type")).toBe(DOCX_CONTENT_TYPE);
  expect(draftDocxExport.headers.get("content-disposition")).toBe(
    `attachment; filename="response-${responseId}.docx"`
  );
  expect(new Uint8Array(await draftDocxExport.arrayBuffer())).toEqual(
    Uint8Array.from(await readObject(draftResponseBeforeExport.draftObjectKey))
  );
  const draftPdfExport = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/pdf`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftPdfExport.status).toBe(200);
  expect(draftPdfExport.headers.get("content-type")).toBe("application/pdf");
  expect(draftPdfExport.headers.get("content-disposition")).toBe(
    `attachment; filename="response-${responseId}.pdf"`
  );
  expect(await draftPdfExport.text()).toBe("%PDF-test");
};
