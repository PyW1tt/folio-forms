// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";
import { createHash } from "node:crypto";

import type { DeletionTombstone } from "@onlyoffice/db";
import { prisma } from "@onlyoffice/db";

import type { createApp } from "../../../src/app";
import { objectExists } from "../../../src/storage";
import { jsonHeaders, bearerFor } from "../../fixtures/http";
import type { DraftLifecycleOutput } from "./draft";
import {
  accountRequest,
  replacePassword,
  sessionStatus,
  capabilityHeaders,
} from "./helpers";
import type { OfficePictureResponseOutput } from "./picture-response";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";
import type { SubmitFailuresAndCompletionOutput } from "./submission";

export interface AccountAuthorityInput {
  app: ReturnType<typeof createApp>;
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
}

export interface AccountAuthorityOutput {
  authorityAdminA: {
    email: string;
    id: string;
    password: string;
    temporaryPassword: string;
    token: string;
  };
  authorityAdminB: {
    email: string;
    id: string;
    password: string;
    temporaryPassword: string;
    token: string;
  };
  erasureAdminBearer: string;
  erasureAdminId: string;
}

export const runAccountAuthority = async (
  input: AccountAuthorityInput
): Promise<AccountAuthorityOutput> => {
  const { app, adminBearer } = input;

  const provisionAdmin = async (label: string) => {
    const email = `ticket-07-${label}-${crypto.randomUUID()}@example.com`;
    const provisionResponse = await accountRequest(
      app,
      "POST",
      "/api/admin/users",
      adminBearer,
      { email, name: `Ticket 07 ${label}`, role: "admin" }
    );
    expect(provisionResponse.status).toBe(200);
    const body = (await provisionResponse.json()) as {
      temporaryPassword?: unknown;
      user?: Record<string, unknown>;
    };
    if (
      typeof body.temporaryPassword !== "string" ||
      !body.user ||
      typeof body.user.id !== "string"
    ) {
      throw new Error("The Admin provisioning response was malformed");
    }
    const temporaryToken = await bearerFor(
      app,
      email,
      body.temporaryPassword,
      `ticket-07-${label}-temporary-${crypto.randomUUID()}`
    );
    const newPassword = `Ticket07-${label}-permanent-password`;
    const passwordChange = await replacePassword(
      app,
      temporaryToken,
      body.temporaryPassword,
      newPassword
    );
    expect(passwordChange.status).toBe(200);
    return {
      email,
      id: body.user.id,
      password: newPassword,
      temporaryPassword: body.temporaryPassword,
      token: await bearerFor(
        app,
        email,
        newPassword,
        `ticket-07-${label}-live-${crypto.randomUUID()}`
      ),
    };
  };
  const authorityAdminA = await provisionAdmin("authority-a");
  const authorityAdminB = await provisionAdmin("authority-b");
  const authorityListResponses = await Promise.all([
    accountRequest(app, "GET", "/api/admin/users", authorityAdminA.token),
    accountRequest(app, "GET", "/api/admin/users", authorityAdminB.token),
  ]);
  expect(authorityListResponses.map((response) => response.status)).toEqual([
    200, 200,
  ]);
  const enabledAdminsBeforeRace = await prisma.user.findMany({
    select: { id: true },
    where: { enabled: true, role: "admin" },
  });
  for (const enabledAdminBeforeRace of enabledAdminsBeforeRace) {
    if (
      enabledAdminBeforeRace.id === authorityAdminA.id ||
      enabledAdminBeforeRace.id === authorityAdminB.id
    ) {
      continue;
    }
    const demoteResponse = await accountRequest(
      app,
      "PATCH",
      `/api/admin/users/${enabledAdminBeforeRace.id}`,
      authorityAdminA.token,
      { role: "user" }
    );
    expect(demoteResponse.status).toBe(200);
  }
  expect(
    await prisma.user.count({ where: { enabled: true, role: "admin" } })
  ).toBe(2);
  const [disableRaceResponse, demoteRaceResponse] = await Promise.all([
    accountRequest(
      app,
      "PATCH",
      `/api/admin/users/${authorityAdminA.id}`,
      authorityAdminA.token,
      { enabled: false }
    ),
    accountRequest(
      app,
      "PATCH",
      `/api/admin/users/${authorityAdminB.id}`,
      authorityAdminB.token,
      { role: "user" }
    ),
  ]);
  expect(
    [disableRaceResponse.status, demoteRaceResponse.status].toSorted()
  ).toEqual([200, 409]);
  expect(
    [disableRaceResponse, demoteRaceResponse].some(
      (response) => response.status === 409
    )
  ).toBe(true);
  const raceBodies = await Promise.all([
    disableRaceResponse.json(),
    demoteRaceResponse.json(),
  ]);
  expect(
    raceBodies.some(
      (body) =>
        (body as Record<string, unknown>).error === "final_admin_required"
    )
  ).toBe(true);
  expect(
    await prisma.user.count({ where: { enabled: true, role: "admin" } })
  ).toBe(1);
  if (disableRaceResponse.status === 200) {
    expect(await sessionStatus(app, authorityAdminA.token)).toBe(401);
    expect(await sessionStatus(app, authorityAdminB.token)).toBe(200);
  } else {
    expect(await sessionStatus(app, authorityAdminA.token)).toBe(200);
    expect(await sessionStatus(app, authorityAdminB.token)).toBe(401);
  }

  const erasureAdmin =
    disableRaceResponse.status === 200 ? authorityAdminB : authorityAdminA;
  const erasureAdminBearer = erasureAdmin.token;
  const erasureAdminId = erasureAdmin.id;
  const finalAdminDelete = await accountRequest(
    app,
    "DELETE",
    `/api/admin/users/${erasureAdmin.id}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(finalAdminDelete.status).toBe(409);
  expect(await finalAdminDelete.json()).toMatchObject({
    error: "final_admin_required",
  });
  return {
    authorityAdminA,
    authorityAdminB,
    erasureAdminBearer,
    erasureAdminId,
  };
};

export interface ResponseErasureInput {
  app: ReturnType<typeof createApp>;
  pictureResponseId: OfficePictureResponseOutput["pictureResponseId"];
  pictureCorrectionEditor: OfficePictureResponseOutput["pictureCorrectionEditor"];
  erasureAdminBearer: AccountAuthorityOutput["erasureAdminBearer"];
  userBearer: PrefillEntryOutput["userBearer"];
  erasureAdminId: AccountAuthorityOutput["erasureAdminId"];
  userId: PrefillEntryOutput["userId"];
  responseId: DraftLifecycleOutput["responseId"];
  userLease: SubmitFailuresAndCompletionOutput["userLease"];
  postSubmitHandoffRecord: SubmitFailuresAndCompletionOutput["postSubmitHandoffRecord"];
  mock: PrefillEntryOutput["mock"];
  postSubmitExternalReference: SubmitFailuresAndCompletionOutput["postSubmitExternalReference"];
  formRecord: PrefillEntryOutput["formRecord"];
  responseDocumentKey: SubmitFailuresAndCompletionOutput["responseDocumentKey"];
  saveDraftCapability: SubmitFailuresAndCompletionOutput["saveDraftCapability"];
  userEmail: BootstrapAndCreationOutput["userEmail"];
  savedDraftData: DraftLifecycleOutput["savedDraftData"];
}

export interface ResponseErasureOutput {
  mainTombstone: DeletionTombstone;
}

export const runResponseErasure = async (
  input: ResponseErasureInput
): Promise<ResponseErasureOutput> => {
  const {
    app,
    pictureResponseId,
    pictureCorrectionEditor,
    erasureAdminBearer,
    userBearer,
    erasureAdminId,
    userId,
    responseId,
    userLease,
    postSubmitHandoffRecord,
    mock,
    postSubmitExternalReference,
    formRecord,
    responseDocumentKey,
    saveDraftCapability,
    userEmail,
    savedDraftData,
  } = input;
  const releaseLeaseIfPresent = async (
    targetId: string,
    releaseUrl: string,
    token: string
  ) => {
    const leaseCount = await prisma.editorLease.count({
      where: { targetId },
    });
    if (leaseCount === 0) {
      return;
    }
    const leaseReleaseResponse = await app.handle(
      new Request(new URL(releaseUrl, "http://test.local").toString(), {
        headers: { Authorization: `Bearer ${token}` },
        method: "DELETE",
      })
    );
    expect(leaseReleaseResponse.status).toBe(200);
  };

  const pictureResponseBeforeDeletion = await prisma.response.findUniqueOrThrow(
    {
      select: {
        draftObjectKey: true,
        externalReferenceDigest: true,
      },
      where: { id: pictureResponseId },
    }
  );
  const pictureSubmissionBeforeDeletion =
    await prisma.submission.findUniqueOrThrow({
      select: { id: true, objectKey: true },
      where: { responseId: pictureResponseId },
    });
  const pictureCorrectionBeforeDeletion =
    await prisma.correction.findFirstOrThrow({
      select: { id: true, objectKey: true },
      where: { responseId: pictureResponseId },
    });
  await releaseLeaseIfPresent(
    pictureResponseId,
    pictureCorrectionEditor.bridge.lease.releaseUrl,
    erasureAdminBearer
  );
  const forbiddenResponseDeletion = await app.handle(
    new Request(`http://test.local/api/admin/responses/${pictureResponseId}`, {
      body: JSON.stringify({ confirm: true }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(forbiddenResponseDeletion.status).toBe(403);
  const malformedResponseDeletion = await accountRequest(
    app,
    "DELETE",
    `/api/admin/responses/${pictureResponseId}`,
    erasureAdminBearer,
    { confirm: false }
  );
  expect(malformedResponseDeletion.status).toBe(400);
  const pictureDeletionResponse = await accountRequest(
    app,
    "DELETE",
    `/api/admin/responses/${pictureResponseId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(pictureDeletionResponse.status).toBe(200);
  expect(await pictureDeletionResponse.json()).toEqual({ deleted: true });
  expect(
    await prisma.response.findUnique({ where: { id: pictureResponseId } })
  ).toBeNull();
  expect(
    await prisma.submission.findUnique({
      where: { id: pictureSubmissionBeforeDeletion.id },
    })
  ).toBeNull();
  expect(
    await prisma.correction.findUnique({
      where: { id: pictureCorrectionBeforeDeletion.id },
    })
  ).toBeNull();
  expect(
    await prisma.prefillSnapshot.count({
      where: { responseId: pictureResponseId },
    })
  ).toBe(0);
  expect(
    await prisma.editorLease.count({ where: { targetId: pictureResponseId } })
  ).toBe(0);
  expect(
    await prisma.operation.count({
      where: {
        OR: [
          { responseId: pictureResponseId },
          { submissionId: pictureSubmissionBeforeDeletion.id },
          { correctionId: pictureCorrectionBeforeDeletion.id },
        ],
      },
    })
  ).toBe(0);
  if (pictureResponseBeforeDeletion.draftObjectKey) {
    expect(
      await objectExists(pictureResponseBeforeDeletion.draftObjectKey)
    ).toBe(false);
  }
  expect(await objectExists(pictureSubmissionBeforeDeletion.objectKey)).toBe(
    false
  );
  expect(await objectExists(pictureCorrectionBeforeDeletion.objectKey)).toBe(
    false
  );
  const pictureDeletionAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "delete_response",
      outcome: "success",
      targetId: pictureResponseId,
    },
  });
  expect(pictureDeletionAudit.safeMetadata).toEqual({});
  expect(JSON.stringify(pictureDeletionAudit)).not.toContain("ตรวจสอบรูปภาพ");
  expect(JSON.stringify(pictureDeletionAudit)).not.toContain(
    pictureSubmissionBeforeDeletion.objectKey
  );
  const pictureTombstone = await prisma.deletionTombstone.findUniqueOrThrow({
    where: {
      responseLookupDigest: createHash("sha256")
        .update(pictureResponseId)
        .digest("hex"),
    },
  });
  expect(pictureTombstone).toMatchObject({
    actorId: erasureAdminId,
    externalReferenceDigest:
      pictureResponseBeforeDeletion.externalReferenceDigest,
    outcome: "success",
  });

  const blockedUserDeletion = await accountRequest(
    app,
    "DELETE",
    `/api/admin/users/${userId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(blockedUserDeletion.status).toBe(409);
  expect(await blockedUserDeletion.json()).toMatchObject({
    error: "personal_data_remains",
  });
  await releaseLeaseIfPresent(responseId, userLease.releaseUrl, userBearer);
  const mainResponseBeforeDeletion = await prisma.response.findUniqueOrThrow({
    select: {
      draftObjectKey: true,
      externalReferenceDigest: true,
    },
    where: { id: responseId },
  });
  const mainSubmissionBeforeDeletion =
    await prisma.submission.findUniqueOrThrow({
      select: { id: true, objectKey: true },
      where: { responseId },
    });
  const mainDeletionResponse = await accountRequest(
    app,
    "DELETE",
    `/api/admin/responses/${responseId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(mainDeletionResponse.status).toBe(200);
  expect(await mainDeletionResponse.json()).toEqual({ deleted: true });
  expect(await sessionStatus(app, userBearer)).toBe(401);
  expect(
    await prisma.response.findUnique({ where: { id: responseId } })
  ).toBeNull();
  expect(
    await prisma.submission.findUnique({
      where: { id: mainSubmissionBeforeDeletion.id },
    })
  ).toBeNull();
  expect(await prisma.prefillSnapshot.count({ where: { responseId } })).toBe(0);
  expect(
    await prisma.editorLease.count({ where: { targetId: responseId } })
  ).toBe(0);
  expect(
    await prisma.operation.count({
      where: {
        OR: [{ responseId }, { submissionId: mainSubmissionBeforeDeletion.id }],
      },
    })
  ).toBe(0);
  if (mainResponseBeforeDeletion.draftObjectKey) {
    expect(await objectExists(mainResponseBeforeDeletion.draftObjectKey)).toBe(
      false
    );
  }
  expect(await objectExists(mainSubmissionBeforeDeletion.objectKey)).toBe(
    false
  );
  const deletedHandoff = await prisma.handoff.findUniqueOrThrow({
    where: { id: postSubmitHandoffRecord.id },
  });
  expect(deletedHandoff).toMatchObject({
    codeDigest: null,
    configurationHash: null,
    deletionResponseLookupDigest: createHash("sha256")
      .update(responseId)
      .digest("hex"),
    filteredValues: null,
    formId: null,
    normalizedEmail: null,
    responseId: null,
    status: "deleted",
  });
  const deletedExternalStatus = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(deletedExternalStatus.status).toBe(200);
  const deletedExternalBody = await deletedExternalStatus.json();
  expect(deletedExternalBody).toMatchObject({
    deletedAt: expect.any(String),
    status: "deleted",
  });
  const staleCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect([401, 404]).toContain(staleCapabilityResponse.status);
  const pendingStatusExternalReference = `ticket-22-pending-status-${crypto.randomUUID()}`;
  const pendingStatusResponseDigest = createHash("sha256")
    .update(`ticket-22-pending-response-${crypto.randomUUID()}`)
    .digest("hex");
  await prisma.deletionTombstone.create({
    data: {
      actorId: erasureAdminId,
      externalReferenceDigest: createHash("sha256")
        .update(pendingStatusExternalReference)
        .digest("hex"),
      id: crypto.randomUUID(),
      outcome: "success",
      responseLookupDigest: pendingStatusResponseDigest,
    },
  });
  const pendingStatusObjectKey = `ticket-22-pending-status/${crypto.randomUUID()}`;
  await prisma.objectCleanupIntent.create({
    data: {
      deletionResponseLookupDigest: pendingStatusResponseDigest,
      objectKey: pendingStatusObjectKey,
    },
  });
  const pendingExternalStatus = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({
      externalReference: pendingStatusExternalReference,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(pendingExternalStatus.status).toBe(404);
  expect(await pendingExternalStatus.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  await prisma.objectCleanupIntent.delete({
    where: { objectKey: pendingStatusObjectKey },
  });
  const completedExternalStatus = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({
      externalReference: pendingStatusExternalReference,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(completedExternalStatus.status).toBe(200);
  expect(await completedExternalStatus.json()).toMatchObject({
    status: "deleted",
  });
  expect(JSON.stringify(deletedExternalBody)).not.toContain(userEmail);
  expect(JSON.stringify(deletedExternalBody)).not.toContain(
    "Ticket 17 Prefill"
  );
  const mainDeletionAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "delete_response",
      outcome: "success",
      targetId: responseId,
    },
  });
  expect(mainDeletionAudit.safeMetadata).toEqual({});
  expect(JSON.stringify(mainDeletionAudit)).not.toContain(
    JSON.stringify(savedDraftData)
  );
  expect(JSON.stringify(mainDeletionAudit)).not.toContain(
    postSubmitExternalReference
  );
  const mainTombstone = await prisma.deletionTombstone.findUniqueOrThrow({
    where: {
      responseLookupDigest: createHash("sha256")
        .update(responseId)
        .digest("hex"),
    },
  });
  expect(mainTombstone).toMatchObject({
    actorId: erasureAdminId,
    externalReferenceDigest: mainResponseBeforeDeletion.externalReferenceDigest,
    outcome: "success",
  });
  const repeatedMainDeletion = await accountRequest(
    app,
    "DELETE",
    `/api/admin/responses/${responseId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(repeatedMainDeletion.status).toBe(200);
  expect(await repeatedMainDeletion.json()).toEqual({ deleted: true });
  const deletedOwnerAccount = await accountRequest(
    app,
    "DELETE",
    `/api/admin/users/${userId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(deletedOwnerAccount.status).toBe(200);
  expect(await deletedOwnerAccount.json()).toEqual({ deleted: true });
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_user",
        outcome: "success",
        targetId: userId,
      },
    })
  ).toMatchObject({ actorId: erasureAdminId, targetId: userId });
  expect(await prisma.user.findUnique({ where: { id: userId } })).toBeNull();
  const deletedExternalStatusAfterAccount = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(deletedExternalStatusAfterAccount.status).toBe(200);
  expect(await deletedExternalStatusAfterAccount.json()).toMatchObject({
    deletedAt: expect.any(String),
    status: "deleted",
  });
  return {
    mainTombstone,
  };
};
