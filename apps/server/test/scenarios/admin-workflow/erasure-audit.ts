// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect, vi } from "bun:test";
import { createHash } from "node:crypto";

import { prisma } from "@onlyoffice/db";

import { AiAuthoringSessions } from "../../../src/ai-authoring";
import { createApp } from "../../../src/app";
import { objectExists, deleteObject } from "../../../src/storage";
import { jsonHeaders, bearerFor } from "../../fixtures/http";
import type {
  AccountAuthorityOutput,
  ResponseErasureOutput,
} from "./account-erasure";
import type { ManagedAccountsOutput } from "./account-lifecycle";
import type { DraftLifecycleOutput } from "./draft";
import type {
  FieldConfigurationOutput,
  EditorAccessOutput,
} from "./editor-access";
import { accountRequest, sessionStatus } from "./helpers";
import type { OfficePictureResponseOutput } from "./picture-response";
import type { PrefillEntryOutput } from "./prefill-entry";
import type {
  BootstrapAndCreationOutput,
  TemplateUploadsOutput,
} from "./setup";
import type {
  SubmissionExportsOutput,
  SubmitFailuresAndCompletionOutput,
} from "./submission";

export interface ErasureFailuresInput {
  app: ReturnType<typeof createApp>;
  formRecord: PrefillEntryOutput["formRecord"];
  otherUserBearer: SubmissionExportsOutput["otherUserBearer"];
  erasureAdminBearer: AccountAuthorityOutput["erasureAdminBearer"];
  otherUser: SubmissionExportsOutput["otherUser"];
  otherUserEmail: BootstrapAndCreationOutput["otherUserEmail"];
  password: BootstrapAndCreationOutput["password"];
  handoffCandidateValues: PrefillEntryOutput["handoffCandidateValues"];
  prefillHandoffSecret: PrefillEntryOutput["prefillHandoffSecret"];
}

export const runErasureFailures = async (
  input: ErasureFailuresInput
): Promise<void> => {
  const {
    app,
    formRecord,
    otherUserBearer,
    erasureAdminBearer,
    otherUser,
    otherUserEmail,
    password,
    handoffCandidateValues,
    prefillHandoffSecret,
  } = input;

  const cleanupStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${otherUserBearer}` },
      method: "POST",
    })
  );
  expect(cleanupStart.status).toBe(200);
  const cleanupStartBody = (await cleanupStart.json()) as {
    response?: { id?: string };
  };
  const cleanupResponseId = cleanupStartBody.response?.id;
  if (!cleanupResponseId) {
    throw new Error("The cleanup failure response was not created");
  }
  const cleanupResponse = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: cleanupResponseId },
  });
  if (!cleanupResponse.draftObjectKey) {
    throw new Error("The cleanup failure object was not created");
  }
  const databaseFailureResponseId = cleanupResponseId;
  await prisma.$executeRawUnsafe(`
    CREATE OR REPLACE FUNCTION ticket22_abort_response_delete()
    RETURNS trigger
    LANGUAGE plpgsql
    AS $function$
    BEGIN
      RAISE EXCEPTION 'ticket22 database failure';
    END;
    $function$;
  `);
  await prisma.$executeRawUnsafe(`
    CREATE TRIGGER ticket22_abort_response_delete
    BEFORE DELETE ON "responses"
    FOR EACH ROW
    EXECUTE FUNCTION ticket22_abort_response_delete();
  `);
  let databaseFailureDeletion: Response;
  try {
    databaseFailureDeletion = await app.handle(
      new Request(
        `http://test.local/api/admin/responses/${databaseFailureResponseId}`,
        {
          body: JSON.stringify({ confirm: true }),
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${erasureAdminBearer}`,
          },
          method: "DELETE",
        }
      )
    );
  } finally {
    await prisma.$executeRawUnsafe(`
      DROP TRIGGER IF EXISTS ticket22_abort_response_delete ON "responses";
    `);
    await prisma.$executeRawUnsafe(`
      DROP FUNCTION IF EXISTS ticket22_abort_response_delete();
    `);
  }
  expect(databaseFailureDeletion.status).toBe(500);
  expect(await databaseFailureDeletion.json()).toMatchObject({
    error: "internal_error",
  });
  expect(
    await prisma.response.findUnique({
      where: { id: databaseFailureResponseId },
    })
  ).not.toBeNull();
  expect(
    await prisma.objectCleanupIntent.count({
      where: {
        deletionResponseLookupDigest: createHash("sha256")
          .update(databaseFailureResponseId)
          .digest("hex"),
      },
    })
  ).toBe(0);
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_response",
        outcome: "failure",
        targetId: databaseFailureResponseId,
      },
    })
  ).toMatchObject({ safeMetadata: { errorCode: "internal_error" } });
  let failCleanup = true;
  const cleanupFailureApp = createApp({
    deleteObject: async (key) => {
      if (failCleanup) {
        throw new Error(`deterministic cleanup failure: ${key}`);
      }
      await deleteObject(key);
    },
  });
  const failedCleanupDeletion = await cleanupFailureApp.handle(
    new Request(`http://test.local/api/admin/responses/${cleanupResponseId}`, {
      body: JSON.stringify({ confirm: true }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${erasureAdminBearer}`,
      },
      method: "DELETE",
    })
  );
  expect(failedCleanupDeletion.status).toBe(503);
  expect(await failedCleanupDeletion.json()).toMatchObject({
    error: "deletion_cleanup_failed",
  });
  expect(
    await prisma.response.findUnique({ where: { id: cleanupResponseId } })
  ).toBeNull();
  expect(
    await prisma.objectCleanupIntent.count({
      where: {
        deletionResponseLookupDigest: createHash("sha256")
          .update(cleanupResponseId)
          .digest("hex"),
      },
    })
  ).toBeGreaterThan(0);
  expect(
    await prisma.auditEvent.findFirst({
      where: {
        action: "delete_response",
        outcome: "success",
        targetId: cleanupResponseId,
      },
    })
  ).toBeNull();
  expect(
    await prisma.auditEvent.findFirst({
      where: {
        action: "delete_response_pending",
        outcome: "success",
        targetId: cleanupResponseId,
      },
    })
  ).toMatchObject({ safeMetadata: {} });
  const blockedCleanupOwner = await accountRequest(
    app,
    "DELETE",
    `/api/admin/users/${otherUser.id}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(blockedCleanupOwner.status).toBe(409);
  expect(await blockedCleanupOwner.json()).toMatchObject({
    error: "personal_data_remains",
  });
  failCleanup = false;
  const retriedCleanupDeletion = await cleanupFailureApp.handle(
    new Request(`http://test.local/api/admin/responses/${cleanupResponseId}`, {
      body: JSON.stringify({ confirm: true }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${erasureAdminBearer}`,
      },
      method: "DELETE",
    })
  );
  expect(retriedCleanupDeletion.status).toBe(200);
  expect(await retriedCleanupDeletion.json()).toEqual({ deleted: true });
  expect(
    await prisma.objectCleanupIntent.count({
      where: {
        deletionResponseLookupDigest: createHash("sha256")
          .update(cleanupResponseId)
          .digest("hex"),
      },
    })
  ).toBe(0);
  expect(
    await prisma.auditEvent.findFirst({
      where: {
        action: "delete_response",
        outcome: "success",
        targetId: cleanupResponseId,
      },
    })
  ).toMatchObject({ safeMetadata: {} });
  expect(await objectExists(cleanupResponse.draftObjectKey)).toBe(false);
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_response",
        outcome: "failure",
        targetId: cleanupResponseId,
      },
    })
  ).toMatchObject({
    safeMetadata: { errorCode: "deletion_cleanup_failed" },
  });

  const cleanupAfterAiBearer = await bearerFor(
    app,
    otherUserEmail,
    password,
    `ticket-17-ai-cleanup-${crypto.randomUUID()}`
  );
  const cleanupAfterAiHandoff = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: otherUserEmail,
        externalReference: `ticket-17-ai-cleanup-${crypto.randomUUID()}`,
        publicId: formRecord.publicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(cleanupAfterAiHandoff.status).toBe(200);
  const cleanupAfterAiHandoffBody: unknown = await cleanupAfterAiHandoff.json();
  if (
    !cleanupAfterAiHandoffBody ||
    typeof cleanupAfterAiHandoffBody !== "object" ||
    !("code" in cleanupAfterAiHandoffBody) ||
    typeof cleanupAfterAiHandoffBody.code !== "string" ||
    !cleanupAfterAiHandoffBody.code
  ) {
    throw new Error("The AI cleanup handoff was not created");
  }
  const cleanupAfterAiHandoffCode = cleanupAfterAiHandoffBody.code;
  const cleanupAfterAiLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: cleanupAfterAiHandoffCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    })
  );
  expect(cleanupAfterAiLaunch.status).toBe(303);
  const cleanupAfterAiPendingCookie = cleanupAfterAiLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!cleanupAfterAiPendingCookie) {
    throw new Error(
      "The AI cleanup handoff did not set a pending claim cookie"
    );
  }
  const cleanupAfterAiStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${cleanupAfterAiBearer}`,
        Cookie: cleanupAfterAiPendingCookie,
      },
      method: "POST",
    })
  );
  const cleanupAfterAiBody = (await cleanupAfterAiStart.json()) as {
    response?: { id?: string };
  };
  const cleanupAfterAiResponseId = cleanupAfterAiBody.response?.id;
  if (!cleanupAfterAiResponseId) {
    throw new Error("The AI cleanup response was not created");
  }
  const cleanupAfterAiResponse = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: cleanupAfterAiResponseId },
  });
  if (!cleanupAfterAiResponse.draftObjectKey) {
    throw new Error("The AI cleanup response object was not created");
  }
  const cleanupAfterAiDigest = createHash("sha256")
    .update(cleanupAfterAiResponseId)
    .digest("hex");
  const endForSession = vi
    .spyOn(AiAuthoringSessions.prototype, "endForSession")
    .mockRejectedValueOnce(new Error("AI authoring cleanup failed"));
  let cleanupAfterAiDeletion: Response;
  try {
    cleanupAfterAiDeletion = await cleanupFailureApp.handle(
      new Request(
        `http://test.local/api/admin/responses/${cleanupAfterAiResponseId}`,
        {
          body: JSON.stringify({ confirm: true }),
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${erasureAdminBearer}`,
          },
          method: "DELETE",
        }
      )
    );
  } finally {
    endForSession.mockRestore();
  }
  expect(cleanupAfterAiDeletion.status).toBe(500);
  expect(await cleanupAfterAiDeletion.json()).toMatchObject({
    error: "internal_error",
  });
  expect(
    await prisma.response.findUnique({
      where: { id: cleanupAfterAiResponseId },
    })
  ).toBeNull();
  expect(
    await prisma.objectCleanupIntent.count({
      where: { deletionResponseLookupDigest: cleanupAfterAiDigest },
    })
  ).toBe(0);
  expect(await objectExists(cleanupAfterAiResponse.draftObjectKey)).toBe(false);
  expect(await sessionStatus(app, cleanupAfterAiBearer)).toBe(401);
};

export interface ImmutableAuditsInput {
  app: ReturnType<typeof createApp>;
  managedId: ManagedAccountsOutput["managedId"];
  erasureAdminBearer: AccountAuthorityOutput["erasureAdminBearer"];
  managedToken: ManagedAccountsOutput["managedToken"];
  erasureAdminId: AccountAuthorityOutput["erasureAdminId"];
  managedTemporaryPassword: ManagedAccountsOutput["managedTemporaryPassword"];
  resetTemporaryPassword: ManagedAccountsOutput["resetTemporaryPassword"];
  authorityAdminA: AccountAuthorityOutput["authorityAdminA"];
  authorityAdminB: AccountAuthorityOutput["authorityAdminB"];
  managedPassword: ManagedAccountsOutput["managedPassword"];
  resetPassword: ManagedAccountsOutput["resetPassword"];
  invalidTargetSecret: ManagedAccountsOutput["invalidTargetSecret"];
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
  competingAdminBearer: TemplateUploadsOutput["competingAdminBearer"];
  secretFormTitle: BootstrapAndCreationOutput["secretFormTitle"];
  secretFormDescription: BootstrapAndCreationOutput["secretFormDescription"];
  templateDocumentKey: BootstrapAndCreationOutput["templateDocumentKey"];
  createdFormRecord: BootstrapAndCreationOutput["createdFormRecord"];
  publishCapability: FieldConfigurationOutput["publishCapability"];
  saveTemplateCapability: EditorAccessOutput["saveTemplateCapability"];
  formPublicId: BootstrapAndCreationOutput["formPublicId"];
  uploadPublicId: TemplateUploadsOutput["uploadPublicId"];
  competingAdmin: TemplateUploadsOutput["competingAdmin"];
  otherUserEmail: BootstrapAndCreationOutput["otherUserEmail"];
  password: BootstrapAndCreationOutput["password"];
  pictureResponseId: OfficePictureResponseOutput["pictureResponseId"];
  userBearer: PrefillEntryOutput["userBearer"];
  saveDraftCapability: SubmitFailuresAndCompletionOutput["saveDraftCapability"];
  submitCapability: SubmitFailuresAndCompletionOutput["submitCapability"];
  postSubmitExternalReference: SubmitFailuresAndCompletionOutput["postSubmitExternalReference"];
  savedDraftData: DraftLifecycleOutput["savedDraftData"];
  mainTombstone: ResponseErasureOutput["mainTombstone"];
}

export const runImmutableAudits = async (
  input: ImmutableAuditsInput
): Promise<void> => {
  const {
    app,
    managedId,
    erasureAdminBearer,
    managedToken,
    erasureAdminId,
    managedTemporaryPassword,
    resetTemporaryPassword,
    authorityAdminA,
    authorityAdminB,
    managedPassword,
    resetPassword,
    invalidTargetSecret,
    adminBearer,
    competingAdminBearer,
    secretFormTitle,
    secretFormDescription,
    templateDocumentKey,
    createdFormRecord,
    publishCapability,
    saveTemplateCapability,
    formPublicId,
    uploadPublicId,
    competingAdmin,
    otherUserEmail,
    password,
    pictureResponseId,
    userBearer,
    saveDraftCapability,
    submitCapability,
    postSubmitExternalReference,
    savedDraftData,
    mainTombstone,
  } = input;

  const managedDeleteResponse = await accountRequest(
    app,
    "DELETE",
    `/api/admin/users/${managedId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(managedDeleteResponse.status).toBe(200);
  expect(await managedDeleteResponse.json()).toEqual({ deleted: true });
  expect(await sessionStatus(app, managedToken)).toBe(401);
  expect(await prisma.user.findUnique({ where: { id: managedId } })).toBeNull();
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_user",
        outcome: "success",
        targetId: managedId,
      },
    })
  ).toMatchObject({ actorId: erasureAdminId, targetId: managedId });

  const accountAuditEvents = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: { targetType: "user" },
  });
  const auditText = JSON.stringify(accountAuditEvents);
  for (const secret of [
    managedTemporaryPassword,
    resetTemporaryPassword,
    authorityAdminA.temporaryPassword,
    authorityAdminB.temporaryPassword,
    managedPassword,
    resetPassword,
    authorityAdminA.password,
    authorityAdminB.password,
    invalidTargetSecret,
  ]) {
    expect(auditText).not.toContain(secret);
  }
  for (const event of accountAuditEvents) {
    const metadata =
      event.safeMetadata &&
      typeof event.safeMetadata === "object" &&
      !Array.isArray(event.safeMetadata)
        ? (event.safeMetadata as Record<string, unknown>)
        : null;
    expect(
      Object.keys(metadata ?? {}).every(
        (key) => key === "change" || key === "errorCode"
      )
    ).toBe(true);
  }
  const managedAuditActions = new Set(
    accountAuditEvents
      .filter((event) => event.targetId === managedId)
      .map((event) => event.action)
  );
  for (const action of [
    "create_user",
    "enable_user",
    "disable_user",
    "change_user_email",
    "promote_user",
    "demote_user",
    "reset_user_password",
    "delete_user",
  ]) {
    expect(managedAuditActions.has(action)).toBe(true);
  }
  expect(
    accountAuditEvents.some(
      (event) =>
        event.action === "create_user" &&
        event.outcome === "failure" &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "email_in_use"
    )
  ).toBe(true);
  expect(
    accountAuditEvents.some(
      (event) =>
        event.outcome === "failure" &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "final_admin_required"
    )
  ).toBe(true);
  expect(
    accountAuditEvents.some(
      (event) =>
        event.action === "update_user" &&
        event.outcome === "failure" &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "invalid_request"
    )
  ).toBe(true);
  const formAuditEvents = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: { targetType: "form" },
  });
  const formAuditText = JSON.stringify(formAuditEvents);
  for (const secret of [
    adminBearer,
    competingAdminBearer,
    secretFormTitle,
    secretFormDescription,
    templateDocumentKey,
    createdFormRecord.templateDraft.objectKey,
    publishCapability,
    saveTemplateCapability,
  ]) {
    expect(formAuditText).not.toContain(secret);
  }
  for (const event of formAuditEvents) {
    const metadata =
      event.safeMetadata &&
      typeof event.safeMetadata === "object" &&
      !Array.isArray(event.safeMetadata)
        ? (event.safeMetadata as Record<string, unknown>)
        : null;
    expect(
      Object.entries(metadata ?? {}).every(
        ([key, value]) =>
          key === "errorCode" ||
          key === "source" ||
          key === "sourcePublicId" ||
          key === "status" ||
          (key === "fillMethod" &&
            (value === "native" || value === "onlyoffice"))
      )
    ).toBe(true);
  }
  for (const expected of [
    {
      action: "create_form",
      outcome: "success",
      targetId: formPublicId,
    },
    {
      action: "create_form",
      outcome: "failure",
      targetId: null,
    },
    {
      action: "delete_form",
      outcome: "success",
      targetId: uploadPublicId,
    },
    {
      action: "delete_form",
      outcome: "failure",
      targetId: formPublicId,
    },
    {
      action: "save_template_draft",
      outcome: "success",
      targetId: formPublicId,
    },
    {
      action: "save_template_draft",
      outcome: "failure",
      targetId: formPublicId,
    },
    {
      action: "publish_form",
      outcome: "success",
      targetId: formPublicId,
    },
    {
      action: "publish_form",
      outcome: "failure",
      targetId: formPublicId,
    },
  ] as const) {
    expect(
      formAuditEvents.some(
        (event) =>
          event.action === expected.action &&
          event.outcome === expected.outcome &&
          event.targetId === expected.targetId
      )
    ).toBe(true);
  }
  expect(
    formAuditEvents.some(
      (event) =>
        event.action === "delete_form" &&
        event.actorId === competingAdmin.id &&
        event.outcome === "failure" &&
        event.targetId === uploadPublicId &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "editor_in_use"
    )
  ).toBe(true);
  const auditUnauthenticatedResponse = await app.handle(
    new Request("http://test.local/api/admin/audit-events")
  );
  expect(auditUnauthenticatedResponse.status).toBe(401);
  const auditUserBearer = await bearerFor(
    app,
    otherUserEmail,
    password,
    `ticket-23-audit-user-${crypto.randomUUID()}`
  );
  const auditForbiddenResponse = await app.handle(
    new Request("http://test.local/api/admin/audit-events", {
      headers: { Authorization: `Bearer ${auditUserBearer}` },
    })
  );
  expect(auditForbiddenResponse.status).toBe(403);
  expect(await auditForbiddenResponse.json()).toMatchObject({
    error: "forbidden",
  });
  const filteredAuditResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/audit-events?action=delete_response&target=${encodeURIComponent(pictureResponseId)}&outcome=success`,
      { headers: { Authorization: `Bearer ${erasureAdminBearer}` } }
    )
  );
  expect(filteredAuditResponse.status).toBe(200);
  const filteredAuditBody = (await filteredAuditResponse.json()) as {
    events?: {
      action: string;
      outcome: string;
      safeMetadata: Record<string, unknown>;
      targetId: string | null;
    }[];
    nextCursor?: string | null;
  };
  expect(filteredAuditBody.events).toHaveLength(1);
  expect(filteredAuditBody.events?.[0]).toMatchObject({
    action: "delete_response",
    outcome: "success",
    safeMetadata: {},
    targetId: pictureResponseId,
  });
  const auditPageResponse = await app.handle(
    new Request("http://test.local/api/admin/audit-events", {
      headers: { Authorization: `Bearer ${erasureAdminBearer}` },
    })
  );
  expect(auditPageResponse.status).toBe(200);
  const auditPageBody = (await auditPageResponse.json()) as {
    events?: {
      action: string;
      createdAt: string;
      id: string;
      safeMetadata: Record<string, unknown>;
    }[];
    nextCursor?: string | null;
  };
  if (!auditPageBody.events || auditPageBody.events.length === 0) {
    throw new Error("The audit API returned no events");
  }
  for (let index = 1; index < auditPageBody.events.length; index += 1) {
    const previous = auditPageBody.events[index - 1];
    const current = auditPageBody.events[index];
    if (!previous || !current) {
      continue;
    }
    expect(
      new Date(previous.createdAt).getTime() >=
        new Date(current.createdAt).getTime()
    ).toBe(true);
  }
  if (auditPageBody.nextCursor) {
    const nextAuditPageResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/audit-events?cursor=${encodeURIComponent(auditPageBody.nextCursor)}`,
        { headers: { Authorization: `Bearer ${erasureAdminBearer}` } }
      )
    );
    expect(nextAuditPageResponse.status).toBe(200);
    const nextAuditPageBody = (await nextAuditPageResponse.json()) as {
      events?: { id: string }[];
    };
    const auditPageIds = new Set(auditPageBody.events.map((event) => event.id));
    for (const event of nextAuditPageBody.events ?? []) {
      expect(auditPageIds.has(event.id)).toBe(false);
    }
  }
  for (const invalidAuditQuery of [
    "?actor=not-a-uuid",
    "?from=2026-09-20T00:00:00.000Z&to=2026-09-19T00:00:00.000Z",
  ]) {
    const invalidAuditResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/audit-events${invalidAuditQuery}`,
        { headers: { Authorization: `Bearer ${erasureAdminBearer}` } }
      )
    );
    expect(invalidAuditResponse.status).toBe(400);
  }
  const auditApiText = JSON.stringify({
    events: auditPageBody.events,
  });
  for (const secret of [
    adminBearer,
    userBearer,
    saveDraftCapability,
    submitCapability,
    postSubmitExternalReference,
    savedDraftData,
    password,
    managedPassword,
    resetPassword,
  ]) {
    expect(auditApiText).not.toContain(
      typeof secret === "string" ? secret : JSON.stringify(secret)
    );
  }
  let auditMutationBlocked = false;
  const auditProbe = await prisma.auditEvent.findFirstOrThrow();
  try {
    await prisma.auditEvent.update({
      data: { action: "tampered" },
      where: { id: auditProbe.id },
    });
  } catch {
    auditMutationBlocked = true;
  }
  expect(auditMutationBlocked).toBe(true);
  let tombstoneMutationBlocked = false;
  try {
    await prisma.deletionTombstone.update({
      data: { outcome: "failure" },
      where: { responseLookupDigest: mainTombstone.responseLookupDigest },
    });
  } catch {
    tombstoneMutationBlocked = true;
  }
  expect(tombstoneMutationBlocked).toBe(true);
};
