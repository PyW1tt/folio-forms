// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";

import type { createApp } from "../../../src/app";
import {
  createCredentialFixture,
  bearerFor,
  signIn,
} from "../../fixtures/http";
import { replacePassword, accountRequest, sessionStatus } from "./helpers";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";

export interface PasswordAndThrottleInput {
  app: ReturnType<typeof createApp>;
  formRecord: PrefillEntryOutput["formRecord"];
}

export const runPasswordAndThrottle = async (
  input: PasswordAndThrottleInput
): Promise<void> => {
  const { app, formRecord } = input;

  const passwordEmail = `ticket-04-password-${crypto.randomUUID()}@example.com`;
  const currentPassword = "Ticket04-current-password";
  const passwordUser = await createCredentialFixture({
    email: passwordEmail,
    mustChangePassword: true,
    name: "Ticket 04 Password User",
    password: currentPassword,
  });
  const mandatoryToken = await bearerFor(
    app,
    passwordEmail,
    currentPassword,
    `ticket-04-password-${crypto.randomUUID()}`
  );
  const mandatorySessionResponse = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${mandatoryToken}` },
    })
  );
  expect(mandatorySessionResponse.status).toBe(200);
  expect(await mandatorySessionResponse.json()).toMatchObject({
    user: { mustChangePassword: true },
  });
  const restrictedProductResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}`, {
      headers: { Authorization: `Bearer ${mandatoryToken}` },
    })
  );
  expect(restrictedProductResponse.status).toBe(403);
  expect(await restrictedProductResponse.json()).toMatchObject({
    error: "password_change_required",
  });

  const tooShortResponse = await replacePassword(
    app,
    mandatoryToken,
    currentPassword,
    "x".repeat(11)
  );
  expect(tooShortResponse.status).toBe(400);
  expect(await tooShortResponse.json()).toMatchObject({
    error: "password_too_short",
  });

  const spacesPassword = " ".repeat(12);
  const spacesReplacementResponse = await replacePassword(
    app,
    mandatoryToken,
    currentPassword,
    spacesPassword
  );
  expect(spacesReplacementResponse.status).toBe(200);
  expect(await spacesReplacementResponse.json()).toEqual({ ok: true });
  const invalidatedMandatorySession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${mandatoryToken}` },
    })
  );
  expect(invalidatedMandatorySession.status).toBe(401);

  const spacesToken = await bearerFor(
    app,
    passwordEmail,
    spacesPassword,
    `ticket-04-spaces-${crypto.randomUUID()}`
  );
  const persistedSession = await prisma.session.findFirst({
    orderBy: { createdAt: "desc" },
    where: { userId: passwordUser.id },
  });
  if (!persistedSession) {
    throw new Error("The replacement password session was not persisted");
  }
  const sessionDurationMs = 60 * 60 * 1000;
  expect(
    Math.abs(
      persistedSession.expiresAt.getTime() -
        persistedSession.createdAt.getTime() -
        sessionDurationMs
    )
  ).toBeLessThanOrEqual(1000);

  interface SessionBody {
    session?: { expiresAt?: string };
    user?: { mustChangePassword?: boolean };
  }
  const readSession = async (token: string): Promise<SessionBody> => {
    const response = await app.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${token}` },
      })
    );
    expect(response.status).toBe(200);
    return (await response.json()) as SessionBody;
  };
  const firstSessionRead = await readSession(spacesToken);
  const secondSessionRead = await readSession(spacesToken);
  const firstExpiresAt = firstSessionRead.session?.expiresAt;
  const secondExpiresAt = secondSessionRead.session?.expiresAt;
  expect(firstSessionRead.user?.mustChangePassword).toBe(false);
  expect(firstExpiresAt).toBe(secondExpiresAt);
  expect(firstExpiresAt).toBe(persistedSession.expiresAt.toISOString());

  const maximumPassword = "x".repeat(128);
  const maximumReplacementResponse = await replacePassword(
    app,
    spacesToken,
    spacesPassword,
    maximumPassword
  );
  expect(maximumReplacementResponse.status).toBe(200);
  expect(await maximumReplacementResponse.json()).toEqual({ ok: true });
  const invalidatedSpacesSession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${spacesToken}` },
    })
  );
  expect(invalidatedSpacesSession.status).toBe(401);

  const maximumToken = await bearerFor(
    app,
    passwordEmail,
    maximumPassword,
    `ticket-04-maximum-${crypto.randomUUID()}`
  );
  const tooLongResponse = await replacePassword(
    app,
    maximumToken,
    maximumPassword,
    "x".repeat(129)
  );
  expect(tooLongResponse.status).toBe(400);
  expect(await tooLongResponse.json()).toMatchObject({
    error: "password_too_long",
  });
  const maximumSession = await readSession(maximumToken);
  expect(maximumSession.user?.mustChangePassword).toBe(false);

  const logoutResponse = await app.handle(
    new Request("http://test.local/api/auth/sign-out", {
      headers: { Authorization: `Bearer ${maximumToken}` },
      method: "POST",
    })
  );
  expect(logoutResponse.status).toBe(200);
  const revokedLogoutSession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${maximumToken}` },
    })
  );
  expect(revokedLogoutSession.status).toBe(401);

  const throttleEmail = `ticket-04-throttle-${crypto.randomUUID()}@example.com`;
  const throttlePassword = "Ticket04-throttle-password";
  await createCredentialFixture({
    email: throttleEmail,
    name: "Ticket 04 Throttle User",
    password: throttlePassword,
  });
  const knownWrongResponse = await signIn(
    app,
    throttleEmail,
    "wrong-password",
    `ticket-04-known-${crypto.randomUUID()}`
  );
  const knownWrongBody = (await knownWrongResponse.json()) as {
    error?: string;
  };
  const unknownWrongResponse = await signIn(
    app,
    `ticket-04-unknown-${crypto.randomUUID()}@example.com`,
    "wrong-password",
    `ticket-04-unknown-${crypto.randomUUID()}`
  );
  const unknownWrongBody = (await unknownWrongResponse.json()) as {
    error?: string;
  };
  expect(knownWrongResponse.status).toBe(401);
  expect(unknownWrongResponse.status).toBe(401);
  expect(knownWrongBody.error).toBe("invalid_credentials");
  expect(unknownWrongBody.error).toBe(knownWrongBody.error);

  const parallelEmail = `ticket-04-parallel-${crypto.randomUUID()}@example.com`;
  const parallelIp = `ticket-04-parallel-ip-${crypto.randomUUID()}`;
  const parallelResponses = await Promise.all(
    Array.from({ length: 8 }, () =>
      signIn(app, parallelEmail, "wrong-password", parallelIp)
    )
  );
  const parallelStatuses = parallelResponses
    .map((response) => response.status)
    .toSorted((left, right) => left - right);
  expect(parallelStatuses).toEqual([401, 401, 401, 401, 401, 429, 429, 429]);

  const normalizedIp = `ticket-04-normalized-ip-${crypto.randomUUID()}`;
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const emailVariant =
      attempt % 2 === 0 ? ` ${throttleEmail.toUpperCase()} ` : throttleEmail;
    const response = await signIn(
      app,
      emailVariant,
      "wrong-password",
      normalizedIp
    );
    expect(response.status).toBe(401);
    expect(await response.json()).toMatchObject({
      error: "invalid_credentials",
    });
  }
  const normalizedSixthResponse = await signIn(
    app,
    ` ${throttleEmail.toUpperCase()} `,
    "wrong-password",
    normalizedIp
  );
  expect(normalizedSixthResponse.status).toBe(429);
  expect(await normalizedSixthResponse.json()).toMatchObject({
    error: "login_throttled",
  });

  const isolatedIp = `ticket-04-isolated-ip-${crypto.randomUUID()}`;
  const isolatedResponse = await signIn(
    app,
    throttleEmail,
    "wrong-password",
    isolatedIp
  );
  expect(isolatedResponse.status).toBe(401);
  expect(await isolatedResponse.json()).toMatchObject({
    error: "invalid_credentials",
  });

  const clearingIp = `ticket-04-clearing-ip-${crypto.randomUUID()}`;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const response = await signIn(
      app,
      throttleEmail,
      "wrong-password",
      clearingIp
    );
    expect(response.status).toBe(401);
  }
  const successfulLoginToken = await bearerFor(
    app,
    throttleEmail,
    throttlePassword,
    clearingIp
  );
  expect(successfulLoginToken).toBeTruthy();
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const response = await signIn(
      app,
      throttleEmail,
      "wrong-password",
      clearingIp
    );
    expect(response.status).toBe(401);
  }
  const clearedSixthResponse = await signIn(
    app,
    throttleEmail,
    "wrong-password",
    clearingIp
  );
  expect(clearedSixthResponse.status).toBe(429);
  expect(await clearedSixthResponse.json()).toMatchObject({
    error: "login_throttled",
  });
};

export interface ManagedAccountsInput {
  app: ReturnType<typeof createApp>;
  adminBearer: BootstrapAndCreationOutput["adminBearer"];
}

export interface ManagedAccountsOutput {
  managedTemporaryPassword: string;
  managedId: string;
  managedPassword: string;
  managedToken: string;
  resetTemporaryPassword: string;
  resetPassword: string;
  invalidTargetSecret: string;
}

export const runManagedAccounts = async (
  input: ManagedAccountsInput
): Promise<ManagedAccountsOutput> => {
  const { app, adminBearer } = input;

  const managedEmail = `ticket-07-managed-${crypto.randomUUID()}@example.com`;
  const managedCreateResponse = await accountRequest(
    app,
    "POST",
    "/api/admin/users",
    adminBearer,
    {
      email: ` ${managedEmail.toUpperCase()} `,
      name: "Ticket 07 Managed User",
      role: "user",
    }
  );
  expect(managedCreateResponse.status).toBe(200);
  const managedCreateBody = (await managedCreateResponse.json()) as {
    temporaryPassword?: unknown;
    user?: Record<string, unknown>;
  };
  const managedTemporaryPassword = managedCreateBody.temporaryPassword;
  const managedUser = managedCreateBody.user;
  if (
    typeof managedTemporaryPassword !== "string" ||
    !managedUser ||
    typeof managedUser.id !== "string"
  ) {
    throw new TypeError("The account creation response was malformed");
  }
  const managedId = managedUser.id;
  expect(Object.keys(managedUser).toSorted()).toEqual([
    "createdAt",
    "email",
    "enabled",
    "id",
    "mustChangePassword",
    "name",
    "role",
    "updatedAt",
  ]);
  expect(managedUser).toMatchObject({
    email: managedEmail,
    enabled: true,
    mustChangePassword: true,
    role: "user",
  });
  expect(managedTemporaryPassword).toMatch(/^[A-Za-z0-9_-]+$/u);
  expect(managedTemporaryPassword.length).toBe(32);
  const managedAccount = await prisma.account.findFirst({
    where: { providerId: "credential", userId: managedId },
  });
  expect(managedAccount?.password).not.toBe(managedTemporaryPassword);

  const cursorSeed = Array.from({ length: 21 }, (_, index) => ({
    email: `ticket-07-page-${crypto.randomUUID()}-${index}@example.com`,
    emailVerified: true,
    id: crypto.randomUUID(),
    name: `Ticket 07 Page ${index}`,
    role: "user" as const,
  }));
  await prisma.user.createMany({ data: cursorSeed });
  const firstPageResponse = await accountRequest(
    app,
    "GET",
    "/api/admin/users",
    adminBearer
  );
  expect(firstPageResponse.status).toBe(200);
  const firstPageBody = (await firstPageResponse.json()) as {
    nextCursor?: unknown;
    users?: Record<string, unknown>[];
  };
  if (
    !Array.isArray(firstPageBody.users) ||
    typeof firstPageBody.nextCursor !== "string"
  ) {
    throw new TypeError("The first account page was malformed");
  }
  expect(firstPageBody.users).toHaveLength(20);
  const firstPageIds = new Set(
    firstPageBody.users.map((pageUser) => String(pageUser.id))
  );
  const secondPageResponse = await accountRequest(
    app,
    "GET",
    `/api/admin/users?cursor=${encodeURIComponent(firstPageBody.nextCursor)}`,
    adminBearer
  );
  expect(secondPageResponse.status).toBe(200);
  const secondPageBody = (await secondPageResponse.json()) as {
    users?: Record<string, unknown>[];
  };
  if (!Array.isArray(secondPageBody.users)) {
    throw new TypeError("The second account page was malformed");
  }
  expect(
    secondPageBody.users.every(
      (pageUser) => !firstPageIds.has(String(pageUser.id))
    )
  ).toBe(true);
  const invalidCursorResponse = await accountRequest(
    app,
    "GET",
    "/api/admin/users?cursor=------------------------------------",
    adminBearer
  );
  expect(invalidCursorResponse.status).toBe(400);
  const normalizedListResponse = await accountRequest(
    app,
    "GET",
    `/api/admin/users?email=${encodeURIComponent(` ${managedEmail.toUpperCase()} `)}&role=user&enabled=true`,
    adminBearer
  );
  expect(normalizedListResponse.status).toBe(200);
  const normalizedListBody = (await normalizedListResponse.json()) as {
    users?: Record<string, unknown>[];
  };
  expect(normalizedListBody.users).toHaveLength(1);
  expect(normalizedListBody.users?.[0]).toMatchObject({
    email: managedEmail,
    id: managedId,
    role: "user",
  });
  expect(JSON.stringify(normalizedListBody)).not.toContain(
    managedTemporaryPassword
  );

  const managedTemporaryToken = await bearerFor(
    app,
    managedEmail,
    managedTemporaryPassword,
    `ticket-07-managed-temporary-${crypto.randomUUID()}`
  );
  const managedTemporarySession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${managedTemporaryToken}` },
    })
  );
  expect(managedTemporarySession.status).toBe(200);
  expect(await managedTemporarySession.json()).toMatchObject({
    user: { mustChangePassword: true },
  });
  const managedPassword = "Ticket07-managed-permanent-password";
  const managedPasswordChange = await replacePassword(
    app,
    managedTemporaryToken,
    managedTemporaryPassword,
    managedPassword
  );
  expect(managedPasswordChange.status).toBe(200);
  expect(await sessionStatus(app, managedTemporaryToken)).toBe(401);
  let managedToken = await bearerFor(
    app,
    managedEmail,
    managedPassword,
    `ticket-07-managed-live-${crypto.randomUUID()}`
  );

  const managedDisableResponse = await accountRequest(
    app,
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { enabled: false }
  );
  expect(managedDisableResponse.status).toBe(200);
  expect(await sessionStatus(app, managedToken)).toBe(401);
  const managedEnableResponse = await accountRequest(
    app,
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { enabled: true }
  );
  expect(managedEnableResponse.status).toBe(200);
  managedToken = await bearerFor(
    app,
    managedEmail,
    managedPassword,
    `ticket-07-managed-enabled-${crypto.randomUUID()}`
  );
  const changedManagedEmail = `ticket-07-managed-renamed-${crypto.randomUUID()}@example.com`;
  const managedEmailResponse = await accountRequest(
    app,
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { email: ` ${changedManagedEmail.toUpperCase()} ` }
  );
  expect(managedEmailResponse.status).toBe(200);
  expect(await sessionStatus(app, managedToken)).toBe(401);
  managedToken = await bearerFor(
    app,
    changedManagedEmail,
    managedPassword,
    `ticket-07-managed-renamed-${crypto.randomUUID()}`
  );
  const managedPromoteResponse = await accountRequest(
    app,
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { role: "admin" }
  );
  expect(managedPromoteResponse.status).toBe(200);
  expect(await sessionStatus(app, managedToken)).toBe(401);
  const promotedManagedToken = await bearerFor(
    app,
    changedManagedEmail,
    managedPassword,
    `ticket-07-managed-promoted-${crypto.randomUUID()}`
  );
  const promotedListResponse = await accountRequest(
    app,
    "GET",
    "/api/admin/users",
    promotedManagedToken
  );
  expect(promotedListResponse.status).toBe(200);
  const managedDemoteResponse = await accountRequest(
    app,
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { role: "user" }
  );
  expect(managedDemoteResponse.status).toBe(200);
  expect(await sessionStatus(app, promotedManagedToken)).toBe(401);
  managedToken = await bearerFor(
    app,
    changedManagedEmail,
    managedPassword,
    `ticket-07-managed-demoted-${crypto.randomUUID()}`
  );

  const managedResetResponse = await accountRequest(
    app,
    "POST",
    `/api/admin/users/${managedId}/password-reset`,
    adminBearer
  );
  expect(managedResetResponse.status).toBe(200);
  const managedResetBody = (await managedResetResponse.json()) as {
    temporaryPassword?: unknown;
    user?: Record<string, unknown>;
  };
  if (
    typeof managedResetBody.temporaryPassword !== "string" ||
    !managedResetBody.user
  ) {
    throw new Error("The password reset response was malformed");
  }
  const resetTemporaryPassword = managedResetBody.temporaryPassword;
  expect(resetTemporaryPassword).not.toBe(managedTemporaryPassword);
  expect(managedResetBody.user).toMatchObject({
    id: managedId,
    mustChangePassword: true,
    role: "user",
  });
  expect(await sessionStatus(app, managedToken)).toBe(401);
  const resetTemporaryToken = await bearerFor(
    app,
    changedManagedEmail,
    resetTemporaryPassword,
    `ticket-07-managed-reset-${crypto.randomUUID()}`
  );
  expect(await sessionStatus(app, resetTemporaryToken)).toBe(200);
  const resetPassword = "Ticket07-managed-reset-permanent-password";
  const resetPasswordChange = await replacePassword(
    app,
    resetTemporaryToken,
    resetTemporaryPassword,
    resetPassword
  );
  expect(resetPasswordChange.status).toBe(200);
  expect(await sessionStatus(app, resetTemporaryToken)).toBe(401);
  managedToken = await bearerFor(
    app,
    changedManagedEmail,
    resetPassword,
    `ticket-07-managed-reset-live-${crypto.randomUUID()}`
  );

  const duplicateCreateResponse = await accountRequest(
    app,
    "POST",
    "/api/admin/users",
    adminBearer,
    {
      email: ` ${changedManagedEmail.toUpperCase()} `,
      name: "Ticket 07 Duplicate",
      role: "user",
    }
  );
  expect(duplicateCreateResponse.status).toBe(409);
  expect(await duplicateCreateResponse.json()).toMatchObject({
    error: "email_in_use",
  });
  const malformedPatchResponse = await accountRequest(
    app,
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { email: "another@example.com", role: "user" }
  );
  expect(malformedPatchResponse.status).toBe(400);
  expect(await malformedPatchResponse.json()).toMatchObject({
    error: "invalid_request",
  });
  const unknownResetResponse = await accountRequest(
    app,
    "POST",
    `/api/admin/users/${crypto.randomUUID()}/password-reset`,
    adminBearer
  );
  expect(unknownResetResponse.status).toBe(404);
  const invalidTargetSecret = `ticket-07-target-secret-${crypto.randomUUID()}`;
  const invalidTargetResponse = await accountRequest(
    app,
    "POST",
    `/api/admin/users/${encodeURIComponent(invalidTargetSecret)}/password-reset`,
    adminBearer
  );
  expect(invalidTargetResponse.status).toBe(404);
  const userForbiddenResponse = await accountRequest(
    app,
    "POST",
    "/api/admin/users",
    managedToken,
    {
      email: `ticket-07-forbidden-${crypto.randomUUID()}@example.com`,
      name: "Ticket 07 Forbidden",
      role: "user",
    }
  );
  expect(userForbiddenResponse.status).toBe(403);
  expect(await userForbiddenResponse.json()).toMatchObject({
    error: "forbidden",
  });
  return {
    invalidTargetSecret,
    managedId,
    managedPassword,
    managedTemporaryPassword,
    managedToken,
    resetPassword,
    resetTemporaryPassword,
  };
};
