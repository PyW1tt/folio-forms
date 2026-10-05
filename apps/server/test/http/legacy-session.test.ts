import { test, expect, afterEach } from "bun:test";
import { createHash } from "node:crypto";

import { prisma } from "@onlyoffice/db";
import { startLegacySsoMock } from "prefill-mock/mock";
import type { LegacySsoMockServer } from "prefill-mock/mock";

import { createApp } from "../../src/app";
import {
  verifyEditorCapability,
  createEditorCapability,
} from "../../src/onlyoffice";
import { objectKey, putObject, DOCX_CONTENT_TYPE } from "../../src/storage";
import { docxFixture } from "../fixtures/documents";
import {
  createCredentialFixture,
  jsonHeaders,
  bearerFor,
} from "../fixtures/http";
import type { EditorConfigBody } from "../fixtures/http";
import {
  legacySsoCallbackUrl,
  createLegacySsoTestApp,
  startLegacySsoBrowser,
  cookieHeaderFrom,
  legacySsoWebOrigin,
  cookiePairFrom,
  legacySsoFailedLocation,
  completeAuthenticatedLegacySsoCallback,
  claimLegacySsoBearer,
  legacySsoPendingLocation,
} from "../fixtures/legacy-sso";
import type { LegacySsoHttpHandler } from "../fixtures/legacy-sso";

// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  throw new Error(
    "The HTTP application test requires DATABASE_URL for an isolated PostgreSQL database"
  );
}
const convertedDocumentKeys: string[] = [];
const app = createApp({
  legacySso: null,
  onlyOffice: {
    convertDocxToPdf: (documentKey) => {
      convertedDocumentKeys.push(documentKey);
      return Promise.resolve(new TextEncoder().encode("%PDF-test"));
    },
    forceSave: () => Promise.resolve(false),
  },
  prefillReturnUrl: "https://source.example.test/forms/return",
  requestIp: (request) => request.headers.get("x-test-ip"),
});

let legacySsoMock: LegacySsoMockServer | undefined;

afterEach(() => {
  legacySsoMock?.close();
  legacySsoMock = undefined;
});

test("linked role=user completes browser-bound SSO and receives a normal session", async () => {
  const email = `ticket-12-linked-${crypto.randomUUID()}@example.com`;
  const password = "Ticket12-local-password";
  const user = await createCredentialFixture({
    email,
    mustChangePassword: true,
    name: "Ticket 12 linked User",
    password,
  });
  const identity = {
    email,
    email_verified: false,
    name: "PDMS must not rename a linked User",
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const authorizationCode = "ticket-12+/=opaque-code";
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    authorizationCode,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    identity,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const returnTo = `/forms/${"a".repeat(32)}/fill?responseId=${crypto.randomUUID()}`;
  const browserStart = await startLegacySsoBrowser(handle, returnTo);
  expect(browserStart.startResponse.status).toBe(200);
  const startCookieHeader = cookieHeaderFrom(
    browserStart.startResponse,
    "__Host-folio-sso"
  );
  expect(startCookieHeader).toContain("HttpOnly");
  expect(startCookieHeader).toContain("Secure");
  expect(startCookieHeader).toContain("SameSite=None");
  expect(startCookieHeader).toContain("Path=/");

  const authorizeUrl = new URL(browserStart.authorizationUrl);
  expect(authorizeUrl.searchParams.get("client_id")).toBe(clientId);
  expect(authorizeUrl.searchParams.get("redirect_uri")).toBe(
    legacySsoCallbackUrl
  );
  expect(authorizeUrl.searchParams.get("code_challenge_method")).toBe("S256");
  expect(authorizeUrl.searchParams.get("state")).toMatch(
    /^[A-Za-z0-9_-]{43}$/u
  );
  expect(authorizeUrl.searchParams.get("code_challenge")).toMatch(
    /^[A-Za-z0-9_-]{43}$/u
  );
  expect(authorizeUrl.searchParams.has("returnTo")).toBe(false);
  expect(browserStart.authorizationUrl).not.toContain(clientSecret);

  const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(oldBackendResponse.status).toBe(303);
  const callbackLocation = oldBackendResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not return a callback");
  }
  const callbackUrl = new URL(callbackLocation);
  expect([...callbackUrl.searchParams.keys()].toSorted()).toEqual([
    "code",
    "state",
  ]);
  expect(callbackUrl.searchParams.get("code")).toBe(authorizationCode);
  const callbackResponse = await handle(
    new Request(callbackUrl.href, {
      headers: {
        Cookie: browserStart.preLoginCookie,
        Host: "attacker.example.test",
        "X-Forwarded-Host": "attacker.example.test",
        "X-Forwarded-Proto": "https",
      },
    })
  );
  expect(callbackResponse.status).toBe(303);
  expect(callbackResponse.headers.get("location")).toBe(
    new URL(returnTo, legacySsoWebOrigin).href
  );
  expect(callbackResponse.headers.get("referrer-policy")).toBe("no-referrer");
  expect(callbackLocation).not.toContain(clientSecret);
  const callbackPreLoginCookieHeader = cookieHeaderFrom(
    callbackResponse,
    "__Host-folio-sso"
  );
  expect(callbackPreLoginCookieHeader).toContain("Secure");
  expect(callbackPreLoginCookieHeader).toContain("SameSite=Lax");
  const callbackCookieHeader = cookieHeaderFrom(
    callbackResponse,
    "__Host-folio-sso-session"
  );
  expect(callbackCookieHeader).toContain("HttpOnly");
  expect(callbackCookieHeader).toContain("Secure");
  expect(callbackCookieHeader).toContain("SameSite=None");

  const sessionCookie = cookiePairFrom(
    callbackResponse,
    "__Host-folio-sso-session"
  );
  const wrongOriginClaim = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: sessionCookie,
        Origin: "https://attacker.example.test",
      },
      method: "POST",
    })
  );
  expect(wrongOriginClaim.status).toBe(403);
  expect(wrongOriginClaim.headers.has("set-cookie")).toBe(false);
  const claimResponse = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: sessionCookie,
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(claimResponse.status).toBe(200);
  const claimCookieHeader = cookieHeaderFrom(
    claimResponse,
    "__Host-folio-sso-session"
  );
  expect(claimCookieHeader).toContain("Secure");
  expect(claimCookieHeader).toContain("SameSite=None");
  const claimBody = await claimResponse.json();
  if (
    !claimBody ||
    typeof claimBody !== "object" ||
    Array.isArray(claimBody) ||
    !("token" in claimBody) ||
    typeof claimBody.token !== "string"
  ) {
    throw new Error("The SSO session handoff did not return a bearer");
  }
  const bearer = claimBody.token;
  expect(callbackResponse.headers.get("location")).not.toContain(bearer);
  const sessionResponse = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${bearer}` },
    })
  );
  expect(sessionResponse.status).toBe(200);
  expect(await sessionResponse.json()).toMatchObject({
    user: {
      email,
      id: user.id,
      mustChangePassword: false,
      name: "Ticket 12 linked User",
      role: "user",
    },
  });
  expect(
    await prisma.user.findUnique({ where: { id: user.id } })
  ).toMatchObject({
    email,
    emailVerified: true,
    mustChangePassword: true,
    name: "Ticket 12 linked User",
  });
  const ssoSession = await prisma.session.findFirst({
    orderBy: { createdAt: "desc" },
    where: { isSso: true, userId: user.id },
  });
  if (!ssoSession) {
    throw new Error("The linked User SSO session was not created");
  }
  const boundedSessionExpiry = new Date(Date.now() + 600_000);
  await prisma.session.update({
    data: { expiresAt: boundedSessionExpiry },
    where: { id: ssoSession.id },
  });
  const formId = crypto.randomUUID();
  const formPublicId = crypto.randomUUID().replaceAll("-", "");
  const templateDocumentKey = `ticket-12-template-${crypto.randomUUID()}`;
  const templateObjectKey = objectKey(
    "forms",
    formId,
    "published",
    crypto.randomUUID(),
    "ticket-12.docx"
  );
  const templateDocument = docxFixture(`ticket-12-${crypto.randomUUID()}`);
  await putObject(templateObjectKey, templateDocument, DOCX_CONTENT_TYPE);
  await prisma.form.create({
    data: {
      createdBy: user.id,
      id: formId,
      publicId: formPublicId,
      publishedTemplate: {
        create: {
          contentHash: createHash("sha256")
            .update(templateDocument)
            .digest("hex"),
          documentKey: templateDocumentKey,
          id: crypto.randomUUID(),
          manifest: {
            create: {
              configurationHash: createHash("sha256")
                .update("ticket-12-manifest")
                .digest("hex"),
              id: crypto.randomUUID(),
            },
          },
          objectKey: templateObjectKey,
          version: 1,
        },
      },
      status: "published",
      title: "Ticket 12 SSO capability form",
      version: 1,
    },
  });
  const startResponse = await handle(
    new Request(`https://folio.example.test/api/forms/${formPublicId}/start`, {
      headers: { Authorization: `Bearer ${bearer}` },
      method: "POST",
    })
  );
  expect(startResponse.status).toBe(200);
  const startBody = (await startResponse.json()) as {
    response?: { id?: string };
  };
  const responseId = startBody.response?.id;
  if (!responseId) {
    throw new Error("The linked User response was not started");
  }
  const editorResponse = await handle(
    new Request(
      `https://folio.example.test/api/forms/${formPublicId}/editor-config?responseId=${responseId}&action=fill`,
      { headers: { Authorization: `Bearer ${bearer}` } }
    )
  );
  expect(editorResponse.status).toBe(200);
  const editorConfig = (await editorResponse.json()) as EditorConfigBody;
  const ssoCapability = editorConfig.bridge.capabilities["save-draft"];
  if (!ssoCapability) {
    throw new Error("The linked User editor capability was not returned");
  }
  const ssoCapabilityClaims = verifyEditorCapability(ssoCapability);
  if (
    !ssoCapabilityClaims ||
    !ssoCapabilityClaims.leaseId ||
    !ssoCapabilityClaims.leaseProof
  ) {
    throw new Error("The linked User editor capability did not verify");
  }
  expect(ssoCapabilityClaims).toMatchObject({
    action: "save-draft",
    actorId: user.id,
    isSso: true,
    role: "user",
  });
  expect(ssoCapabilityClaims.sessionExpiresAt).toBe(
    Math.floor(boundedSessionExpiry.getTime() / 1000)
  );
  expect(ssoCapabilityClaims.expiresAt).toBeLessThanOrEqual(
    Math.floor(boundedSessionExpiry.getTime() / 1000)
  );
  const capabilityRequest = (token: string) =>
    handle(
      new Request(
        `https://folio.example.test/api/forms/${formPublicId}/draft`,
        {
          body: JSON.stringify({
            data: {},
            documentKey: ssoCapabilityClaims.documentKey,
            responseId,
          }),
          headers: {
            ...jsonHeaders,
            "X-Editor-Capability": token,
          },
          method: "POST",
        }
      )
    );
  const ssoDraftResponse = await capabilityRequest(ssoCapability);
  expect(ssoDraftResponse.status).toBe(202);
  const ssoDraftBody = (await ssoDraftResponse.json()) as {
    operationCapability?: string;
  };
  if (!ssoDraftBody.operationCapability) {
    throw new Error("The SSO save did not return an operation capability");
  }
  const operationClaims = verifyEditorCapability(
    ssoDraftBody.operationCapability
  );
  if (!operationClaims?.operationId) {
    throw new Error("The SSO operation capability did not verify");
  }
  expect(operationClaims).toMatchObject({
    action: "poll-operation",
    actorId: user.id,
    isSso: true,
    role: "user",
    sessionExpiresAt: Math.floor(boundedSessionExpiry.getTime() / 1000),
  });
  expect(operationClaims.expiresAt).toBeGreaterThan(
    ssoCapabilityClaims.expiresAt
  );
  expect(operationClaims.expiresAt).toBeLessThanOrEqual(
    operationClaims.issuedAt + 6 * 60
  );
  expect(operationClaims.expiresAt).toBeLessThanOrEqual(
    Math.floor(boundedSessionExpiry.getTime() / 1000)
  );
  const ssoPollResponse = await handle(
    new Request(
      `https://folio.example.test/api/operations/${operationClaims.operationId}`,
      {
        headers: {
          "X-Editor-Capability": ssoDraftBody.operationCapability,
        },
      }
    )
  );
  expect(ssoPollResponse.status).toBe(200);
  const malformedSsoCapability = createEditorCapability({
    action: ssoCapabilityClaims.action,
    actorId: user.id,
    documentKey: ssoCapabilityClaims.documentKey,
    expiresAt: ssoCapabilityClaims.expiresAt,
    formId: ssoCapabilityClaims.formId,
    isSso: "true" as unknown as boolean,
    leaseId: ssoCapabilityClaims.leaseId,
    leaseProof: ssoCapabilityClaims.leaseProof,
    role: "user",
    targetId: ssoCapabilityClaims.targetId,
    targetType: ssoCapabilityClaims.targetType,
  });
  expect(verifyEditorCapability(malformedSsoCapability)).toBeNull();
  const missingSsoExpiryCapability = createEditorCapability({
    action: ssoCapabilityClaims.action,
    actorId: user.id,
    documentKey: ssoCapabilityClaims.documentKey,
    expiresAt: ssoCapabilityClaims.expiresAt,
    formId: ssoCapabilityClaims.formId,
    isSso: true,
    leaseId: ssoCapabilityClaims.leaseId,
    leaseProof: ssoCapabilityClaims.leaseProof,
    role: "user",
    targetId: ssoCapabilityClaims.targetId,
    targetType: ssoCapabilityClaims.targetType,
  });
  expect(verifyEditorCapability(missingSsoExpiryCapability)).toBeNull();
  const missingSsoExpiryResponse = await capabilityRequest(
    missingSsoExpiryCapability
  );
  expect(missingSsoExpiryResponse.status).toBe(401);
  const admin = await createCredentialFixture({
    email: `ticket-12-admin-${crypto.randomUUID()}@example.com`,
    name: "Ticket 12 Admin",
    password: "Ticket12-admin-password",
    role: "admin",
  });
  const adminSsoCapability = createEditorCapability({
    action: ssoCapabilityClaims.action,
    actorId: admin.id,
    documentKey: ssoCapabilityClaims.documentKey,
    expiresAt: ssoCapabilityClaims.expiresAt,
    formId: ssoCapabilityClaims.formId,
    isSso: true,
    leaseId: ssoCapabilityClaims.leaseId,
    leaseProof: ssoCapabilityClaims.leaseProof,
    role: "admin",
    sessionExpiresAt: ssoCapabilityClaims.sessionExpiresAt,
    targetId: ssoCapabilityClaims.targetId,
    targetType: ssoCapabilityClaims.targetType,
  });
  const adminSsoResponse = await capabilityRequest(adminSsoCapability);
  expect(adminSsoResponse.status).toBe(401);
  await prisma.user.update({
    data: { role: "user" },
    where: { id: admin.id },
  });

  const localBearer = await bearerFor(
    app,
    email,
    password,
    `ticket-12-local-${crypto.randomUUID()}`
  );
  const localFormResponse = await handle(
    new Request(`https://folio.example.test/api/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${localBearer}` },
    })
  );
  expect(localFormResponse.status).toBe(403);
  expect(await localFormResponse.json()).toMatchObject({
    error: "password_change_required",
  });
  const localPasswordCapability = createEditorCapability({
    action: ssoCapabilityClaims.action,
    actorId: user.id,
    documentKey: ssoCapabilityClaims.documentKey,
    expiresAt: ssoCapabilityClaims.expiresAt,
    formId: ssoCapabilityClaims.formId,
    leaseId: ssoCapabilityClaims.leaseId,
    leaseProof: ssoCapabilityClaims.leaseProof,
    role: "user",
    targetId: ssoCapabilityClaims.targetId,
    targetType: ssoCapabilityClaims.targetType,
  });
  expect(verifyEditorCapability(localPasswordCapability)).toMatchObject({
    actorId: user.id,
    isSso: false,
  });
  const localPasswordResponse = await capabilityRequest(
    localPasswordCapability
  );
  expect(localPasswordResponse.status).toBe(401);

  const callbackReplay = await handle(
    new Request(callbackUrl.href, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(callbackReplay.status).toBe(303);
  expect(callbackReplay.headers.get("location")).toBe(legacySsoFailedLocation);
  const callbackReplaySessionCookieHeader = cookieHeaderFrom(
    callbackReplay,
    "__Host-folio-sso-session"
  );
  expect(callbackReplaySessionCookieHeader).toContain("Secure");
  expect(callbackReplaySessionCookieHeader).toContain("SameSite=None");
  const callbackReplayPreLoginCookieHeader = cookieHeaderFrom(
    callbackReplay,
    "__Host-folio-sso"
  );
  expect(callbackReplayPreLoginCookieHeader).toContain("Secure");
  expect(callbackReplayPreLoginCookieHeader).toContain("SameSite=Lax");
  const sessionReplay = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: sessionCookie,
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(sessionReplay.status).toBe(401);
  const sessionReplayCookieHeader = cookieHeaderFrom(
    sessionReplay,
    "__Host-folio-sso-session"
  );
  expect(sessionReplayCookieHeader).toContain("Secure");
  expect(sessionReplayCookieHeader).toContain("SameSite=None");
});

test("Legacy SSO switch requires owner confirmation and preserves current session on cancel", async () => {
  const owner = await createCredentialFixture({
    email: `ticket-15-owner-${crypto.randomUUID()}@example.com`,
    name: "Ticket 15 current User",
    password: "Ticket15-owner-password",
  });
  const target = await createCredentialFixture({
    email: `ticket-15-target-${crypto.randomUUID()}@example.com`,
    name: "Ticket 15 Legacy User",
    password: "Ticket15-target-password",
  });
  const other = await createCredentialFixture({
    email: `ticket-15-other-${crypto.randomUUID()}@example.com`,
    name: "Ticket 15 Other User",
    password: "Ticket15-other-password",
  });
  const admin = await createCredentialFixture({
    email: `ticket-15-admin-${crypto.randomUUID()}@example.com`,
    name: "Ticket 15 Admin",
    password: "Ticket15-admin-password",
    role: "admin",
  });
  const identity = {
    email: target.email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const targetSubject = identity.sub;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-15-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: target.id,
    },
  });
  const otherSubject = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: otherSubject,
      id: crypto.randomUUID(),
      providerId,
      userId: other.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    identity,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const invalidBearerStart = await handle(
    new Request("https://folio.example.test/api/legacy-sso/start", {
      body: JSON.stringify({}),
      headers: {
        ...jsonHeaders,
        Authorization: "Bearer invalid-ticket-15-token",
      },
      method: "POST",
    })
  );
  expect(invalidBearerStart.status).toBe(401);
  expect(invalidBearerStart.headers.has("set-cookie")).toBe(false);
  const ownerBearer = await bearerFor(
    app,
    owner.email,
    "Ticket15-owner-password"
  );
  const otherBearer = await bearerFor(
    app,
    other.email,
    "Ticket15-other-password"
  );
  const returnTo = `/forms/${"f".repeat(32)}/fill?responseId=${crypto.randomUUID()}`;
  const adminBearer = await bearerFor(
    app,
    admin.email,
    "Ticket15-admin-password"
  );
  const adminStart = await handle(
    new Request("https://folio.example.test/api/legacy-sso/start", {
      body: JSON.stringify({ returnTo }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${adminBearer}`,
      },
      method: "POST",
    })
  );
  expect(adminStart.status).toBe(403);
  expect(adminStart.headers.has("set-cookie")).toBe(false);
  const baselineResponses = await handle(
    new Request("https://folio.example.test/api/responses/me", {
      headers: { Authorization: `Bearer ${ownerBearer}` },
    })
  );
  const beforeResponses = await baselineResponses.json();
  const ownerSessionsBefore = await prisma.session.count({
    where: { userId: owner.id },
  });
  const targetSessionsBefore = await prisma.session.count({
    where: { userId: target.id },
  });

  const { browserStart, callback } =
    await completeAuthenticatedLegacySsoCallback(handle, ownerBearer, returnTo);
  expect(browserStart.startResponse.status).toBe(200);
  expect(callback.status).toBe(303);
  expect(callback.headers.get("location")).toBe(
    new URL("/legacy-sso/confirm", legacySsoWebOrigin).href
  );
  expect(callback.headers.get("location")).not.toContain(owner.id);
  expect(callback.headers.get("location")).not.toContain(target.id);
  expect(callback.headers.get("location")).not.toContain("code");
  expect(await prisma.session.count({ where: { userId: owner.id } })).toBe(
    ownerSessionsBefore
  );
  expect(await prisma.session.count({ where: { userId: target.id } })).toBe(
    targetSessionsBefore
  );
  const switchCookie = cookiePairFrom(callback, "__Host-folio-sso-switch");
  const switchCookieHeader = cookieHeaderFrom(
    callback,
    "__Host-folio-sso-switch"
  );
  expect(switchCookieHeader).toContain("HttpOnly");
  expect(switchCookieHeader).toContain("Secure");
  expect(switchCookieHeader).toContain("SameSite=None");
  expect(switchCookie).not.toContain(owner.id);
  expect(switchCookie).not.toContain(target.id);

  const switchRequest = async (
    bearer: string,
    path: string,
    method = "GET",
    cookie = switchCookie,
    confirmationFingerprint?: string
  ): Promise<Response> => {
    const headers = new Headers({
      Authorization: `Bearer ${bearer}`,
      Cookie: cookie,
      Origin: "https://folio.example.test",
    });
    const body =
      method === "POST" && confirmationFingerprint
        ? JSON.stringify({ confirmationFingerprint })
        : undefined;
    if (body) {
      headers.set("Content-Type", "application/json");
    }
    return await handle(
      new Request(`https://folio.example.test${path}`, {
        // oxlint-disable-next-line unicorn/no-invalid-fetch-options -- Preserve the shared GET/POST fixture: GET has an undefined body; POST carries confirmation data.
        body,
        headers,
        method,
      })
    );
  };
  const wrongOwnerRead = await switchRequest(
    otherBearer,
    "/api/legacy-sso/switch"
  );
  expect(wrongOwnerRead.status).toBe(401);
  const wrongOwnerConfirm = await switchRequest(
    otherBearer,
    "/api/legacy-sso/switch/confirm",
    "POST"
  );
  expect(wrongOwnerConfirm.status).toBe(401);
  const wrongOwnerCancel = await switchRequest(
    otherBearer,
    "/api/legacy-sso/switch/cancel",
    "POST"
  );
  expect(wrongOwnerCancel.status).toBe(401);

  const switchDetails = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch"
  );
  expect(switchDetails.status).toBe(200);
  expect(switchDetails.headers.get("cache-control")).toBe("no-store");
  const switchDetailsBody = (await switchDetails.json()) as {
    confirmationFingerprint: string;
    current: { email: string; name: string };
    legacy: { email: string; name: string };
    returnTo: string;
    sameUser: boolean;
  };
  expect(switchDetailsBody).toEqual({
    confirmationFingerprint: expect.any(String),
    current: { email: owner.email, name: "Ticket 15 current User" },
    legacy: { email: target.email, name: "Ticket 15 Legacy User" },
    returnTo,
    sameUser: false,
  });
  const cancelResponse = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch/cancel",
    "POST",
    switchCookie,
    switchDetailsBody.confirmationFingerprint
  );
  expect(cancelResponse.status).toBe(200);
  expect(cancelResponse.headers.get("set-cookie")).toContain(
    "__Host-folio-sso-switch="
  );
  expect(cancelResponse.headers.get("set-cookie")).toContain("Max-Age=0");
  expect(cancelResponse.headers.has("location")).toBe(false);
  expect(await prisma.session.count({ where: { userId: target.id } })).toBe(
    targetSessionsBefore
  );
  const ownerSessionAfterCancel = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${ownerBearer}` },
    })
  );
  expect(await ownerSessionAfterCancel.json()).toMatchObject({
    user: { id: owner.id },
  });
  const ownerResponsesAfterCancel = await handle(
    new Request("https://folio.example.test/api/responses/me", {
      headers: { Authorization: `Bearer ${ownerBearer}` },
    })
  );
  expect(await ownerResponsesAfterCancel.json()).toEqual(beforeResponses);
  const cancelledSwitchRead = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch",
    "GET"
  );
  expect(cancelledSwitchRead.status).toBe(401);

  const otherSessionsBeforeOverlap = await prisma.session.count({
    where: { userId: other.id },
  });
  const staleFlow = await completeAuthenticatedLegacySsoCallback(
    handle,
    ownerBearer,
    returnTo
  );
  const staleCookie = cookiePairFrom(
    staleFlow.callback,
    "__Host-folio-sso-switch"
  );
  const staleDetailsResponse = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch",
    "GET",
    staleCookie
  );
  const staleDetails = (await staleDetailsResponse.json()) as {
    confirmationFingerprint: string;
  };
  identity.email = other.email;
  identity.sub = otherSubject;
  const replacementFlow = await completeAuthenticatedLegacySsoCallback(
    handle,
    ownerBearer,
    returnTo
  );
  const replacementCookie = cookiePairFrom(
    replacementFlow.callback,
    "__Host-folio-sso-switch"
  );
  const staleConfirm = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch/confirm",
    "POST",
    replacementCookie,
    staleDetails.confirmationFingerprint
  );
  expect(staleConfirm.status).toBe(401);
  expect(await prisma.session.count({ where: { userId: other.id } })).toBe(
    otherSessionsBeforeOverlap
  );
  const replacementDetailsResponse = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch",
    "GET",
    replacementCookie
  );
  const replacementDetails = (await replacementDetailsResponse.json()) as {
    confirmationFingerprint: string;
    legacy: { email: string; name: string };
  };
  expect(replacementDetails.legacy).toEqual({
    email: other.email,
    name: "Ticket 15 Other User",
  });
  expect(replacementDetails.confirmationFingerprint).not.toBe(
    staleDetails.confirmationFingerprint
  );
  const replacementCancel = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch/cancel",
    "POST",
    replacementCookie,
    replacementDetails.confirmationFingerprint
  );
  expect(replacementCancel.status).toBe(200);
  const [, staleChallengeId] = staleCookie.split("=", 2);
  await prisma.verification.deleteMany({ where: { id: staleChallengeId } });
  identity.email = target.email;
  identity.sub = targetSubject;
  const pendingSubject = `legacy-${crypto.randomUUID()}`;
  identity.email = owner.email;
  identity.sub = pendingSubject;
  const pendingFlow = await completeAuthenticatedLegacySsoCallback(
    handle,
    ownerBearer,
    returnTo
  );
  const pendingLocation = new URL(
    pendingFlow.callback.headers.get("location") ?? ""
  );
  expect(pendingLocation.pathname).toBe("/login");
  expect(pendingLocation.searchParams.get("legacySso")).toBe("pending");
  expect(pendingLocation.searchParams.get("returnTo")).toBe(returnTo);
  expect(await prisma.session.count({ where: { userId: owner.id } })).toBe(
    ownerSessionsBefore
  );
  const ownerAfterPending = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${ownerBearer}` },
    })
  );
  expect(await ownerAfterPending.json()).toMatchObject({
    user: { id: owner.id },
  });
  identity.email = target.email;
  identity.sub = targetSubject;
  const confirmedFlow = await completeAuthenticatedLegacySsoCallback(
    handle,
    ownerBearer,
    returnTo
  );
  const confirmCookie = cookiePairFrom(
    confirmedFlow.callback,
    "__Host-folio-sso-switch"
  );
  const confirmDetailsResponse = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch",
    "GET",
    confirmCookie
  );
  const confirmDetails = (await confirmDetailsResponse.json()) as {
    confirmationFingerprint: string;
  };
  const confirmed = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch/confirm",
    "POST",
    confirmCookie,
    confirmDetails.confirmationFingerprint
  );
  expect(confirmed.status).toBe(200);
  expect(await confirmed.json()).toEqual({ returnTo });
  expect(confirmed.headers.get("set-cookie")).toContain(
    "__Host-folio-sso-session="
  );
  expect(confirmed.headers.get("set-cookie")).not.toContain(ownerBearer);
  expect(await prisma.session.count({ where: { userId: target.id } })).toBe(
    targetSessionsBefore + 1
  );
  const claimedBearer = await claimLegacySsoBearer(handle, confirmed);
  const targetSession = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${claimedBearer}` },
    })
  );
  expect(await targetSession.json()).toMatchObject({
    user: { id: target.id, role: "user" },
  });
  const sameUserSubject = `legacy-${crypto.randomUUID()}`;
  identity.email = owner.email;
  identity.sub = sameUserSubject;
  await prisma.account.create({
    data: {
      accountId: sameUserSubject,
      id: crypto.randomUUID(),
      providerId,
      userId: owner.id,
    },
  });
  const sameUserFlow = await completeAuthenticatedLegacySsoCallback(
    handle,
    ownerBearer,
    returnTo
  );
  const sameUserCookie = cookiePairFrom(
    sameUserFlow.callback,
    "__Host-folio-sso-switch"
  );
  const sameUserDetails = await handle(
    new Request("https://folio.example.test/api/legacy-sso/switch", {
      headers: {
        Authorization: `Bearer ${ownerBearer}`,
        Cookie: sameUserCookie,
      },
    })
  );
  const sameUserConfirmWithoutOrigin = await handle(
    new Request("https://folio.example.test/api/legacy-sso/switch/confirm", {
      headers: {
        Authorization: `Bearer ${ownerBearer}`,
        Cookie: sameUserCookie,
      },
      method: "POST",
    })
  );
  expect(sameUserConfirmWithoutOrigin.status).toBe(401);
  expect(sameUserDetails.status).toBe(200);
  const sameUserDetailsBody = (await sameUserDetails.json()) as {
    confirmationFingerprint: string;
    legacy: { email: string; name: string };
    returnTo: string;
    sameUser: boolean;
  };
  const sameUserFingerprint = sameUserDetailsBody.confirmationFingerprint;
  expect(sameUserDetailsBody).toMatchObject({
    confirmationFingerprint: expect.any(String),
    legacy: { email: owner.email, name: "Ticket 15 current User" },
    returnTo,
    sameUser: true,
  });
  const sameUserConfirm = await switchRequest(
    ownerBearer,
    "/api/legacy-sso/switch/confirm",
    "POST",
    sameUserCookie,
    sameUserFingerprint
  );
  expect(await sameUserConfirm.json()).toEqual({ returnTo });
  const sameUserBearer = await claimLegacySsoBearer(handle, sameUserConfirm);
  const sameUserSession = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${sameUserBearer}` },
    })
  );
  expect(await sameUserSession.json()).toMatchObject({
    user: { id: owner.id },
  });
});

test("Legacy SSO exchange failure and expired switch challenge preserve current bearer", async () => {
  const now = new Date();
  const owner = await createCredentialFixture({
    email: `ticket-15-safe-${crypto.randomUUID()}@example.com`,
    name: "Ticket 15 Safe User",
    password: "Ticket15-safe-password",
  });
  const target = await createCredentialFixture({
    email: `ticket-15-safe-target-${crypto.randomUUID()}@example.com`,
    name: "Ticket 15 Safe Target",
    password: "Ticket15-safe-target-password",
  });
  const identity = {
    email: target.email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-15-safe-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: target.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    clock: () => new Date(now),
    codeLifetimeMs: 60_000,
    identity,
  });
  const failedExchangeApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: new URL("/exchange-failed", legacySsoMock.exchangeUrl).href,
    providerId,
  });
  const failedExchangeHandle: LegacySsoHttpHandler = (request) =>
    failedExchangeApp.handle(request);
  const ownerBearer = await bearerFor(
    app,
    owner.email,
    "Ticket15-safe-password"
  );
  const targetSessionsBefore = await prisma.session.count({
    where: { userId: target.id },
  });
  const failedFlow = await completeAuthenticatedLegacySsoCallback(
    failedExchangeHandle,
    ownerBearer,
    `/forms/${"b".repeat(32)}/fill`
  );
  const failedLocation = new URL(
    failedFlow.callback.headers.get("location") ?? ""
  );
  expect(failedLocation.pathname).toBe("/login");
  expect(failedLocation.searchParams.get("legacySso")).toBe("failed");
  expect(failedLocation.searchParams.get("returnTo")).toBe(
    `/forms/${"b".repeat(32)}/fill`
  );
  expect(await prisma.session.count({ where: { userId: target.id } })).toBe(
    targetSessionsBefore
  );
  const ownerSessionAfterFailure = await failedExchangeHandle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${ownerBearer}` },
    })
  );
  expect(await ownerSessionAfterFailure.json()).toMatchObject({
    user: { id: owner.id },
  });

  const appNow = () => new Date(now);
  const ssoApp = createLegacySsoTestApp(
    {
      authorizeUrl: legacySsoMock.authorizeUrl,
      callbackUrl: legacySsoCallbackUrl,
      clientId,
      clientSecret,
      exchangeUrl: legacySsoMock.exchangeUrl,
      providerId,
    },
    appNow
  );
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const expiredFlow = await completeAuthenticatedLegacySsoCallback(
    handle,
    ownerBearer,
    `/forms/${"b".repeat(32)}/fill`
  );
  expect(expiredFlow.callback.headers.get("location")).toBe(
    new URL("/legacy-sso/confirm", legacySsoWebOrigin).href
  );
  const switchCookie = cookiePairFrom(
    expiredFlow.callback,
    "__Host-folio-sso-switch"
  );
  const [, challengeId] = switchCookie.split("=", 2);
  if (!challengeId) {
    throw new Error("The switch challenge cookie was empty");
  }
  await prisma.verification.update({
    data: { expiresAt: new Date(now.getTime() - 1) },
    where: { id: challengeId },
  });
  const expiredRead = await handle(
    new Request("https://folio.example.test/api/legacy-sso/switch", {
      headers: {
        Authorization: `Bearer ${ownerBearer}`,
        Cookie: switchCookie,
        Origin: "https://folio.example.test",
      },
    })
  );
  expect(expiredRead.status).toBe(401);
  const expiredConfirm = await handle(
    new Request("https://folio.example.test/api/legacy-sso/switch/confirm", {
      headers: {
        Authorization: `Bearer ${ownerBearer}`,
        Cookie: switchCookie,
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(expiredConfirm.status).toBe(401);
  expect(await prisma.session.count({ where: { userId: target.id } })).toBe(
    targetSessionsBefore
  );
  const ownerSessionAfterExpiry = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${ownerBearer}` },
    })
  );
  expect(await ownerSessionAfterExpiry.json()).toMatchObject({
    user: { id: owner.id },
  });
});

test("rejects unsafe returns and mismatched callback state or cookie", async () => {
  const email = `ticket-12-binding-${crypto.randomUUID()}@example.com`;
  const password = "Ticket12-binding-test-password";
  const user = await createCredentialFixture({
    email,
    name: "Ticket 12 binding User",
    password,
  });
  const identity = {
    email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    identity,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  for (const unsafeReturn of [
    "//attacker.example/",
    "https://attacker.example/",
    "/admin",
    `/forms/${"a".repeat(32)}/fill?access_token=secret`,
  ]) {
    const unsafeStart = await handle(
      new Request("https://folio.example.test/api/legacy-sso/start", {
        body: JSON.stringify({ returnTo: unsafeReturn }),
        headers: jsonHeaders,
        method: "POST",
      })
    );
    expect(unsafeStart.status).toBe(400);
    expect(unsafeStart.headers.has("set-cookie")).toBe(false);
  }

  const returnTo = `/forms/${"c".repeat(32)}/fill`;
  const browserStart = await startLegacySsoBrowser(handle, returnTo);
  const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(oldBackendResponse.status).toBe(303);
  const callbackLocation = oldBackendResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not return a callback");
  }
  const callbackUrl = new URL(callbackLocation);
  const wrongStateUrl = new URL(callbackUrl);
  wrongStateUrl.searchParams.set("state", "x".repeat(43));
  const wrongStateResponse = await handle(
    new Request(wrongStateUrl.href, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(wrongStateResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  const wrongCookieResponse = await handle(
    new Request(callbackUrl.href, {
      headers: { Cookie: `__Host-folio-sso=${"y".repeat(43)}` },
    })
  );
  expect(wrongCookieResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  const wrongHostUrl = new URL(callbackUrl);
  wrongHostUrl.hostname = "attacker.example.test";
  const wrongHostResponse = await handle(
    new Request(wrongHostUrl.href, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(wrongHostResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  const wrongProtocolResponse = await handle(
    new Request(callbackUrl.href, {
      headers: {
        Cookie: browserStart.preLoginCookie,
        "X-Forwarded-Proto": "http",
      },
    })
  );
  expect(wrongProtocolResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );

  const validCallback = await handle(
    new Request(callbackUrl.href, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(validCallback.status).toBe(303);
  expect(validCallback.headers.get("location")).toBe(
    new URL(returnTo, legacySsoWebOrigin).href
  );
  const enabledStatusResponse = await handle(
    new Request("https://folio.example.test/api/legacy-sso/status")
  );
  expect(await enabledStatusResponse.json()).toEqual({ enabled: true });

  const handoffGetResponse = await handle(
    new Request("https://folio.example.test/prefill/handoff")
  );
  expect(handoffGetResponse.status).toBe(405);
  const disabledStatusResponse = await app.handle(
    new Request("http://test.local/api/legacy-sso/status")
  );
  expect(await disabledStatusResponse.json()).toEqual({ enabled: false });
});

test("rejects disabled and Admin identities but queues email collisions", async () => {
  const email = `ticket-12-eligibility-${crypto.randomUUID()}@example.com`;
  const password = "Ticket12-eligibility-test-password";
  const user = await createCredentialFixture({
    email,
    name: "Ticket 12 eligible User",
    password,
  });
  const identity = {
    email,
    email_verified: false,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const linkedSubject = identity.sub;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    identity,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const expectIdentityLocation = async (
    expectedLocation: string
  ): Promise<void> => {
    const browserStart = await startLegacySsoBrowser(
      handle,
      `/forms/${"d".repeat(32)}/fill`
    );
    const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
      redirect: "manual",
    });
    expect(oldBackendResponse.status).toBe(303);
    const location = oldBackendResponse.headers.get("location");
    if (!location) {
      throw new Error("The test old backend did not return a callback");
    }
    const callbackResponse = await handle(
      new Request(location, {
        headers: { Cookie: browserStart.preLoginCookie },
      })
    );
    expect(callbackResponse.status).toBe(303);
    expect(callbackResponse.headers.get("location")).toBe(expectedLocation);
  };
  const userCountBefore = await prisma.user.count();
  identity.sub = `unlinked-${crypto.randomUUID()}`;
  await expectIdentityLocation(legacySsoPendingLocation);
  identity.sub = linkedSubject;
  await prisma.user.update({
    data: { enabled: false },
    where: { id: user.id },
  });
  await expectIdentityLocation(legacySsoFailedLocation);
  await prisma.user.update({
    data: { enabled: true, role: "admin" },
    where: { id: user.id },
  });
  await expectIdentityLocation(legacySsoFailedLocation);
  await prisma.user.update({
    data: { role: "user" },
    where: { id: user.id },
  });
  expect(await prisma.user.count()).toBe(userCountBefore);
  expect(await prisma.session.count({ where: { userId: user.id } })).toBe(0);
});

test("expires browser-bound SSO transactions and session handoffs", async () => {
  let appNow = new Date();
  const backendNow = new Date();
  const email = `ticket-12-expiry-${crypto.randomUUID()}@example.com`;
  const user = await createCredentialFixture({
    email,
    name: "Ticket 12 expiry User",
    password: "Ticket12-expiry-test-password",
  });
  const identity = {
    email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    clock: () => new Date(backendNow),
    codeLifetimeMs: 60_000,
    identity,
  });
  const ssoApp = createLegacySsoTestApp(
    {
      authorizeUrl: legacySsoMock.authorizeUrl,
      callbackUrl: legacySsoCallbackUrl,
      clientId,
      clientSecret,
      exchangeUrl: legacySsoMock.exchangeUrl,
      providerId,
    },
    () => new Date(appNow)
  );
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const returnTo = `/forms/${"e".repeat(32)}/fill`;
  const expiredStart = await startLegacySsoBrowser(handle, returnTo);
  const expiredAuthorization = await fetch(expiredStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(expiredAuthorization.status).toBe(303);
  const expiredCallbackUrl = expiredAuthorization.headers.get("location");
  if (!expiredCallbackUrl) {
    throw new Error("The test old backend did not return an expiring code");
  }
  appNow = new Date(appNow.getTime() + 5 * 60_000 + 1);
  const expiredCallback = await handle(
    new Request(expiredCallbackUrl, {
      headers: { Cookie: expiredStart.preLoginCookie },
    })
  );
  expect(expiredCallback.headers.get("location")).toBe(legacySsoFailedLocation);

  const validStart = await startLegacySsoBrowser(handle, returnTo);
  const validAuthorization = await fetch(validStart.authorizationUrl, {
    redirect: "manual",
  });
  const validCallbackUrl = validAuthorization.headers.get("location");
  if (!validCallbackUrl) {
    throw new Error("The test old backend did not return a valid code");
  }
  const validCallback = await handle(
    new Request(validCallbackUrl, {
      headers: { Cookie: validStart.preLoginCookie },
    })
  );
  expect(validCallback.headers.get("location")).toBe(
    new URL(returnTo, legacySsoWebOrigin).href
  );
  const sessionCookie = cookiePairFrom(
    validCallback,
    "__Host-folio-sso-session"
  );
  appNow = new Date(appNow.getTime() + 60_001);
  const expiredSession = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: sessionCookie,
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(expiredSession.status).toBe(401);
});

test("old-backend codes bind client, callback, verifier, expiry, and single use", async () => {
  let now = new Date();
  const callbackUrl = legacySsoCallbackUrl;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const identity = {
    email: `ticket-12-mock-${crypto.randomUUID()}@example.com`,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const legacySsoServer = startLegacySsoMock({
    callbackUrl,
    clientId,
    clientSecret,
    clock: () => new Date(now),
    codeLifetimeMs: 60_000,
    exchangeResponse: () => Response.json(identity),
    identity,
  });
  legacySsoMock = legacySsoServer;
  const verifier = "v".repeat(43);
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  const state = "s".repeat(43);
  const authorize = new URL(legacySsoServer.authorizeUrl);
  authorize.searchParams.set("client_id", clientId);
  authorize.searchParams.set("redirect_uri", callbackUrl);
  authorize.searchParams.set("state", state);
  authorize.searchParams.set("code_challenge", challenge);
  authorize.searchParams.set("code_challenge_method", "S256");

  const wrongClientAuthorize = new URL(authorize);
  wrongClientAuthorize.searchParams.set("client_id", "other-client");
  const wrongClientAuthorizeResponse = await fetch(wrongClientAuthorize, {
    redirect: "manual",
  });
  expect(wrongClientAuthorizeResponse.status).toBe(400);
  const wrongCallbackAuthorize = new URL(authorize);
  wrongCallbackAuthorize.searchParams.set(
    "redirect_uri",
    "https://attacker.example.test/callback"
  );
  const wrongCallbackAuthorizeResponse = await fetch(wrongCallbackAuthorize, {
    redirect: "manual",
  });
  expect(wrongCallbackAuthorizeResponse.status).toBe(400);
  const authorizeResponse = await fetch(authorize, { redirect: "manual" });
  expect(authorizeResponse.status).toBe(303);
  const callbackLocation = authorizeResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not issue an authorization code");
  }
  const code = new URL(callbackLocation).searchParams.get("code");
  if (!code) {
    throw new Error("The test old backend callback omitted its code");
  }

  const exchange = async ({
    clientId: exchangeClientId = clientId,
    codeVerifier = verifier,
    redirectUri = callbackUrl,
  }: {
    clientId?: string;
    codeVerifier?: string;
    redirectUri?: string;
  } = {}) =>
    await fetch(legacySsoServer.exchangeUrl, {
      body: new URLSearchParams({
        client_id: exchangeClientId,
        code,
        code_verifier: codeVerifier,
        grant_type: "authorization_code",
        redirect_uri: redirectUri,
      }),
      headers: {
        Authorization: `Basic ${Buffer.from(
          `${clientId}:${clientSecret}`
        ).toString("base64")}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      method: "POST",
    });
  const wrongClientExchange = await exchange({ clientId: "other-client" });
  expect(wrongClientExchange.status).toBe(400);
  const wrongRedirectExchange = await exchange({
    redirectUri: "https://attacker.example.test/callback",
  });
  expect(wrongRedirectExchange.status).toBe(400);
  const wrongVerifierExchange = await exchange({
    codeVerifier: "w".repeat(43),
  });
  expect(wrongVerifierExchange.status).toBe(400);
  const concurrentExchanges = await Promise.all([exchange(), exchange()]);
  expect(
    concurrentExchanges.map((response) => response.status).toSorted()
  ).toEqual([200, 400]);
  const replayedExchange = await exchange();
  expect(replayedExchange.status).toBe(400);

  const expiredAuthorization = await fetch(authorize, { redirect: "manual" });
  const expiredLocation = expiredAuthorization.headers.get("location");
  if (!expiredLocation) {
    throw new Error("The test old backend did not issue an expiring code");
  }
  const expiredCode = new URL(expiredLocation).searchParams.get("code");
  if (!expiredCode) {
    throw new Error("The test old backend callback omitted its expiring code");
  }
  now = new Date(now.getTime() + 60_000);
  const expiredResponse = await fetch(legacySsoServer.exchangeUrl, {
    body: new URLSearchParams({
      client_id: clientId,
      code: expiredCode,
      code_verifier: verifier,
      grant_type: "authorization_code",
      redirect_uri: callbackUrl,
    }),
    headers: {
      Authorization: `Basic ${Buffer.from(
        `${clientId}:${clientSecret}`
      ).toString("base64")}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    method: "POST",
  });
  expect(expiredResponse.status).toBe(400);
});
