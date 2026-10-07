import { test, expect, afterEach } from "bun:test";
import { createHash } from "node:crypto";

import { auth } from "@onlyoffice/auth";
import { prisma } from "@onlyoffice/db";
import { unzipSync } from "fflate";
import { startLegacySsoMock } from "prefill-mock/mock";
import type { LegacySsoMockServer } from "prefill-mock/mock";

import { createApp } from "../../src/app";
import { verifyEditorCapability } from "../../src/onlyoffice";
import {
  objectKey,
  putObject,
  DOCX_CONTENT_TYPE,
  readObject,
} from "../../src/storage";
import {
  docxFixture,
  docxXmlFixture,
  contentControlDocument,
  contentControl,
} from "../fixtures/documents";
import {
  jsonHeaders,
  createCredentialFixture,
  bearerFor,
  waitForOperation,
} from "../fixtures/http";
import {
  legacySsoCallbackUrl,
  createLegacySsoTestApp,
  completeLegacySsoCallback,
  legacySsoWebOrigin,
  claimLegacySsoBearer,
  legacySsoFailedLocation,
  cookiePairFrom,
  startLegacySsoBrowser,
  completeAuthenticatedLegacySsoCallback,
  legacySsoPendingLocation,
  createPdmsReviewFixture,
  legacyAccountLinkSnapshot,
  submitLegacyAccountLinkReview,
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

const retainLegacySsoMock = (mock: LegacySsoMockServer): void => {
  legacySsoMock = mock;
};

afterEach(() => {
  legacySsoMock?.close();
  legacySsoMock = undefined;
});

const pdmsInitialNameCases: {
  label: string;
  name: unknown;
  expectedName?: string;
}[] = [
  {
    expectedName: "PDMS Test User",
    label: "trimmed name",
    name: "  PDMS Test User  ",
  },
  {
    expectedName: "n".repeat(120),
    label: "120-unit name",
    name: "n".repeat(120),
  },
  {
    expectedName: "\u{1F600}".repeat(60),
    label: "120 UTF-16-unit name",
    name: "\u{1F600}".repeat(60),
  },
  { label: "absent name", name: undefined },
  { label: "nonstring name", name: 123 },
  { label: "null name", name: null },
  { label: "blank name", name: "   " },
  { label: "121-unit name", name: "n".repeat(121) },
  { label: "over-limit UTF-16 name", name: "\u{1F600}".repeat(61) },
  { label: "newline name", name: "\nPDMS Test User\n" },
  { label: "tab name", name: "\tPDMS Test User\t" },
  { label: "NUL name", name: "PDMS\u0000Test User" },
  { label: "DEL name", name: "PDMS\u007FTest User" },
  { label: "C1-control name", name: "PDMS\u0085Test User" },
];

const registerPdmsNameCase = (
  nameCase: (typeof pdmsInitialNameCases)[number]
): void => {
  const currentNameCase = nameCase;
  test(`PDMS truthful identity provisions ${currentNameCase.label} without mailbox verification or provider privileges`, async () => {
    let assertedName = currentNameCase.name;
    const identity = {
      email: `PDMS-Name-${crypto.randomUUID()}@EXAMPLE.COM`,
      email_verified: false,
      sub: "12345",
    };
    const email = identity.email.toLowerCase();
    const clientId = `folio-${crypto.randomUUID()}`;
    const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
    const providerId = `pdms-${crypto.randomUUID()}`;
    legacySsoMock = startLegacySsoMock({
      callbackUrl: legacySsoCallbackUrl,
      clientId,
      clientSecret,
      codeLifetimeMs: 60_000,
      exchangeResponse: () =>
        Response.json({
          ...identity,
          access_token: "not-a-folio-credential",
          name: assertedName,
          role: "admin",
        }),
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
    const callback = await completeLegacySsoCallback(handle);
    expect(callback.status).toBe(303);
    expect(callback.headers.get("location")).toBe(
      new URL("/dashboard", legacySsoWebOrigin).href
    );
    const bearer = await claimLegacySsoBearer(handle, callback);
    const user = await prisma.user.findUniqueOrThrow({ where: { email } });
    expect(user).toMatchObject({
      email,
      emailVerified: false,
      enabled: true,
      mustChangePassword: false,
      name: currentNameCase.expectedName ?? email,
      role: "user",
    });
    expect(await prisma.user.count({ where: { email } })).toBe(1);
    expect(
      await prisma.account.findMany({ where: { userId: user.id } })
    ).toEqual([
      expect.objectContaining({
        accessToken: null,
        accessTokenExpiresAt: null,
        accountId: identity.sub,
        idToken: null,
        issuer: providerId,
        password: null,
        providerId,
        refreshToken: null,
        refreshTokenExpiresAt: null,
        userId: user.id,
      }),
    ]);
    const session = await handle(
      new Request("https://folio.example.test/api/session", {
        headers: { Authorization: `Bearer ${bearer}` },
      })
    );
    expect(session.status).toBe(200);
    expect(await session.json()).toMatchObject({
      user: {
        email,
        id: user.id,
        name: currentNameCase.expectedName ?? email,
        role: "user",
      },
    });
    const forbiddenAdmin = await handle(
      new Request("https://folio.example.test/api/admin/users", {
        headers: { Authorization: `Bearer ${bearer}` },
      })
    );
    expect(forbiddenAdmin.status).toBe(403);

    identity.email = `PDMS-Changed-${crypto.randomUUID()}@EXAMPLE.COM`;
    identity.email_verified = true;
    assertedName = "A later PDMS name must not replace the local profile";
    const repeatedCallback = await completeLegacySsoCallback(handle);
    const repeatedBearer = await claimLegacySsoBearer(handle, repeatedCallback);
    const repeatedSession = await handle(
      new Request("https://folio.example.test/api/session", {
        headers: { Authorization: `Bearer ${repeatedBearer}` },
      })
    );
    expect(repeatedSession.status).toBe(200);
    expect(await repeatedSession.json()).toMatchObject({
      user: { email, id: user.id, name: user.name, role: "user" },
    });
    expect(
      await prisma.user.findUnique({ where: { id: user.id } })
    ).toMatchObject({
      email,
      emailVerified: false,
      mustChangePassword: false,
      name: user.name,
    });
  });
};

for (const nameCase of pdmsInitialNameCases) {
  registerPdmsNameCase(nameCase);
}

const pdmsInvalidExchangeCases: {
  label: string;
  response: (identity: {
    email: string;
    email_verified: boolean;
    sub: string;
  }) => Response | Promise<Response>;
}[] = [
  {
    label: "missing sub",
    response: (identity) => Response.json({ ...identity, sub: undefined }),
  },
  {
    label: "subject-only payload",
    response: (identity) =>
      Response.json({ ...identity, sub: undefined, subject: identity.sub }),
  },
  {
    label: "nonstring sub",
    response: (identity) => Response.json({ ...identity, sub: 12_345 }),
  },
  {
    label: "empty sub",
    response: (identity) => Response.json({ ...identity, sub: "" }),
  },
  {
    label: "over-limit sub",
    response: (identity) =>
      Response.json({ ...identity, sub: "s".repeat(256) }),
  },
  {
    label: "untrimmed sub",
    response: (identity) => Response.json({ ...identity, sub: " 12345" }),
  },
  {
    label: "control-character sub",
    response: (identity) => Response.json({ ...identity, sub: "123\u007F45" }),
  },
  {
    label: "missing email",
    response: (identity) => Response.json({ ...identity, email: undefined }),
  },
  {
    label: "nonstring email",
    response: (identity) => Response.json({ ...identity, email: 123 }),
  },
  {
    label: "malformed email",
    response: (identity) =>
      Response.json({ ...identity, email: "not-an-email" }),
  },
  {
    label: "over-limit email",
    response: (identity) =>
      Response.json({ ...identity, email: `${"e".repeat(243)}@example.com` }),
  },
  {
    label: "missing email_verified",
    response: (identity) =>
      Response.json({ ...identity, email_verified: undefined }),
  },
  {
    label: "string email_verified",
    response: (identity) =>
      Response.json({ ...identity, email_verified: "false" }),
  },
  {
    label: "numeric email_verified",
    response: (identity) => Response.json({ ...identity, email_verified: 1 }),
  },
  {
    label: "null email_verified",
    response: (identity) =>
      Response.json({ ...identity, email_verified: null }),
  },
  { label: "array payload", response: (identity) => Response.json([identity]) },
  { label: "null payload", response: () => Response.json(null) },
  {
    label: "malformed JSON",
    response: () => new Response("{", { headers: jsonHeaders }),
  },
  {
    label: "non-JSON content",
    response: (identity) =>
      Response.json(identity, { headers: { "Content-Type": "text/plain" } }),
  },
  {
    label: "oversized streamed JSON",
    response: (identity) => {
      const body = new TextEncoder().encode(
        JSON.stringify({ ...identity, padding: "x".repeat(8192) })
      );
      return new Response(
        new ReadableStream<Uint8Array>({
          start(controller) {
            controller.enqueue(body.subarray(0, 4096));
            controller.enqueue(body.subarray(4096));
            controller.close();
          },
        }),
        { headers: jsonHeaders }
      );
    },
  },
  {
    label: "exchange failure",
    response: (identity) => Response.json(identity, { status: 503 }),
  },
  {
    label: "redirecting exchange",
    response: () =>
      Response.redirect("https://attacker.example.test/exchange", 302),
  },
  {
    label: "exchange timeout",
    // Exercise the real fetch AbortSignal.timeout(5000); fake timers cannot drive the platform timeout.
    response: () => Promise.withResolvers<Response>().promise,
  },
];

const registerPdmsInvalidExchangeCase = (
  exchangeCase: (typeof pdmsInvalidExchangeCases)[number]
): void => {
  const currentExchangeCase = exchangeCase;
  test(`PDMS denies ${currentExchangeCase.label} without provisioning or granting a Session`, async () => {
    const identity = {
      email: `pdms-invalid-${crypto.randomUUID()}@example.com`,
      email_verified: false,
      sub: `pdms-${crypto.randomUUID()}`,
    };
    const clientId = `folio-${crypto.randomUUID()}`;
    const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
    const providerId = `pdms-${crypto.randomUUID()}`;
    legacySsoMock = startLegacySsoMock({
      callbackUrl: legacySsoCallbackUrl,
      clientId,
      clientSecret,
      codeLifetimeMs: 60_000,
      exchangeResponse: () => currentExchangeCase.response(identity),
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
    const initialCounts = {
      accounts: await prisma.account.count(),
      requests: await prisma.legacyAccountLinkRequest.count(),
      sessions: await prisma.session.count(),
      users: await prisma.user.count(),
    };
    const callback = await completeLegacySsoCallback(handle);
    expect(callback.status).toBe(303);
    expect(callback.headers.get("location")).toBe(legacySsoFailedLocation);
    expect(cookiePairFrom(callback, "__Host-folio-sso-session")).toBe(
      "__Host-folio-sso-session="
    );
    const claim = await handle(
      new Request("https://folio.example.test/api/legacy-sso/session", {
        headers: {
          Cookie: "__Host-folio-sso-session=",
          Origin: "https://folio.example.test",
        },
        method: "POST",
      })
    );
    expect(claim.status).toBe(401);
    expect(await prisma.user.count()).toBe(initialCounts.users);
    expect(await prisma.account.count()).toBe(initialCounts.accounts);
    expect(await prisma.session.count()).toBe(initialCounts.sessions);
    expect(await prisma.legacyAccountLinkRequest.count()).toBe(
      initialCounts.requests
    );
  }, 15_000);
};

for (const exchangeCase of pdmsInvalidExchangeCases) {
  registerPdmsInvalidExchangeCase(exchangeCase);
}

test("PDMS accepts JSON at the 8 KiB boundary and a 255-unit stable subject", async () => {
  const identity = {
    email: `pdms-boundary-${crypto.randomUUID()}@example.com`,
    email_verified: true,
    sub: "s".repeat(255),
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
  const providerId = `pdms-${crypto.randomUUID()}`;
  const unpadded = JSON.stringify({ ...identity, padding: "" });
  const body = JSON.stringify({
    ...identity,
    padding: "x".repeat(8192 - unpadded.length),
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    exchangeResponse: () =>
      new Response(body, {
        headers: { "Content-Type": "application/json; charset=utf-8" },
      }),
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
  const callback = await completeLegacySsoCallback(handle);
  expect(callback.headers.get("location")).toBe(
    new URL("/dashboard", legacySsoWebOrigin).href
  );
  const bearer = await claimLegacySsoBearer(handle, callback);
  const user = await prisma.user.findUniqueOrThrow({
    where: { email: identity.email },
  });
  expect(user).toMatchObject({
    emailVerified: false,
    name: identity.email,
    role: "user",
  });
  const session = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${bearer}` },
    })
  );
  expect(session.status).toBe(200);
  expect(await session.json()).toMatchObject({
    user: { id: user.id, role: "user" },
  });
});

test("PDMS rejects failed Basic authentication without provisioning or granting a Session", async () => {
  const identity = {
    email: `pdms-basic-${crypto.randomUUID()}@example.com`,
    email_verified: false,
    sub: `pdms-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
  const providerId = `pdms-${crypto.randomUUID()}`;
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    exchangeResponse: () => Response.json(identity),
    identity,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret: `${clientSecret}-wrong`,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const initialCounts = {
    accounts: await prisma.account.count(),
    sessions: await prisma.session.count(),
    users: await prisma.user.count(),
  };
  const callback = await completeLegacySsoCallback((request) =>
    ssoApp.handle(request)
  );
  expect(callback.headers.get("location")).toBe(legacySsoFailedLocation);
  expect(cookiePairFrom(callback, "__Host-folio-sso-session")).toBe(
    "__Host-folio-sso-session="
  );
  expect(await prisma.user.count()).toBe(initialCounts.users);
  expect(await prisma.account.count()).toBe(initialCounts.accounts);
  expect(await prisma.session.count()).toBe(initialCounts.sessions);
});

test("PDMS changed-email completed mapping retains historical linked ownership and local profile", async () => {
  const owner = await createCredentialFixture({
    email: `pdms-owner-${crypto.randomUUID()}@example.com`,
    mustChangePassword: true,
    name: "Original local profile",
    password: "PDMS-local-temporary-password",
  });
  const other = await createCredentialFixture({
    email: `pdms-other-${crypto.randomUUID()}@example.com`,
    name: "Other local profile",
    password: "PDMS-other-local-password",
  });
  const identity = {
    email: `pdms-free-${crypto.randomUUID()}@example.com`,
    email_verified: false,
    name: "Changed PDMS profile",
    sub: `pdms-${crypto.randomUUID()}`,
  };
  const freeEmail = identity.email;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
  const providerId = `pdms-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      issuer: providerId,
      providerId,
      userId: owner.id,
    },
  });
  await prisma.legacyAccountLinkRequest.create({
    data: {
      email: owner.email,
      providerId,
      reviewedAt: new Date(),
      reviewedGeneration: 1n,
      status: "linked",
      subject: identity.sub,
      userId: owner.id,
    },
  });
  const credentialBefore = await prisma.account.findFirstOrThrow({
    where: { providerId: "credential", userId: owner.id },
  });
  const formId = crypto.randomUUID();
  const publishedTemplateId = crypto.randomUUID();
  const document = docxFixture(`pdms-ownership-${crypto.randomUUID()}`);
  const templateObjectKey = objectKey(
    "forms",
    formId,
    "published",
    crypto.randomUUID(),
    "pdms.docx"
  );
  await putObject(templateObjectKey, document, DOCX_CONTENT_TYPE);
  await prisma.form.create({
    data: {
      createdBy: owner.id,
      id: formId,
      publicId: crypto.randomUUID().replaceAll("-", ""),
      publishedTemplate: {
        create: {
          contentHash: createHash("sha256").update(document).digest("hex"),
          documentKey: `pdms-template-${crypto.randomUUID()}`,
          id: publishedTemplateId,
          objectKey: templateObjectKey,
          version: 1,
        },
      },
      status: "published",
      title: "PDMS completed mapping ownership",
      version: 1,
    },
  });
  const responses = [];
  for (const responseOwner of [owner, other]) {
    const responseId = crypto.randomUUID();
    const draftObjectKey = objectKey(
      "responses",
      responseId,
      "draft",
      crypto.randomUUID(),
      "pdms.docx"
    );
    await putObject(draftObjectKey, document, DOCX_CONTENT_TYPE);
    responses.push(
      await prisma.response.create({
        data: {
          draftData: {
            full_name:
              responseOwner.id === owner.id
                ? "Original owner values"
                : "Other owner private values",
          },
          draftDocumentKey: `pdms-draft-${crypto.randomUUID()}`,
          draftObjectKey,
          formId,
          id: responseId,
          publishedTemplateId,
          publishedVersion: 1,
          userId: responseOwner.id,
        },
      })
    );
  }
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
  // Reserve a generation so real callbacks are newer than the historical cutoff, even in an empty database.
  await startLegacySsoBrowser(handle);
  for (const emailVerified of [true, false]) {
    await prisma.user.update({
      data: { emailVerified },
      where: { id: owner.id },
    });
    for (const email of [freeEmail, other.email]) {
      identity.email = email;
      const callback = await completeLegacySsoCallback(handle);
      expect(callback.status).toBe(303);
      expect(callback.headers.get("location")).toBe(
        new URL("/dashboard", legacySsoWebOrigin).href
      );
      const bearer = await claimLegacySsoBearer(handle, callback);
      const session = await handle(
        new Request("https://folio.example.test/api/session", {
          headers: { Authorization: `Bearer ${bearer}` },
        })
      );
      expect(session.status).toBe(200);
      expect(await session.json()).toMatchObject({
        user: {
          email: owner.email,
          id: owner.id,
          name: "Original local profile",
          role: "user",
        },
      });
      expect(
        await prisma.user.findUnique({ where: { id: owner.id } })
      ).toMatchObject({
        email: owner.email,
        emailVerified,
        enabled: true,
        mustChangePassword: true,
        name: "Original local profile",
        role: "user",
      });
      expect(
        await prisma.account.findUnique({
          where: {
            providerId_accountId: { accountId: identity.sub, providerId },
          },
        })
      ).toMatchObject({ userId: owner.id });
      expect(
        await prisma.account.findUnique({ where: { id: credentialBefore.id } })
      ).toEqual(credentialBefore);
      for (const response of responses) {
        const isOwned = response.userId === owner.id;
        const exportedJson = await handle(
          new Request(
            `https://folio.example.test/api/responses/${response.id}/draft/json`,
            {
              headers: { Authorization: `Bearer ${bearer}` },
            }
          )
        );
        expect(exportedJson.status).toBe(isOwned ? 200 : 403);
        if (isOwned) {
          expect(await exportedJson.json()).toEqual({
            full_name: "Original owner values",
          });
        }
        const exportedDocx = await handle(
          new Request(
            `https://folio.example.test/api/responses/${response.id}/draft/docx`,
            {
              headers: { Authorization: `Bearer ${bearer}` },
            }
          )
        );
        expect(exportedDocx.status).toBe(isOwned ? 200 : 403);
        if (isOwned) {
          expect([...new Uint8Array(await exportedDocx.arrayBuffer())]).toEqual(
            [...document]
          );
        }
      }
    }
  }
  expect(
    await prisma.user.findUnique({ where: { id: other.id } })
  ).toMatchObject({
    email: other.email,
    name: "Other local profile",
  });
});

test("PDMS completed mapping ignores review metadata but retains immutable stale-attempt cutoff", async () => {
  const owner = await createCredentialFixture({
    email: `pdms-mapped-${crypto.randomUUID()}@example.com`,
    name: "Mapped owner",
    password: "PDMS-mapped-owner-password",
  });
  const candidate = await createCredentialFixture({
    email: `pdms-history-${crypto.randomUUID()}@example.com`,
    name: "Historical candidate",
    password: "PDMS-historical-candidate-password",
  });
  const identity = {
    email: `pdms-current-${crypto.randomUUID()}@example.com`,
    email_verified: false,
    sub: `pdms-${crypto.randomUUID()}`,
  };
  const providerId = `pdms-${crypto.randomUUID()}`;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
  const mapping = await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      issuer: providerId,
      providerId,
      userId: owner.id,
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
  const noReviewCallback = await completeLegacySsoCallback(handle);
  expect(noReviewCallback.headers.get("location")).toBe(
    new URL("/dashboard", legacySsoWebOrigin).href
  );
  const noReviewBearer = await claimLegacySsoBearer(handle, noReviewCallback);
  const noReviewSession = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${noReviewBearer}` },
    })
  );
  expect(await noReviewSession.json()).toMatchObject({
    user: { id: owner.id },
  });
  const request = await prisma.legacyAccountLinkRequest.create({
    data: {
      email: candidate.email,
      providerId,
      status: "linked",
      subject: identity.sub,
      userId: candidate.id,
    },
  });
  for (const status of ["pending", "approved", "rejected", "linked"] as const) {
    await prisma.legacyAccountLinkRequest.update({
      data: { status },
      where: { id: request.id },
    });
    const callback = await completeLegacySsoCallback(handle);
    expect(callback.headers.get("location")).toBe(
      new URL("/dashboard", legacySsoWebOrigin).href
    );
    const bearer = await claimLegacySsoBearer(handle, callback);
    const session = await handle(
      new Request("https://folio.example.test/api/session", {
        headers: { Authorization: `Bearer ${bearer}` },
      })
    );
    expect(session.status).toBe(200);
    expect(await session.json()).toMatchObject({
      user: { email: owner.email, id: owner.id },
    });
    expect(
      await prisma.legacyAccountLinkRequest.findUnique({
        where: { id: request.id },
      })
    ).toMatchObject({
      email: candidate.email,
      reviewedGeneration: null,
      status,
      userId: candidate.id,
    });
  }
  const browsers = [
    await startLegacySsoBrowser(handle),
    await startLegacySsoBrowser(handle),
  ] as const;
  const [, equalBrowser] = browsers;
  const [, transactionId] = equalBrowser.preLoginCookie.split("=", 2);
  if (!transactionId) {
    throw new Error("Equal-generation SSO transaction is missing");
  }
  const transaction = await prisma.verification.findUniqueOrThrow({
    where: { id: transactionId },
  });
  const { generation } = JSON.parse(transaction.value) as {
    generation: string;
  };
  await prisma.legacyAccountLinkRequest.update({
    data: { reviewedGeneration: BigInt(generation) },
    where: { id: request.id },
  });
  const fencedRequest = await prisma.legacyAccountLinkRequest.findUniqueOrThrow(
    { where: { id: request.id } }
  );
  const sessionCountBefore = await prisma.session.count({
    where: { userId: owner.id },
  });
  for (const browser of browsers) {
    const authorization = await fetch(browser.authorizationUrl, {
      redirect: "manual",
    });
    const location = authorization.headers.get("location");
    if (!location) {
      throw new Error("Fenced SSO authorization did not return a callback");
    }
    const callback = await handle(
      new Request(location, { headers: { Cookie: browser.preLoginCookie } })
    );
    expect(callback.headers.get("location")).toBe(legacySsoFailedLocation);
    expect(cookiePairFrom(callback, "__Host-folio-sso-session")).toBe(
      "__Host-folio-sso-session="
    );
  }
  expect(await prisma.session.count({ where: { userId: owner.id } })).toBe(
    sessionCountBefore
  );
  expect(
    await prisma.legacyAccountLinkRequest.findUnique({
      where: { id: request.id },
    })
  ).toEqual(fencedRequest);
  expect(
    await prisma.account.findUnique({ where: { id: mapping.id } })
  ).toEqual(mapping);
  const newerCallback = await completeLegacySsoCallback(handle);
  expect(newerCallback.headers.get("location")).toBe(
    new URL("/dashboard", legacySsoWebOrigin).href
  );
  const newerBearer = await claimLegacySsoBearer(handle, newerCallback);
  const newerSession = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${newerBearer}` },
    })
  );
  expect(await newerSession.json()).toMatchObject({ user: { id: owner.id } });
  await prisma.account.delete({ where: { id: mapping.id } });
  identity.email = candidate.email;
  const sessionCountAfterFresh = await prisma.session.count({
    where: { userId: { in: [owner.id, candidate.id] } },
  });
  const missingMappingCallback = await completeLegacySsoCallback(handle);
  expect(missingMappingCallback.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  expect(
    await prisma.account.findUnique({
      where: { providerId_accountId: { accountId: identity.sub, providerId } },
    })
  ).toBeNull();
  expect(
    await prisma.session.count({
      where: { userId: { in: [owner.id, candidate.id] } },
    })
  ).toBe(sessionCountAfterFresh);
});

test("PDMS mapped owner eligibility is rechecked at callback, claim, and switch confirmation", async () => {
  const owner = await createCredentialFixture({
    email: `pdms-eligible-${crypto.randomUUID()}@example.com`,
    name: "Mapped target",
    password: "PDMS-eligible-target-password",
  });
  const initiatorPassword = "PDMS-switch-initiator-password";
  const initiator = await createCredentialFixture({
    email: `pdms-initiator-${crypto.randomUUID()}@example.com`,
    name: "Switch initiator",
    password: initiatorPassword,
  });
  const initiatingBearer = await bearerFor(
    app,
    initiator.email,
    initiatorPassword
  );
  const identity = {
    email: `pdms-changed-${crypto.randomUUID()}@example.com`,
    email_verified: false,
    sub: `pdms-${crypto.randomUUID()}`,
  };
  const providerId = `pdms-${crypto.randomUUID()}`;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      issuer: providerId,
      providerId,
      userId: owner.id,
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
  for (const ineligible of [
    { enabled: false, role: "user" },
    { enabled: true, role: "admin" },
  ] as const) {
    await prisma.user.update({
      data: { enabled: true, role: "user" },
      where: { id: owner.id },
    });
    const claimCallback = await completeLegacySsoCallback(handle);
    expect(claimCallback.headers.get("location")).toBe(
      new URL("/dashboard", legacySsoWebOrigin).href
    );
    const switchFlow = await completeAuthenticatedLegacySsoCallback(
      handle,
      initiatingBearer,
      "/dashboard"
    );
    expect(switchFlow.callback.headers.get("location")).toBe(
      new URL("/legacy-sso/confirm", legacySsoWebOrigin).href
    );
    const switchCookie = cookiePairFrom(
      switchFlow.callback,
      "__Host-folio-sso-switch"
    );
    const details = await handle(
      new Request("https://folio.example.test/api/legacy-sso/switch", {
        headers: {
          Authorization: `Bearer ${initiatingBearer}`,
          Cookie: switchCookie,
        },
      })
    );
    expect(details.status).toBe(200);
    const { confirmationFingerprint } = (await details.json()) as {
      confirmationFingerprint: string;
    };
    const ownerSessionCount = await prisma.session.count({
      where: { userId: owner.id },
    });
    await prisma.user.update({ data: ineligible, where: { id: owner.id } });
    const deniedCallback = await completeLegacySsoCallback(handle);
    expect(deniedCallback.headers.get("location")).toBe(
      legacySsoFailedLocation
    );
    const deniedClaim = await handle(
      new Request("https://folio.example.test/api/legacy-sso/session", {
        headers: {
          Cookie: cookiePairFrom(claimCallback, "__Host-folio-sso-session"),
          Origin: legacySsoWebOrigin,
        },
        method: "POST",
      })
    );
    expect(deniedClaim.status).toBe(401);
    const deniedSwitch = await handle(
      new Request("https://folio.example.test/api/legacy-sso/switch/confirm", {
        body: JSON.stringify({ confirmationFingerprint }),
        headers: {
          ...jsonHeaders,
          Authorization: `Bearer ${initiatingBearer}`,
          Cookie: switchCookie,
          Origin: legacySsoWebOrigin,
        },
        method: "POST",
      })
    );
    expect(deniedSwitch.status).toBe(401);
    expect(
      await prisma.session.count({ where: { userId: owner.id } })
    ).toBeLessThanOrEqual(ownerSessionCount);
    const retainedSession = await handle(
      new Request("https://folio.example.test/api/session", {
        headers: { Authorization: `Bearer ${initiatingBearer}` },
      })
    );
    expect(retainedSession.status).toBe(200);
    expect(await retainedSession.json()).toMatchObject({
      user: { id: initiator.id },
    });
  }
  await prisma.user.update({
    data: { enabled: true, role: "user" },
    where: { id: owner.id },
  });
  for (const sessionChange of ["expired", "revoked"]) {
    const callback = await completeLegacySsoCallback(handle);
    expect(callback.headers.get("location")).toBe(
      new URL("/dashboard", legacySsoWebOrigin).href
    );
    const transferCookie = cookiePairFrom(callback, "__Host-folio-sso-session");
    const [, transferId] = transferCookie.split("=", 2);
    if (!transferId) {
      throw new Error("SSO Session transfer is missing");
    }
    const transfer = await prisma.verification.findUniqueOrThrow({
      where: { id: transferId },
    });
    await (sessionChange === "expired"
      ? prisma.session.update({
          data: { expiresAt: new Date(0) },
          where: { token: transfer.value },
        })
      : prisma.session.delete({ where: { token: transfer.value } }));
    const deniedClaim = await handle(
      new Request("https://folio.example.test/api/legacy-sso/session", {
        headers: { Cookie: transferCookie, Origin: legacySsoWebOrigin },
        method: "POST",
      })
    );
    expect(deniedClaim.status).toBe(401);
  }
});

test("expired SSO re-entry restores the owned draft", async () => {
  const email = `Ticket-13-${crypto.randomUUID()}@EXAMPLE.COM`;
  const normalizedEmail = email.toLowerCase();
  const identity = {
    email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-13-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
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
  const userCountBefore = await prisma.user.count();
  const accountCountBefore = await prisma.account.count();
  const firstCallback = await completeLegacySsoCallback(handle);
  expect(firstCallback.status).toBe(303);
  expect(firstCallback.headers.get("location")).toBe(
    new URL("/dashboard", legacySsoWebOrigin).href
  );
  const firstBearer = await claimLegacySsoBearer(handle, firstCallback);
  const user = await prisma.user.findUnique({
    where: { email: normalizedEmail },
  });
  if (!user) {
    throw new Error("First Legacy SSO login did not create a User");
  }
  expect(user).toMatchObject({
    email: normalizedEmail,
    emailVerified: false,
    enabled: true,
    mustChangePassword: false,
    name: normalizedEmail,
    role: "user",
  });
  expect(await prisma.user.count()).toBe(userCountBefore + 1);
  expect(await prisma.account.count()).toBe(accountCountBefore + 1);
  const account = await prisma.account.findUnique({
    where: {
      providerId_accountId: {
        accountId: identity.sub,
        providerId,
      },
    },
  });
  expect(account).toMatchObject({
    issuer: providerId,
    password: null,
    userId: user.id,
  });
  const firstSessionResponse = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${firstBearer}` },
    })
  );
  expect(firstSessionResponse.status).toBe(200);
  expect(await firstSessionResponse.json()).toMatchObject({
    user: {
      email: normalizedEmail,
      id: user.id,
      mustChangePassword: false,
      name: normalizedEmail,
      role: "user",
    },
  });

  const formId = crypto.randomUUID();
  const publicId = crypto.randomUUID().replaceAll("-", "");
  const templateBytes = docxXmlFixture({
    document: contentControlDocument(
      contentControl({
        alias: "Full name",
        placeholderText: "Enter full name",
        tag: "full_name",
        type: "<w:text/>",
      })
    ),
  });
  const templateObjectKey = objectKey(
    "forms",
    formId,
    "published",
    crypto.randomUUID(),
    "ticket-13.docx"
  );
  await putObject(templateObjectKey, templateBytes, DOCX_CONTENT_TYPE);
  await prisma.form.create({
    data: {
      createdBy: user.id,
      fillMethod: "native",
      id: formId,
      publicId,
      publishedTemplate: {
        create: {
          contentHash: createHash("sha256").update(templateBytes).digest("hex"),
          documentKey: `ticket-13-template-${crypto.randomUUID()}`,
          id: crypto.randomUUID(),
          manifest: {
            create: {
              configurationHash: createHash("sha256")
                .update(templateBytes)
                .digest("hex"),
              fields: {
                create: {
                  label: "Full name",
                  placeholder: "Enter full name",
                  position: 0,
                  tag: "full_name",
                  type: "text",
                },
              },
            },
          },
          objectKey: templateObjectKey,
          version: 1,
        },
      },
      status: "published",
      title: "Ticket 13 Response owner continuity",
      version: 1,
    },
  });
  const firstStartResponse = await handle(
    new Request(`https://folio.example.test/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${firstBearer}` },
      method: "POST",
    })
  );
  expect(firstStartResponse.status).toBe(200);
  const firstStartBody = (await firstStartResponse.json()) as {
    response?: { id?: string };
  };
  const responseId = firstStartBody.response?.id;
  if (!responseId) {
    throw new Error("First SSO User could not start a Response");
  }
  expect(
    await prisma.response.findUnique({ where: { id: responseId } })
  ).toMatchObject({ formId, userId: user.id });
  const editorResponse = await handle(
    new Request(
      `https://folio.example.test/api/forms/${publicId}/editor-config?responseId=${responseId}&action=fill`,
      { headers: { Authorization: `Bearer ${firstBearer}` } }
    )
  );
  expect(editorResponse.status).toBe(200);
  const editor = (await editorResponse.json()) as {
    capabilities: Record<"save-draft", string>;
    documentKey: string;
  };
  const savedData = { full_name: "PDMS saved draft" };
  const saveResponse = await handle(
    new Request(`https://folio.example.test/api/forms/${publicId}/draft`, {
      body: JSON.stringify({
        data: savedData,
        documentKey: editor.documentKey,
        fillMethod: "native",
        responseId,
      }),
      headers: {
        ...jsonHeaders,
        "X-Editor-Capability": editor.capabilities["save-draft"],
      },
      method: "POST",
    })
  );
  expect(saveResponse.status).toBe(202);
  const save = (await saveResponse.json()) as {
    operationId?: string;
    operationCapability?: string;
  };
  if (!save.operationId || !save.operationCapability) {
    throw new Error("SSO Draft save did not create an Operation");
  }
  expect(
    await waitForOperation(app, save.operationId, {
      "X-Editor-Capability": save.operationCapability,
    })
  ).toMatchObject({ status: "completed" });
  const savedResponse = await prisma.response.findUniqueOrThrow({
    where: { id: responseId },
  });
  expect(savedResponse.draftData).toEqual(savedData);
  if (!savedResponse.draftObjectKey) {
    throw new Error("SSO Draft save did not persist its document");
  }
  const savedDocument = await readObject(savedResponse.draftObjectKey);
  expect(
    new TextDecoder().decode(unzipSync(savedDocument)["word/document.xml"])
  ).toContain("PDMS saved draft");
  const saveClaims = verifyEditorCapability(editor.capabilities["save-draft"]);
  if (!saveClaims?.leaseId) {
    throw new Error("SSO Draft editor lease is missing");
  }
  // Close the editor after its completed save before leaving for SSO.
  const editorExit = await handle(
    new Request(
      `https://folio.example.test/api/editor-leases/${saveClaims.leaseId}`,
      { headers: { Authorization: `Bearer ${firstBearer}` }, method: "DELETE" }
    )
  );
  expect(editorExit.status).toBe(200);

  const retainedEmail = `ticket-13-profile-${crypto.randomUUID()}@example.com`;
  await prisma.user.update({
    data: { email: retainedEmail, name: "Retained local profile" },
    where: { id: user.id },
  });
  identity.email = `changed-${crypto.randomUUID()}@example.com`;
  const returnTo = `/forms/${publicId}/fill?responseId=${responseId}`;
  const repeatedStart = await startLegacySsoBrowser(
    handle,
    returnTo,
    firstBearer
  );
  expect(repeatedStart.startResponse.status).toBe(200);
  const repeatedAuthorization = await fetch(repeatedStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(repeatedAuthorization.status).toBe(303);
  const repeatedCallbackUrl = repeatedAuthorization.headers.get("location");
  if (!repeatedCallbackUrl) {
    throw new Error("The repeated SSO login did not return a callback");
  }
  const initiatingSession = await prisma.session.findFirst({
    orderBy: { createdAt: "desc" },
    where: { isSso: true, userId: user.id },
  });
  if (!initiatingSession) {
    throw new Error("The initiating Folio Session was unavailable");
  }
  await prisma.session.update({
    data: { expiresAt: new Date(0) },
    where: { id: initiatingSession.id },
  });
  const repeatedCallback = await handle(
    new Request(repeatedCallbackUrl, {
      headers: { Cookie: repeatedStart.preLoginCookie },
    })
  );
  expect(repeatedCallback.status).toBe(303);
  expect(repeatedCallback.headers.get("location")).toBe(
    new URL(returnTo, legacySsoWebOrigin).href
  );
  const repeatedBearer = await claimLegacySsoBearer(handle, repeatedCallback);
  const repeatedSessionResponse = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${repeatedBearer}` },
    })
  );
  expect(repeatedSessionResponse.status).toBe(200);
  expect(await repeatedSessionResponse.json()).toMatchObject({
    user: {
      email: retainedEmail,
      id: user.id,
      name: "Retained local profile",
      role: "user",
    },
  });
  const repeatedStartResponse = await handle(
    new Request(`https://folio.example.test/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${repeatedBearer}` },
      method: "POST",
    })
  );
  expect(repeatedStartResponse.status).toBe(200);
  const repeatedStartBody = (await repeatedStartResponse.json()) as {
    response?: { id?: string };
  };
  expect(repeatedStartBody.response?.id).toBe(responseId);
  expect(
    await prisma.response.findUnique({ where: { id: responseId } })
  ).toMatchObject({ formId, userId: user.id });
  const resumedEditor = await handle(
    new Request(
      `https://folio.example.test/api/forms/${publicId}/editor-config?responseId=${responseId}&action=fill`,
      { headers: { Authorization: `Bearer ${repeatedBearer}` } }
    )
  );
  expect(resumedEditor.status).toBe(200);
  expect(await resumedEditor.json()).toMatchObject({
    data: savedData,
    fillMethod: "native",
    responseId,
  });
  const resumedDocument = await handle(
    new Request(
      `https://folio.example.test/api/responses/${responseId}/draft/docx`,
      {
        headers: { Authorization: `Bearer ${repeatedBearer}` },
      }
    )
  );
  expect(resumedDocument.status).toBe(200);
  expect([...new Uint8Array(await resumedDocument.arrayBuffer())]).toEqual([
    ...savedDocument,
  ]);
  const target = await createCredentialFixture({
    email: `ticket-16-target-${crypto.randomUUID()}@example.com`,
    name: "Ticket 16 different User",
    password: "Ticket16-different-user-password",
  });
  const targetBearer = await bearerFor(
    app,
    target.email,
    "Ticket16-different-user-password"
  );
  for (const format of ["json", "docx"]) {
    const foreignExport = await handle(
      new Request(
        `https://folio.example.test/api/responses/${responseId}/draft/${format}`,
        { headers: { Authorization: `Bearer ${targetBearer}` } }
      )
    );
    expect(foreignExport.status).toBe(403);
  }
  const targetSubject = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: targetSubject,
      id: crypto.randomUUID(),
      providerId,
      userId: target.id,
    },
  });
  const crossAccountStart = await startLegacySsoBrowser(
    handle,
    returnTo,
    repeatedBearer
  );
  identity.email = target.email;
  identity.sub = targetSubject;
  const crossAccountAuthorization = await fetch(
    crossAccountStart.authorizationUrl,
    { redirect: "manual" }
  );
  const crossAccountCallbackUrl =
    crossAccountAuthorization.headers.get("location");
  if (!crossAccountCallbackUrl) {
    throw new Error(
      "The different-account SSO login did not return a callback"
    );
  }
  const crossAccountSession = await auth.api.getSession({
    headers: new Headers({ Authorization: `Bearer ${repeatedBearer}` }),
  });
  if (!crossAccountSession?.session) {
    throw new Error("The re-entered Folio Session was unavailable");
  }
  await prisma.session.update({
    data: { expiresAt: new Date(0) },
    where: { id: crossAccountSession.session.id },
  });
  const targetSessionCount = await prisma.session.count({
    where: { userId: target.id },
  });
  const crossAccountCallback = await handle(
    new Request(crossAccountCallbackUrl, {
      headers: { Cookie: crossAccountStart.preLoginCookie },
    })
  );
  const crossAccountLocation = new URL(
    crossAccountCallback.headers.get("location") ?? ""
  );
  expect(crossAccountLocation.pathname).toBe("/login");
  expect(crossAccountLocation.searchParams.get("legacySso")).toBe("failed");
  expect(crossAccountLocation.searchParams.get("returnTo")).toBe(returnTo);
  expect(await prisma.session.count({ where: { userId: target.id } })).toBe(
    targetSessionCount
  );
});

test("first Legacy SSO rejects invalid claims and queues email collisions for review", async () => {
  const email = `ticket-13-rejected-${crypto.randomUUID()}@example.com`;
  const subject = `legacy-${crypto.randomUUID()}`;
  const identity = { email, email_verified: true, sub: subject };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-13-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
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
  const expectRejectedLogin = async (): Promise<void> => {
    const callback = await completeLegacySsoCallback(handle);
    expect(callback.status).toBe(303);
    expect(callback.headers.get("location")).toBe(legacySsoFailedLocation);
  };
  const initialCounts = {
    accounts: await prisma.account.count(),
    sessions: await prisma.session.count(),
    users: await prisma.user.count(),
  };
  Reflect.deleteProperty(identity, "sub");
  await expectRejectedLogin();
  identity.sub = subject;
  Reflect.deleteProperty(identity, "email");
  await expectRejectedLogin();
  identity.email = email;
  Reflect.deleteProperty(identity, "email_verified");
  await expectRejectedLogin();
  expect(await prisma.account.count()).toBe(initialCounts.accounts);
  expect(await prisma.session.count()).toBe(initialCounts.sessions);
  expect(await prisma.user.count()).toBe(initialCounts.users);

  const duplicateEmail = `Ticket-13-DUPLICATE-${crypto.randomUUID()}@EXAMPLE.COM`;
  const normalizedDuplicateEmail = duplicateEmail.toLowerCase();
  const localUser = await createCredentialFixture({
    email: normalizedDuplicateEmail,
    name: "Existing local User",
    password: "Ticket13-existing-local-password",
  });
  const countsBeforeCollision = {
    accounts: await prisma.account.count(),
    sessions: await prisma.session.count(),
    users: await prisma.user.count(),
  };
  identity.email = duplicateEmail;
  identity.email_verified = false;
  const collision = await completeLegacySsoCallback(handle);
  expect(collision.status).toBe(303);
  expect(collision.headers.get("location")).toBe(legacySsoPendingLocation);
  expect(await prisma.account.count()).toBe(countsBeforeCollision.accounts);
  expect(await prisma.session.count()).toBe(countsBeforeCollision.sessions);
  expect(await prisma.user.count()).toBe(countsBeforeCollision.users);
  expect(
    await prisma.account.findUnique({
      where: {
        providerId_accountId: {
          accountId: subject,
          providerId,
        },
      },
    })
  ).toBeNull();
  expect(
    await prisma.legacyAccountLinkRequest.findUnique({
      where: { providerId_subject: { providerId, subject } },
    })
  ).toMatchObject({
    email: normalizedDuplicateEmail,
    status: "pending",
    userId: localUser.id,
  });
  expect(await prisma.session.count({ where: { userId: localUser.id } })).toBe(
    0
  );
});

test("racing first Legacy SSO logins create one User and identity binding", async () => {
  const identity = {
    email: `ticket-13-race-${crypto.randomUUID()}@example.com`,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-13-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
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
  const userCountBefore = await prisma.user.count();
  const accountCountBefore = await prisma.account.count();
  const [firstCallback, secondCallback] = await Promise.all([
    completeLegacySsoCallback(handle),
    completeLegacySsoCallback(handle),
  ]);
  expect(firstCallback.headers.get("location")).toBe(
    new URL("/dashboard", legacySsoWebOrigin).href
  );
  expect(secondCallback.headers.get("location")).toBe(
    new URL("/dashboard", legacySsoWebOrigin).href
  );
  const [firstBearer, secondBearer] = await Promise.all([
    claimLegacySsoBearer(handle, firstCallback),
    claimLegacySsoBearer(handle, secondCallback),
  ]);
  const user = await prisma.user.findUnique({
    where: { email: identity.email },
  });
  if (!user) {
    throw new Error("Racing Legacy SSO logins did not create a User");
  }
  expect(await prisma.user.count()).toBe(userCountBefore + 1);
  expect(await prisma.account.count()).toBe(accountCountBefore + 1);
  for (const bearer of [firstBearer, secondBearer]) {
    const sessionResponse = await handle(
      new Request("https://folio.example.test/api/session", {
        headers: { Authorization: `Bearer ${bearer}` },
      })
    );
    expect(sessionResponse.status).toBe(200);
    expect(await sessionResponse.json()).toMatchObject({
      user: { id: user.id, role: "user" },
    });
  }
  expect(
    await prisma.session.count({ where: { isSso: true, userId: user.id } })
  ).toBe(2);
});

test("failed first Legacy SSO persistence leaves no partial User or identity binding", async () => {
  const identity = {
    email: `ticket-13-persistence-${crypto.randomUUID()}@example.com`,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-13-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  const constraintName = `ticket13_reject_${crypto.randomUUID().replaceAll("-", "")}`;
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
  const sessionCountBefore = await prisma.session.count();
  await prisma.$executeRawUnsafe(
    `ALTER TABLE "account" ADD CONSTRAINT "${constraintName}" CHECK ("provider_id" <> '${providerId}') NOT VALID`
  );
  try {
    await expect(completeLegacySsoCallback(handle)).rejects.toBeInstanceOf(
      Error
    );
    expect(
      await prisma.user.findUnique({ where: { email: identity.email } })
    ).toBeNull();
    expect(
      await prisma.account.findUnique({
        where: {
          providerId_accountId: {
            accountId: identity.sub,
            providerId,
          },
        },
      })
    ).toBeNull();
    expect(
      await prisma.legacyAccountLinkRequest.findUnique({
        where: {
          providerId_subject: {
            providerId,
            subject: identity.sub,
          },
        },
      })
    ).toBeNull();
    expect(await prisma.session.count()).toBe(sessionCountBefore);
  } finally {
    await prisma.$executeRawUnsafe(
      `ALTER TABLE "account" DROP CONSTRAINT IF EXISTS "${constraintName}"`
    );
  }
});

for (const initialStatus of ["pending", "approved"] as const) {
  test(`PDMS ${initialStatus} review refresh retargets eligible email and fences displayed decisions and delayed callbacks`, async () => {
    const fixture = await createPdmsReviewFixture(app);
    retainLegacySsoMock(fixture.mock);
    const { candidate, handle, identity, request, reviewerBearer } = fixture;
    const target = await createCredentialFixture({
      email: `pdms-retarget-${crypto.randomUUID()}@example.com`,
      name: "Current PDMS candidate",
      password: fixture.password,
    });
    const displayed = await legacyAccountLinkSnapshot(
      handle,
      reviewerBearer,
      request.id
    );
    if (initialStatus === "approved") {
      const initialApproval = await submitLegacyAccountLinkReview(
        handle,
        reviewerBearer,
        displayed,
        "approve"
      );
      expect(initialApproval.status).toBe(200);
    }
    const before = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { id: request.id },
    });
    const gate = fixture.delayNextExchange();
    const delayedCallback = completeLegacySsoCallback(handle);
    try {
      await gate.entered;
      identity.email = target.email.toUpperCase();
      const refreshedCallback = await completeLegacySsoCallback(handle);
      expect(refreshedCallback.headers.get("location")).toBe(
        legacySsoPendingLocation
      );
      expect(
        cookiePairFrom(refreshedCallback, "__Host-folio-sso-session")
      ).toBe("__Host-folio-sso-session=");
      const refreshed = await prisma.legacyAccountLinkRequest.findUniqueOrThrow(
        { where: { id: request.id } }
      );
      expect(refreshed).toMatchObject({
        email: target.email,
        reviewedAt: null,
        reviewedById: null,
        status: "pending",
        userId: target.id,
      });
      expect(refreshed.reviewedGeneration).not.toBeNull();
      expect(refreshed.reviewedGeneration).toBeGreaterThan(
        before.reviewedGeneration ?? 0n
      );
      const current = await legacyAccountLinkSnapshot(
        handle,
        reviewerBearer,
        request.id
      );
      expect(current).toMatchObject({
        email: target.email,
        reviewedGeneration: refreshed.reviewedGeneration?.toString(),
        user: { email: target.email, id: target.id },
      });
      for (const decision of ["approve", "reject"] as const) {
        const staleDecision = await submitLegacyAccountLinkReview(
          handle,
          reviewerBearer,
          displayed,
          decision
        );
        expect(staleDecision.status).toBe(409);
        expect(await staleDecision.json()).toMatchObject({
          error: "account_link_changed",
          message: "Account link request changed; reload and review it again",
        });
        expect(
          await prisma.legacyAccountLinkRequest.findUnique({
            where: { id: request.id },
          })
        ).toEqual(refreshed);
      }
      const olderBrowser = await startLegacySsoBrowser(handle);
      const equalBrowser = await startLegacySsoBrowser(handle);
      const approval = await submitLegacyAccountLinkReview(
        handle,
        reviewerBearer,
        current,
        "approve"
      );
      expect(approval.status).toBe(200);
      const approved = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
        where: { id: request.id },
      });
      expect(approved).toMatchObject({
        email: target.email,
        status: "approved",
        userId: target.id,
      });
      expect(approved.reviewedGeneration).toBeGreaterThan(
        refreshed.reviewedGeneration ?? 0n
      );
      gate.release();
      const delayedResponse = await delayedCallback;
      expect(delayedResponse.headers.get("location")).toBe(
        legacySsoFailedLocation
      );
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: request.id },
        })
      ).toEqual(approved);
      const [, equalTransactionId] = equalBrowser.preLoginCookie.split("=", 2);
      if (!equalTransactionId || approved.reviewedGeneration === null) {
        throw new Error(
          "The approved generation or browser transaction is missing"
        );
      }
      const equalTransaction = await prisma.verification.findUniqueOrThrow({
        where: { id: equalTransactionId },
      });
      await prisma.verification.update({
        data: {
          value: JSON.stringify({
            ...(JSON.parse(equalTransaction.value) as Record<string, unknown>),
            generation: approved.reviewedGeneration.toString(),
          }),
        },
        where: { id: equalTransactionId },
      });
      for (const browser of [olderBrowser, equalBrowser]) {
        const authorization = await fetch(browser.authorizationUrl, {
          redirect: "manual",
        });
        const location = authorization.headers.get("location");
        if (!location) {
          throw new Error("The stale authorization did not return a callback");
        }
        const callback = await handle(
          new Request(location, { headers: { Cookie: browser.preLoginCookie } })
        );
        expect(callback.headers.get("location")).toBe(legacySsoFailedLocation);
        expect(cookiePairFrom(callback, "__Host-folio-sso-session")).toBe(
          "__Host-folio-sso-session="
        );
      }
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: request.id },
        })
      ).toEqual(approved);
      expect(
        await prisma.account.count({
          where: {
            accountId: identity.sub,
            providerId: fixture.config.providerId,
          },
        })
      ).toBe(0);
      expect(
        await prisma.session.count({
          where: { isSso: true, userId: { in: [candidate.id, target.id] } },
        })
      ).toBe(0);
      const freshCallback = await completeLegacySsoCallback(handle);
      expect(freshCallback.headers.get("location")).toBe(
        new URL("/dashboard", legacySsoWebOrigin).href
      );
      const bearer = await claimLegacySsoBearer(handle, freshCallback);
      const session = await handle(
        new Request(new URL("/api/session", legacySsoCallbackUrl).href, {
          headers: { Authorization: `Bearer ${bearer}` },
        })
      );
      expect(session.status).toBe(200);
      expect(await session.json()).toMatchObject({
        user: { email: target.email, id: target.id },
      });
      expect(
        await prisma.account.findUnique({
          where: {
            providerId_accountId: {
              accountId: identity.sub,
              providerId: fixture.config.providerId,
            },
          },
        })
      ).toMatchObject({ userId: target.id });
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: request.id },
        })
      ).toMatchObject({
        email: target.email,
        reviewedGeneration: approved.reviewedGeneration,
        status: "linked",
        userId: target.id,
      });
      expect(
        await prisma.session.count({
          where: { isSso: true, userId: candidate.id },
        })
      ).toBe(0);
    } finally {
      gate.release();
      await delayedCallback.catch(() => {});
    }
  });

  for (const blockedKind of ["unused", "admin", "disabled"] as const) {
    test(`PDMS ${initialStatus} changed ${blockedKind} email stays visibly blocked without provisioning or fence churn`, async () => {
      const fixture = await createPdmsReviewFixture(app);
      retainLegacySsoMock(fixture.mock);
      const { candidate, handle, identity, request, reviewerBearer } = fixture;
      const displayed = await legacyAccountLinkSnapshot(
        handle,
        reviewerBearer,
        request.id
      );
      if (initialStatus === "approved") {
        const initialApproval = await submitLegacyAccountLinkReview(
          handle,
          reviewerBearer,
          displayed,
          "approve"
        );
        expect(initialApproval.status).toBe(200);
      }
      const blockedEmail = `pdms-blocked-${crypto.randomUUID()}@example.com`;
      let blockedUser: { email: string; id: string } | undefined;
      if (blockedKind !== "unused") {
        blockedUser = await createCredentialFixture({
          email: blockedEmail,
          name: "Initially ineligible candidate",
          password: fixture.password,
          role: blockedKind === "admin" ? "admin" : "user",
        });
        if (blockedKind === "disabled") {
          await prisma.user.update({
            data: { enabled: false },
            where: { id: blockedUser.id },
          });
        }
      }
      identity.email = blockedEmail;
      const userCountBefore = await prisma.user.count({
        where: { email: blockedEmail },
      });
      const callback = await completeLegacySsoCallback(handle);
      expect(callback.headers.get("location")).toBe(legacySsoPendingLocation);
      expect(cookiePairFrom(callback, "__Host-folio-sso-session")).toBe(
        "__Host-folio-sso-session="
      );
      const blocked = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
        where: { id: request.id },
      });
      expect(blocked).toMatchObject({
        email: blockedEmail,
        reviewedAt: null,
        reviewedById: null,
        status: "pending",
        userId: candidate.id,
      });
      expect(blocked.reviewedGeneration).not.toBeNull();
      const blockedSnapshot = await legacyAccountLinkSnapshot(
        handle,
        reviewerBearer,
        request.id
      );
      expect(blockedSnapshot).toMatchObject({
        email: blockedEmail,
        reviewedGeneration: blocked.reviewedGeneration?.toString(),
        user: { email: candidate.email, id: candidate.id },
      });
      const approval = await submitLegacyAccountLinkReview(
        handle,
        reviewerBearer,
        blockedSnapshot,
        "approve"
      );
      expect(approval.status).toBe(409);
      expect(await approval.json()).toMatchObject({
        error: "account_link_not_eligible",
      });
      const blockedCallback = await completeLegacySsoCallback(handle);
      expect(blockedCallback.headers.get("location")).toBe(
        legacySsoPendingLocation
      );
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: request.id },
        })
      ).toEqual(blocked);
      expect(await prisma.user.count({ where: { email: blockedEmail } })).toBe(
        userCountBefore
      );
      expect(
        await prisma.account.count({
          where: { providerId: fixture.config.providerId },
        })
      ).toBe(0);
      expect(
        await prisma.session.count({
          where: {
            isSso: true,
            userId: {
              in: [candidate.id, ...(blockedUser ? [blockedUser.id] : [])],
            },
          },
        })
      ).toBe(0);
      if (blockedUser) {
        const eligibility = await handle(
          new Request(
            new URL(`/api/admin/users/${blockedUser.id}`, legacySsoCallbackUrl)
              .href,
            {
              body: JSON.stringify(
                blockedKind === "admin" ? { role: "user" } : { enabled: true }
              ),
              headers: {
                ...jsonHeaders,
                Authorization: `Bearer ${reviewerBearer}`,
              },
              method: "PATCH",
            }
          )
        );
        expect(eligibility.status).toBe(200);
      } else {
        blockedUser = await createCredentialFixture({
          email: blockedEmail,
          name: "Newly available candidate",
          password: fixture.password,
        });
      }
      const reboundCallback = await completeLegacySsoCallback(handle);
      expect(reboundCallback.headers.get("location")).toBe(
        legacySsoPendingLocation
      );
      expect(cookiePairFrom(reboundCallback, "__Host-folio-sso-session")).toBe(
        "__Host-folio-sso-session="
      );
      const rebound = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
        where: { id: request.id },
      });
      expect(rebound).toMatchObject({
        email: blockedEmail,
        reviewedAt: null,
        reviewedById: null,
        status: "pending",
        userId: blockedUser.id,
      });
      expect(rebound.reviewedGeneration).toBeGreaterThan(
        blocked.reviewedGeneration ?? 0n
      );
      for (const decision of ["approve", "reject"] as const) {
        const staleDecision = await submitLegacyAccountLinkReview(
          handle,
          reviewerBearer,
          blockedSnapshot,
          decision
        );
        expect(staleDecision.status).toBe(409);
      }
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: request.id },
        })
      ).toEqual(rebound);
      expect(
        await prisma.account.count({
          where: { providerId: fixture.config.providerId },
        })
      ).toBe(0);
      expect(
        await prisma.session.count({
          where: { isSso: true, userId: blockedUser.id },
        })
      ).toBe(0);
      const current = await legacyAccountLinkSnapshot(
        handle,
        reviewerBearer,
        request.id
      );
      const currentApproval = await submitLegacyAccountLinkReview(
        handle,
        reviewerBearer,
        current,
        "approve"
      );
      expect(currentApproval.status).toBe(200);
      const fresh = await completeLegacySsoCallback(handle);
      const bearer = await claimLegacySsoBearer(handle, fresh);
      const currentSession = await handle(
        new Request(new URL("/api/session", legacySsoCallbackUrl).href, {
          headers: { Authorization: `Bearer ${bearer}` },
        })
      );
      expect(await currentSession.json()).toMatchObject({
        user: { id: blockedUser.id },
      });
    });
  }
}

for (const loss of ["disabled", "admin"] as const) {
  test(`PDMS approved candidate becoming ${loss} invalidates unchanged identity evidence once`, async () => {
    const fixture = await createPdmsReviewFixture(app);
    retainLegacySsoMock(fixture.mock);
    const { candidate, handle, request, reviewerBearer } = fixture;
    const displayed = await legacyAccountLinkSnapshot(
      handle,
      reviewerBearer,
      request.id
    );
    const approval = await submitLegacyAccountLinkReview(
      handle,
      reviewerBearer,
      displayed,
      "approve"
    );
    expect(approval.status).toBe(200);
    const approved = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { id: request.id },
    });
    const patch = await handle(
      new Request(
        new URL(`/api/admin/users/${candidate.id}`, legacySsoCallbackUrl).href,
        {
          body: JSON.stringify(
            loss === "disabled" ? { enabled: false } : { role: "admin" }
          ),
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${reviewerBearer}`,
          },
          method: "PATCH",
        }
      )
    );
    expect(patch.status).toBe(200);
    const callback = await completeLegacySsoCallback(handle);
    expect(callback.headers.get("location")).toBe(legacySsoPendingLocation);
    expect(cookiePairFrom(callback, "__Host-folio-sso-session")).toBe(
      "__Host-folio-sso-session="
    );
    const invalidated = await prisma.legacyAccountLinkRequest.findUniqueOrThrow(
      { where: { id: request.id } }
    );
    expect(invalidated).toMatchObject({
      email: candidate.email,
      reviewedAt: null,
      reviewedById: null,
      status: "pending",
      userId: candidate.id,
    });
    expect(invalidated.reviewedGeneration).toBeGreaterThan(
      approved.reviewedGeneration ?? 0n
    );
    const current = await legacyAccountLinkSnapshot(
      handle,
      reviewerBearer,
      request.id
    );
    expect(current.user).toMatchObject(
      loss === "disabled" ? { enabled: false } : { role: "admin" }
    );
    const blockedApproval = await submitLegacyAccountLinkReview(
      handle,
      reviewerBearer,
      current,
      "approve"
    );
    expect(blockedApproval.status).toBe(409);
    expect(await blockedApproval.json()).toMatchObject({
      error: "account_link_not_eligible",
    });
    const blockedCallback = await completeLegacySsoCallback(handle);
    expect(blockedCallback.headers.get("location")).toBe(
      legacySsoPendingLocation
    );
    expect(
      await prisma.legacyAccountLinkRequest.findUnique({
        where: { id: request.id },
      })
    ).toEqual(invalidated);
    expect(
      await prisma.account.count({
        where: { providerId: fixture.config.providerId },
      })
    ).toBe(0);
    expect(
      await prisma.session.count({
        where: { isSso: true, userId: candidate.id },
      })
    ).toBe(0);
  });
}
