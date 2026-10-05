import { test, expect, afterEach } from "bun:test";
import { createHash } from "node:crypto";

import { prisma, Prisma } from "@onlyoffice/db";
import { startLegacySsoMock } from "prefill-mock/mock";
import type { LegacySsoMockServer } from "prefill-mock/mock";

import { createApp } from "../../src/app";
import { objectKey, putObject, DOCX_CONTENT_TYPE } from "../../src/storage";
import { docxFixture } from "../fixtures/documents";
import {
  jsonHeaders,
  createCredentialFixture,
  bearerFor,
} from "../fixtures/http";
import {
  createPdmsReviewFixture,
  legacyAccountLinkSnapshot,
  submitLegacyAccountLinkReview,
  completeLegacySsoCallback,
  legacySsoCallbackUrl,
  legacySsoFailedLocation,
  claimLegacySsoBearer,
  legacySsoPendingLocation,
  createLegacySsoTestApp,
  cookiePairFrom,
  startLegacySsoBrowser,
  legacySsoWebOrigin,
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

test("PDMS candidate email PATCH and restore invalidate approval and delayed callback without touching terminal reviews", async () => {
  const fixture = await createPdmsReviewFixture(app);
  legacySsoMock = fixture.mock;
  const { candidate, handle, identity, request, reviewerBearer } = fixture;
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
  const terminal = await Promise.all(
    ["linked", "rejected"].map((status) =>
      prisma.legacyAccountLinkRequest.create({
        data: {
          email: candidate.email,
          providerId: fixture.config.providerId,
          reviewedAt: new Date(),
          reviewedById: fixture.reviewer.id,
          reviewedGeneration: approved.reviewedGeneration,
          status: status === "linked" ? "linked" : "rejected",
          subject: `pdms-terminal-${crypto.randomUUID()}`,
          userId: candidate.id,
        },
      })
    )
  );
  const pendingSibling = await prisma.legacyAccountLinkRequest.create({
    data: {
      email: `pdms-asserted-${crypto.randomUUID()}@example.com`,
      providerId: fixture.config.providerId,
      subject: `pdms-pending-${crypto.randomUUID()}`,
      userId: candidate.id,
    },
  });
  const gate = fixture.delayNextExchange();
  const delayedCallback = completeLegacySsoCallback(handle);
  try {
    await gate.entered;
    let lastGeneration = approved.reviewedGeneration;
    let siblingGeneration = pendingSibling.reviewedGeneration;
    for (const email of [
      `pdms-renamed-${crypto.randomUUID()}@example.com`,
      candidate.email,
    ]) {
      const patch = await handle(
        new Request(
          new URL(`/api/admin/users/${candidate.id}`, legacySsoCallbackUrl)
            .href,
          {
            body: JSON.stringify({ email }),
            headers: {
              ...jsonHeaders,
              Authorization: `Bearer ${reviewerBearer}`,
            },
            method: "PATCH",
          }
        )
      );
      expect(patch.status).toBe(200);
      const invalidated =
        await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
          where: { id: request.id },
        });
      expect(invalidated).toMatchObject({
        email: candidate.email,
        reviewedAt: null,
        reviewedById: null,
        status: "pending",
        userId: candidate.id,
      });
      expect(invalidated.reviewedGeneration).toBeGreaterThan(
        lastGeneration ?? 0n
      );
      lastGeneration = invalidated.reviewedGeneration;
      const invalidatedSibling =
        await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
          where: { id: pendingSibling.id },
        });
      expect(invalidatedSibling).toMatchObject({
        email: pendingSibling.email,
        providerId: pendingSibling.providerId,
        reviewedAt: null,
        reviewedById: null,
        status: "pending",
        subject: pendingSibling.subject,
        userId: candidate.id,
      });
      expect(invalidatedSibling.reviewedGeneration).toBeGreaterThan(
        siblingGeneration ?? 0n
      );
      siblingGeneration = invalidatedSibling.reviewedGeneration;
    }
    const restored = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { id: request.id },
    });
    for (const decision of ["approve", "reject"] as const) {
      const staleDecision = await submitLegacyAccountLinkReview(
        handle,
        reviewerBearer,
        displayed,
        decision
      );
      expect(staleDecision.status).toBe(409);
    }
    gate.release();
    const delayedResponse = await delayedCallback;
    expect(delayedResponse.headers.get("location")).toBe(
      legacySsoFailedLocation
    );
    expect(
      await prisma.legacyAccountLinkRequest.findUnique({
        where: { id: request.id },
      })
    ).toEqual(restored);
    for (const unchanged of terminal) {
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: unchanged.id },
        })
      ).toEqual(unchanged);
    }
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
        where: { isSso: true, userId: candidate.id },
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
      user: { id: candidate.id },
    });
  } finally {
    gate.release();
    await delayedCallback.catch(() => {});
  }
});

test("PDMS current rejection remains terminal against delayed callbacks and changed identity namespaces", async () => {
  const fixture = await createPdmsReviewFixture(app);
  legacySsoMock = fixture.mock;
  const { candidate, handle, identity, request, reviewerBearer } = fixture;
  const target = await createCredentialFixture({
    email: `pdms-rejected-target-${crypto.randomUUID()}@example.com`,
    name: "Other candidate",
    password: fixture.password,
  });
  const displayed = await legacyAccountLinkSnapshot(
    handle,
    reviewerBearer,
    request.id
  );
  const gate = fixture.delayNextExchange();
  const delayedCallback = completeLegacySsoCallback(handle);
  try {
    await gate.entered;
    identity.email = target.email;
    const changedCallback = await completeLegacySsoCallback(handle);
    expect(changedCallback.headers.get("location")).toBe(
      legacySsoPendingLocation
    );
    const staleRejection = await submitLegacyAccountLinkReview(
      handle,
      reviewerBearer,
      displayed,
      "reject"
    );
    expect(staleRejection.status).toBe(409);
    const current = await legacyAccountLinkSnapshot(
      handle,
      reviewerBearer,
      request.id
    );
    const currentRejection = await submitLegacyAccountLinkReview(
      handle,
      reviewerBearer,
      current,
      "reject"
    );
    expect(currentRejection.status).toBe(200);
    const rejected = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { id: request.id },
    });
    expect(rejected).toMatchObject({
      email: target.email,
      status: "rejected",
      userId: target.id,
    });
    const originalCandidate = await prisma.user.findUniqueOrThrow({
      where: { id: candidate.id },
    });
    gate.release();
    const delayedResponse = await delayedCallback;
    expect(delayedResponse.headers.get("location")).toBe(
      legacySsoFailedLocation
    );
    for (const email of [
      candidate.email,
      `pdms-rejected-unused-${crypto.randomUUID()}@example.com`,
    ]) {
      identity.email = email;
      const rejectedCallback = await completeLegacySsoCallback(handle);
      expect(rejectedCallback.headers.get("location")).toBe(
        legacySsoFailedLocation
      );
      expect(await prisma.user.findUnique({ where: { email } })).toEqual(
        email === candidate.email ? originalCandidate : null
      );
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: request.id },
        })
      ).toEqual(rejected);
    }
    identity.email = target.email;
    const subject = identity.sub;
    identity.sub = `pdms-other-${crypto.randomUUID()}`;
    const otherIdentityCallback = await completeLegacySsoCallback(handle);
    expect(otherIdentityCallback.headers.get("location")).toBe(
      legacySsoPendingLocation
    );
    const otherSubject =
      await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
        where: {
          providerId_subject: {
            providerId: fixture.config.providerId,
            subject: identity.sub,
          },
        },
      });
    expect(otherSubject).toMatchObject({
      email: target.email,
      status: "pending",
      userId: target.id,
    });
    expect(otherSubject.id).not.toBe(request.id);
    identity.sub = subject;
    const otherProviderId = `pdms-other-provider-${crypto.randomUUID()}`;
    const otherApp = createLegacySsoTestApp({
      ...fixture.config,
      providerId: otherProviderId,
    });
    const otherProviderCallback = await completeLegacySsoCallback((incoming) =>
      otherApp.handle(incoming)
    );
    expect(otherProviderCallback.headers.get("location")).toBe(
      legacySsoPendingLocation
    );
    const otherProvider =
      await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
        where: { providerId_subject: { providerId: otherProviderId, subject } },
      });
    expect(otherProvider).toMatchObject({
      email: target.email,
      status: "pending",
      userId: target.id,
    });
    expect(otherProvider.id).not.toBe(request.id);
    expect(
      await prisma.legacyAccountLinkRequest.findUnique({
        where: { id: request.id },
      })
    ).toEqual(rejected);
    expect(
      await prisma.account.count({
        where: {
          providerId: { in: [fixture.config.providerId, otherProviderId] },
        },
      })
    ).toBe(0);
    expect(
      await prisma.session.count({
        where: { isSso: true, userId: { in: [candidate.id, target.id] } },
      })
    ).toBe(0);
  } finally {
    gate.release();
    await delayedCallback.catch(() => {});
  }
});

test("PDMS Admin decisions require exact canonical generation bodies after authorization", async () => {
  const fixture = await createPdmsReviewFixture(app);
  legacySsoMock = fixture.mock;
  const { candidate, handle, request, reviewerBearer } = fixture;
  const localBearer = await bearerFor(app, candidate.email, fixture.password);
  const displayed = await legacyAccountLinkSnapshot(
    handle,
    reviewerBearer,
    request.id
  );
  expect(displayed.reviewedGeneration).toBeNull();
  const invalidBodies = [
    undefined,
    "",
    "{",
    "null",
    "[]",
    "true",
    '"1"',
    "{}",
    JSON.stringify({ reviewedGeneration: "1".repeat(1100) }),
    '{"reviewedGeneration":null,"extra":true}',
    ...[
      false,
      true,
      1,
      0,
      {},
      [],
      "",
      "0",
      "-1",
      "+1",
      "01",
      " 1",
      "1 ",
      "1.0",
      "1e3",
      "１２",
      "9223372036854775808",
      "10000000000000000000",
    ].map((reviewedGeneration) => JSON.stringify({ reviewedGeneration })),
  ];
  for (const decision of ["approve", "reject"] as const) {
    const url = new URL(
      `/api/admin/account-links/${request.id}/${decision}`,
      legacySsoCallbackUrl
    );
    for (const [bearer, status] of [
      [undefined, 401],
      [localBearer, 403],
    ] as const) {
      const unauthorized = await handle(
        new Request(url.href, {
          body: "{}",
          headers: bearer
            ? { ...jsonHeaders, Authorization: `Bearer ${bearer}` }
            : jsonHeaders,
          method: "POST",
        })
      );
      expect(unauthorized.status).toBe(status);
    }
    for (const body of invalidBodies) {
      const response = await handle(
        new Request(url.href, {
          body,
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${reviewerBearer}`,
          },
          method: "POST",
        })
      );
      expect(response.status).toBe(400);
      expect(await response.json()).toMatchObject({ error: "invalid_request" });
      expect(
        await prisma.legacyAccountLinkRequest.findUnique({
          where: { id: request.id },
        })
      ).toEqual(request);
    }
  }
  const maximum = await prisma.legacyAccountLinkRequest.create({
    data: {
      email: candidate.email,
      providerId: fixture.config.providerId,
      reviewedGeneration: 9_223_372_036_854_775_807n,
      subject: `pdms-maximum-${crypto.randomUUID()}`,
      userId: candidate.id,
    },
  });
  const maximumSnapshot = await legacyAccountLinkSnapshot(
    handle,
    reviewerBearer,
    maximum.id
  );
  expect(maximumSnapshot.reviewedGeneration).toBe("9223372036854775807");
  const rejection = await submitLegacyAccountLinkReview(
    handle,
    reviewerBearer,
    maximumSnapshot,
    "reject"
  );
  expect(rejection.status).toBe(200);
  expect(
    await prisma.legacyAccountLinkRequest.findUnique({
      where: { id: maximum.id },
    })
  ).toMatchObject({
    reviewedGeneration: 9_223_372_036_854_775_807n,
    status: "rejected",
  });
  const approval = await submitLegacyAccountLinkReview(
    handle,
    reviewerBearer,
    displayed,
    "approve"
  );
  expect(approval.status).toBe(200);
  expect(
    await prisma.legacyAccountLinkRequest.findUnique({
      where: { id: request.id },
    })
  ).toMatchObject({
    status: "approved",
    userId: candidate.id,
  });
});

test("Ticket 14 Legacy SSO collision requires fresh Admin-approved identity link", async () => {
  const email = `ticket-14-${crypto.randomUUID()}@example.com`;
  const password = "Ticket14-local-password";
  const candidate = await createCredentialFixture({
    email,
    name: "Ticket 14 Local User",
    password,
  });
  const localBearer = await bearerFor(app, email, password);
  const formId = crypto.randomUUID();
  const publicId = crypto.randomUUID().replaceAll("-", "");
  const templateBytes = docxFixture(`ticket-14-${crypto.randomUUID()}`);
  const templateObjectKey = objectKey(
    "forms",
    formId,
    "published",
    crypto.randomUUID(),
    "ticket-14.docx"
  );
  await putObject(templateObjectKey, templateBytes, DOCX_CONTENT_TYPE);
  await prisma.form.create({
    data: {
      createdBy: candidate.id,
      id: formId,
      publicId,
      publishedTemplate: {
        create: {
          contentHash: createHash("sha256").update(templateBytes).digest("hex"),
          documentKey: `ticket-14-template-${crypto.randomUUID()}`,
          id: crypto.randomUUID(),
          objectKey: templateObjectKey,
          version: 1,
        },
      },
      status: "published",
      title: "Ticket 14 Response ownership",
      version: 1,
    },
  });
  const startResponse = await app.handle(
    new Request(`http://test.local/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${localBearer}` },
      method: "POST",
    })
  );
  expect(startResponse.status).toBe(200);
  const startBody = (await startResponse.json()) as {
    response?: { id?: string };
  };
  const responseId = startBody.response?.id;
  if (!responseId) {
    throw new Error("Local User could not start a Response");
  }

  const identity = {
    email: email.toUpperCase(),
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const subject = identity.sub;
  const providerId = `legacy-${crypto.randomUUID()}`;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-14-secret-${crypto.randomUUID()}`;
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
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
    () => new Date("2026-01-01T00:00:00.000Z")
  );
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const ssoSessionsBefore = await prisma.session.count({
    where: { isSso: true, userId: candidate.id },
  });
  const [firstPending, racedPending] = await Promise.all([
    completeLegacySsoCallback(handle),
    completeLegacySsoCallback(handle),
  ]);
  for (const callback of [firstPending, racedPending]) {
    expect(callback.status).toBe(303);
    expect(callback.headers.get("location")).toBe(legacySsoPendingLocation);
    expect(cookiePairFrom(callback, "__Host-folio-sso-session")).toBe(
      "__Host-folio-sso-session="
    );
  }
  const linkRequest = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
    where: { providerId_subject: { providerId, subject } },
  });
  expect(linkRequest).toMatchObject({
    email,
    status: "pending",
    userId: candidate.id,
  });
  expect(
    await prisma.legacyAccountLinkRequest.count({
      where: { providerId, subject },
    })
  ).toBe(1);
  expect(
    await prisma.account.findUnique({
      where: { providerId_accountId: { accountId: subject, providerId } },
    })
  ).toBeNull();
  expect(
    await prisma.session.count({
      where: { isSso: true, userId: candidate.id },
    })
  ).toBe(ssoSessionsBefore);

  const adminEmail = `ticket-14-admin-${crypto.randomUUID()}@example.com`;
  const adminPassword = "Ticket14-admin-password";
  await createCredentialFixture({
    email: adminEmail,
    name: "Ticket 14 Admin",
    password: adminPassword,
    role: "admin",
  });
  const adminBearer = await bearerFor(app, adminEmail, adminPassword);
  const requestPath = `/api/admin/account-links/${linkRequest.id}/approve`;
  const displayedReview = await legacyAccountLinkSnapshot(
    handle,
    adminBearer,
    linkRequest.id
  );
  const forbiddenApproval = await handle(
    new Request(`https://folio.example.test${requestPath}`, {
      body: JSON.stringify({
        reviewedGeneration: displayedReview.reviewedGeneration,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${localBearer}` },
      method: "POST",
    })
  );
  expect(forbiddenApproval.status).toBe(403);
  expect(displayedReview).toMatchObject({
    email,
    id: linkRequest.id,
    providerId,
    reviewedGeneration: null,
    subject,
    user: {
      email,
      id: candidate.id,
      name: "Ticket 14 Local User",
      role: "user",
    },
  });

  const staleBrowser = await startLegacySsoBrowser(handle);
  const staleBackendResponse = await fetch(staleBrowser.authorizationUrl, {
    redirect: "manual",
  });
  const staleCallbackLocation = staleBackendResponse.headers.get("location");
  if (!staleCallbackLocation) {
    throw new Error("The stale SSO transaction did not return a callback");
  }
  const staleTransaction = await prisma.$transaction(async (tx) => {
    await tx.$queryRaw(
      Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${candidate.id} FOR UPDATE`
    );
    const approvalPromise = Promise.resolve(
      handle(
        new Request(`https://folio.example.test${requestPath}`, {
          body: JSON.stringify({
            reviewedGeneration: displayedReview.reviewedGeneration,
          }),
          headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
          method: "POST",
        })
      )
    );
    const browser = await startLegacySsoBrowser(handle);
    const backendResponse = await fetch(browser.authorizationUrl, {
      redirect: "manual",
    });
    const callbackLocation = backendResponse.headers.get("location");
    if (!callbackLocation) {
      throw new Error("The stale SSO transaction did not return a callback");
    }
    return { approvalPromise, browser, callbackLocation };
  });
  const approved = await staleTransaction.approvalPromise;
  expect(approved.status).toBe(200);
  expect(await approved.json()).toEqual({ ok: true });
  const approvedRequest =
    await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { id: linkRequest.id },
    });
  expect(approvedRequest.status).toBe("approved");
  expect(approvedRequest.reviewedAt).not.toBeNull();
  expect(
    await prisma.account.findUnique({
      where: { providerId_accountId: { accountId: subject, providerId } },
    })
  ).toBeNull();

  const staleCallback = await handle(
    new Request(staleCallbackLocation, {
      headers: { Cookie: staleBrowser.preLoginCookie },
    })
  );
  expect(staleCallback.headers.get("location")).toBe(legacySsoFailedLocation);
  expect(
    await prisma.session.count({
      where: { isSso: true, userId: candidate.id },
    })
  ).toBe(ssoSessionsBefore);

  identity.sub = `legacy-${crypto.randomUUID()}`;
  const wrongSubjectCallback = await completeLegacySsoCallback(handle);
  expect(wrongSubjectCallback.headers.get("location")).toBe(
    legacySsoPendingLocation
  );
  expect(
    await prisma.account.findUnique({
      where: { providerId_accountId: { accountId: subject, providerId } },
    })
  ).toBeNull();
  identity.sub = subject;
  const { reviewedAt } = approvedRequest;
  if (!reviewedAt) {
    throw new Error("Admin approval did not record review time");
  }
  const freshCallbacks = await Promise.all(
    Array.from({ length: 2 }, async () => {
      const browser = await startLegacySsoBrowser(handle);
      const oldBackendResponse = await fetch(browser.authorizationUrl, {
        redirect: "manual",
      });
      const callbackLocation = oldBackendResponse.headers.get("location");
      const [, transactionId] = browser.preLoginCookie.split("=", 2);
      if (!callbackLocation || !transactionId) {
        throw new Error("Fresh SSO transaction did not return a callback");
      }
      await prisma.verification.update({
        data: { createdAt: reviewedAt },
        where: { id: transactionId },
      });
      return handle(
        new Request(callbackLocation, {
          headers: { Cookie: browser.preLoginCookie },
        })
      );
    })
  );
  expect(freshCallbacks).toHaveLength(2);
  expect(
    freshCallbacks.map((response) => response.headers.get("location"))
  ).toEqual([
    new URL("/dashboard", legacySsoWebOrigin).href,
    new URL("/dashboard", legacySsoWebOrigin).href,
  ]);
  const freshBearers: string[] = [];
  for (const callback of freshCallbacks) {
    freshBearers.push(await claimLegacySsoBearer(handle, callback));
  }
  const [linkedBearer] = freshBearers;
  if (!linkedBearer) {
    throw new Error("Fresh SSO callbacks did not yield a bearer");
  }
  for (const bearer of freshBearers) {
    const session = await handle(
      new Request("https://folio.example.test/api/session", {
        headers: { Authorization: `Bearer ${bearer}` },
      })
    );
    expect(session.status).toBe(200);
    expect(await session.json()).toMatchObject({ user: { id: candidate.id } });
  }
  expect(
    await prisma.account.count({
      where: { accountId: subject, providerId },
    })
  ).toBe(1);
  expect(
    await prisma.session.count({
      where: { isSso: true, userId: candidate.id },
    })
  ).toBe(ssoSessionsBefore + 2);
  expect(
    await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { id: linkRequest.id },
    })
  ).toMatchObject({ status: "linked", userId: candidate.id });
  expect(
    await prisma.account.findUnique({
      where: { providerId_accountId: { accountId: subject, providerId } },
    })
  ).toMatchObject({ password: null, userId: candidate.id });
  const staleAfterLinkCallback = await handle(
    new Request(staleTransaction.callbackLocation, {
      headers: { Cookie: staleTransaction.browser.preLoginCookie },
    })
  );
  expect(staleAfterLinkCallback.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  expect(
    await prisma.user.findUniqueOrThrow({ where: { id: candidate.id } })
  ).toMatchObject({
    email,
    name: "Ticket 14 Local User",
    role: "user",
  });
  const linkedSessionResponse = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${linkedBearer}` },
    })
  );
  expect(linkedSessionResponse.status).toBe(200);
  const repeatedStart = await handle(
    new Request(`https://folio.example.test/api/forms/${publicId}/start`, {
      headers: { Authorization: `Bearer ${linkedBearer}` },
      method: "POST",
    })
  );
  expect(repeatedStart.status).toBe(200);
  const repeatedStartBody = (await repeatedStart.json()) as {
    response?: { id?: string };
  };
  expect(repeatedStartBody.response?.id).toBe(responseId);
  expect(
    await prisma.response.findUniqueOrThrow({ where: { id: responseId } })
  ).toMatchObject({ formId, userId: candidate.id });
  const localLoginAfterLink = await bearerFor(app, email, password);
  expect(
    await handle(
      new Request("https://folio.example.test/api/session", {
        headers: { Authorization: `Bearer ${localLoginAfterLink}` },
      })
    )
  ).toHaveProperty("status", 200);
});

test("Ticket 14 Legacy SSO approval rejects Admin and disabled candidates", async () => {
  const disabledEmail = `ticket-14-disabled-${crypto.randomUUID()}@example.com`;
  const disabledPassword = "Ticket14-disabled-password";
  const disabledUser = await createCredentialFixture({
    email: disabledEmail,
    name: "Ticket 14 Disabled User",
    password: disabledPassword,
  });
  await prisma.user.update({
    data: { enabled: false },
    where: { id: disabledUser.id },
  });
  const adminCandidateEmail = `ticket-14-admin-candidate-${crypto.randomUUID()}@example.com`;
  const adminCandidate = await createCredentialFixture({
    email: adminCandidateEmail,
    name: "Ticket 14 Admin Candidate",
    password: "Ticket14-admin-candidate-password",
    role: "admin",
  });
  const reviewerEmail = `ticket-14-reviewer-${crypto.randomUUID()}@example.com`;
  const reviewerPassword = "Ticket14-reviewer-password";
  await createCredentialFixture({
    email: reviewerEmail,
    name: "Ticket 14 Reviewer",
    password: reviewerPassword,
    role: "admin",
  });
  const reviewerBearer = await bearerFor(app, reviewerEmail, reviewerPassword);
  const identity = {
    email: disabledEmail,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const providerId = `legacy-${crypto.randomUUID()}`;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-14-secret-${crypto.randomUUID()}`;
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
  const disabledSubject = identity.sub;
  const disabledPending = await completeLegacySsoCallback(handle);
  expect(disabledPending.headers.get("location")).toBe(
    legacySsoPendingLocation
  );
  const disabledRequest =
    await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: {
        providerId_subject: { providerId, subject: disabledSubject },
      },
    });
  identity.email = adminCandidateEmail;
  identity.sub = `legacy-${crypto.randomUUID()}`;
  const adminSubject = identity.sub;
  const adminPending = await completeLegacySsoCallback(handle);
  expect(adminPending.headers.get("location")).toBe(legacySsoPendingLocation);
  const adminRequest = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
    where: { providerId_subject: { providerId, subject: adminSubject } },
  });
  for (const request of [disabledRequest, adminRequest]) {
    const snapshot = await legacyAccountLinkSnapshot(
      handle,
      reviewerBearer,
      request.id
    );
    const response = await submitLegacyAccountLinkReview(
      handle,
      reviewerBearer,
      snapshot,
      "approve"
    );
    expect(response.status).toBe(409);
  }
  expect(
    await prisma.legacyAccountLinkRequest.findMany({
      orderBy: { id: "asc" },
      select: { id: true, status: true },
      where: { id: { in: [disabledRequest.id, adminRequest.id] } },
    })
  ).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ id: disabledRequest.id, status: "pending" }),
      expect.objectContaining({ id: adminRequest.id, status: "pending" }),
    ])
  );
  const rejection = await submitLegacyAccountLinkReview(
    handle,
    reviewerBearer,
    await legacyAccountLinkSnapshot(handle, reviewerBearer, disabledRequest.id),
    "reject"
  );
  expect(rejection.status).toBe(200);
  expect(
    await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { id: disabledRequest.id },
    })
  ).toMatchObject({ status: "rejected", userId: disabledUser.id });
  identity.email = disabledEmail;
  identity.sub = disabledSubject;
  const rejectedRetry = await completeLegacySsoCallback(handle);
  expect(rejectedRetry.headers.get("location")).toBe(legacySsoFailedLocation);
  expect(
    await prisma.account.findUnique({
      where: {
        providerId_accountId: { accountId: disabledSubject, providerId },
      },
    })
  ).toBeNull();
  expect(
    await prisma.account.findUnique({
      where: {
        providerId_accountId: { accountId: adminSubject, providerId },
      },
    })
  ).toBeNull();
  expect(
    await prisma.session.count({
      where: {
        isSso: true,
        userId: { in: [disabledUser.id, adminCandidate.id] },
      },
    })
  ).toBe(0);
});

test("Ticket 14 stale SSO callback cannot bypass an account link created concurrently", async () => {
  const email = `ticket-14-race-${crypto.randomUUID()}@example.com`;
  const candidate = await createCredentialFixture({
    email,
    name: "Ticket 14 Race Candidate",
    password: "Ticket14-race-password",
  });
  const identity = {
    email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const subject = identity.sub;
  const providerId = `legacy-${crypto.randomUUID()}`;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-14-secret-${crypto.randomUUID()}`;
  const legacyMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    identity,
  });
  legacySsoMock = legacyMock;
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacyMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacyMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const staleBrowser = await startLegacySsoBrowser(handle);
  const staleBackendResponse = await fetch(staleBrowser.authorizationUrl, {
    redirect: "manual",
  });
  const staleCallbackLocation = staleBackendResponse.headers.get("location");
  if (!staleCallbackLocation) {
    throw new Error("Stale SSO transaction did not return a callback");
  }
  const staleCallbackRequest = new Request(staleCallbackLocation, {
    headers: { Cookie: staleBrowser.preLoginCookie },
  });
  const gateSuffix = crypto.randomUUID().replaceAll("-", "");
  const gateSequence = `ticket14_insert_gate_${gateSuffix}`;
  const gateFunction = `ticket14_insert_gate_fn_${gateSuffix}`;
  const gateTrigger = `ticket14_insert_gate_trigger_${gateSuffix}`;
  const gateLockKey = BigInt(
    `0x${crypto.randomUUID().replaceAll("-", "").slice(0, 15)}`
  );
  const staleCallbackState: { callback: Promise<Response> | null } = {
    callback: null,
  };
  try {
    await prisma.$executeRawUnsafe(`CREATE SEQUENCE public."${gateSequence}"`);
    await prisma.$executeRawUnsafe(`
      CREATE FUNCTION public."${gateFunction}"()
      RETURNS trigger
      LANGUAGE plpgsql
      AS $function$
      BEGIN
        PERFORM nextval('"public"."${gateSequence}"');
        PERFORM pg_advisory_xact_lock(${gateLockKey}::bigint);
        RETURN NEW;
      END;
      $function$;
    `);
    await prisma.$executeRawUnsafe(`
      CREATE TRIGGER "${gateTrigger}"
      BEFORE INSERT ON "user"
      FOR EACH ROW
      WHEN (NEW.email = '${email}')
      EXECUTE FUNCTION public."${gateFunction}"()
    `);
    const sessionCountBefore = await prisma.session.count({
      where: { isSso: true, userId: candidate.id },
    });
    await prisma.$transaction(async (tx) => {
      await tx.$executeRawUnsafe(
        `SELECT pg_advisory_xact_lock(${gateLockKey}::bigint)`
      );
      staleCallbackState.callback = Promise.resolve(
        handle(staleCallbackRequest)
      );
      let callbackReachedUserInsert = false;
      const gateDeadline = Date.now() + 10_000;
      while (!callbackReachedUserInsert && Date.now() < gateDeadline) {
        const [gate] = await tx.$queryRawUnsafe<{ is_called: boolean }[]>(
          `SELECT is_called FROM public."${gateSequence}"`
        );
        callbackReachedUserInsert = gate?.is_called === true;
        if (!callbackReachedUserInsert) {
          await Bun.sleep(10);
        }
      }
      expect(callbackReachedUserInsert).toBe(true);
      const [generationRow] = await tx.$queryRaw<{ generation: bigint }[]>`
        SELECT nextval('"legacy_sso_generation_seq"') AS generation
      `;
      if (!generationRow) {
        throw new Error("Database did not return SSO generation");
      }
      await tx.legacyAccountLinkRequest.create({
        data: {
          email,
          providerId,
          reviewedAt: new Date(),
          reviewedGeneration: generationRow.generation,
          status: "linked",
          subject,
          userId: candidate.id,
        },
      });
      await tx.account.create({
        data: {
          accountId: subject,
          id: crypto.randomUUID(),
          issuer: providerId,
          providerId,
          userId: candidate.id,
        },
      });
    });
    const startedCallback = staleCallbackState.callback;
    if (!startedCallback) {
      throw new Error("Stale SSO callback did not start");
    }
    const staleResponse = await startedCallback;
    expect(staleResponse.headers.get("location")).toBe(legacySsoFailedLocation);
    expect(
      await prisma.session.count({
        where: { isSso: true, userId: candidate.id },
      })
    ).toBe(sessionCountBefore);
  } finally {
    const callbackToSettle = staleCallbackState.callback;
    if (callbackToSettle) {
      await callbackToSettle.catch(() => {});
    }
    await prisma.$executeRawUnsafe(
      `DROP TRIGGER IF EXISTS "${gateTrigger}" ON "user"`
    );
    await prisma.$executeRawUnsafe(
      `DROP FUNCTION IF EXISTS public."${gateFunction}"()`
    );
    await prisma.$executeRawUnsafe(
      `DROP SEQUENCE IF EXISTS public."${gateSequence}"`
    );
  }
});
