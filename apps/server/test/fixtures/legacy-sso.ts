import { expect } from "bun:test";

import { prisma } from "@onlyoffice/db";
import { startLegacySsoMock } from "prefill-mock/mock";

import { createApp } from "../../src/app";
import type { LegacySsoConfig } from "../../src/legacy-sso/config";
import { jsonHeaders, createCredentialFixture, bearerFor } from "./http";

export type LegacySsoHttpHandler = (
  request: Request
) => Response | Promise<Response>;

interface LegacySsoBrowserTransaction {
  authorizationUrl: string;
  preLoginCookie: string;
  startResponse: Response;
}

export const createLegacySsoTestApp = (
  config: LegacySsoConfig,
  clock: () => Date = () => new Date()
): ReturnType<typeof createApp> =>
  createApp({
    clock,
    legacySso: config,
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.resolve(new TextEncoder().encode("%PDF-test")),
      forceSave: () => Promise.resolve(false),
    },
    prefillReturnUrl: "https://source.example.test/forms/return",
    requestIp: (request) => request.headers.get("x-test-ip"),
  });

export const cookieHeaderFrom = (response: Response, name: string): string => {
  const setCookie = response.headers.get("set-cookie") ?? "";
  for (const cookie of setCookie.split(/,\s*(?=[^;,]+=)/u)) {
    if (cookie.startsWith(`${name}=`)) {
      return cookie;
    }
  }
  throw new Error(`The ${name} cookie was not set`);
};

export const cookiePairFrom = (response: Response, name: string): string => {
  const pair = cookieHeaderFrom(response, name).split(";", 1)[0]?.trim();
  if (!pair) {
    throw new Error(`The ${name} cookie was not set`);
  }
  return pair;
};

export const startLegacySsoBrowser = async (
  handle: LegacySsoHttpHandler,
  returnTo?: string,
  bearer?: string
): Promise<LegacySsoBrowserTransaction> => {
  const startUrl = new URL(
    "/api/legacy-sso/start",
    "https://folio.example.test"
  );
  const headers = new Headers(jsonHeaders);
  if (bearer) {
    headers.set("Authorization", `Bearer ${bearer}`);
  }
  const startResponse = await handle(
    new Request(startUrl.href, {
      body: JSON.stringify(returnTo === undefined ? {} : { returnTo }),
      headers,
      method: "POST",
    })
  );
  const startBody = (await startResponse.json()) as {
    authorizationUrl?: string;
  };
  if (!startBody.authorizationUrl) {
    throw new Error("The legacy SSO start did not return an authorization URL");
  }
  return {
    authorizationUrl: startBody.authorizationUrl,
    preLoginCookie: cookiePairFrom(startResponse, "__Host-folio-sso"),
    startResponse,
  };
};

export const legacySsoCallbackUrl =
  "https://folio.example.test/api/legacy-sso/callback";

export const legacySsoWebOrigin = new URL(
  process.env.CORS_ORIGIN ?? "http://localhost:5173"
).origin;

export const legacySsoFailedLocation = new URL(
  "/login?legacySso=failed",
  legacySsoWebOrigin
).href;

export const legacySsoPendingLocation = new URL(
  "/login?legacySso=pending",
  legacySsoWebOrigin
).href;

export const completeLegacySsoCallback = async (
  handle: LegacySsoHttpHandler
): Promise<Response> => {
  const browserStart = await startLegacySsoBrowser(handle);
  const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(oldBackendResponse.status).toBe(303);
  const callbackLocation = oldBackendResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not return a callback");
  }
  return handle(
    new Request(callbackLocation, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
};

interface LegacyAccountLinkReviewSnapshot {
  email: string;
  id: string;
  providerId: string;
  reviewedGeneration: string | null;
  subject: string;
  user: {
    email: string;
    enabled: boolean;
    id: string;
    name: string;
    role: string;
  };
}

// oxlint-disable no-await-in-loop -- Preserve the original HTTP test's sequential review pagination.
export const legacyAccountLinkSnapshot = async (
  handle: LegacySsoHttpHandler,
  bearer: string,
  requestId: string
): Promise<LegacyAccountLinkReviewSnapshot> => {
  let cursor: string | null = null;
  do {
    const url = new URL("/api/admin/account-links", legacySsoCallbackUrl);
    if (cursor) {
      url.searchParams.set("cursor", cursor);
    }
    const response = await handle(
      new Request(url.href, {
        headers: { Authorization: `Bearer ${bearer}` },
      })
    );
    expect(response.status).toBe(200);
    const body = (await response.json()) as {
      nextCursor: string | null;
      requests: LegacyAccountLinkReviewSnapshot[];
    };
    const snapshot = body.requests.find((request) => request.id === requestId);
    if (snapshot) {
      return snapshot;
    }
    cursor = body.nextCursor;
  } while (cursor);
  throw new Error("The pending account link is missing from Admin review");
};
// oxlint-enable no-await-in-loop

export const submitLegacyAccountLinkReview = (
  handle: LegacySsoHttpHandler,
  bearer: string,
  snapshot: LegacyAccountLinkReviewSnapshot,
  decision: "approve" | "reject"
): Promise<Response> =>
  Promise.resolve(
    handle(
      new Request(
        new URL(
          `/api/admin/account-links/${snapshot.id}/${decision}`,
          legacySsoCallbackUrl
        ).href,
        {
          body: JSON.stringify({
            reviewedGeneration: snapshot.reviewedGeneration,
          }),
          headers: { ...jsonHeaders, Authorization: `Bearer ${bearer}` },
          method: "POST",
        }
      )
    )
  );

export const createPdmsReviewFixture = async (
  app: ReturnType<typeof createApp>
) => {
  const password = "PDMS-review-fixture-password";
  const candidate = await createCredentialFixture({
    email: `pdms-review-${crypto.randomUUID()}@example.com`,
    name: "Original review candidate",
    password,
  });
  const reviewer = await createCredentialFixture({
    email: `pdms-review-admin-${crypto.randomUUID()}@example.com`,
    name: "PDMS reviewer",
    password,
    role: "admin",
  });
  const reviewerBearer = await bearerFor(app, reviewer.email, password);
  const identity = {
    email: candidate.email,
    email_verified: false,
    sub: `pdms-${crypto.randomUUID()}`,
  };
  const providerId = `pdms-${crypto.randomUUID()}`;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `pdms-secret-${crypto.randomUUID()}`;
  let nextExchange: (() => Promise<Response>) | undefined;
  const mock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 60_000,
    exchangeResponse: () => {
      const delayed = nextExchange;
      nextExchange = undefined;
      return delayed ? delayed() : Response.json(identity);
    },
    identity,
  });
  try {
    const config = {
      authorizeUrl: mock.authorizeUrl,
      callbackUrl: legacySsoCallbackUrl,
      clientId,
      clientSecret,
      exchangeUrl: mock.exchangeUrl,
      providerId,
    };
    const ssoApp = createLegacySsoTestApp(config);
    const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
    const initialCallback = await completeLegacySsoCallback(handle);
    expect(initialCallback.headers.get("location")).toBe(
      legacySsoPendingLocation
    );
    const request = await prisma.legacyAccountLinkRequest.findUniqueOrThrow({
      where: { providerId_subject: { providerId, subject: identity.sub } },
    });
    return {
      candidate,
      config,
      delayNextExchange: () => {
        const entered = Promise.withResolvers<undefined>();
        const release = Promise.withResolvers<undefined>();
        nextExchange = async () => {
          const claims = { ...identity };
          entered.resolve();
          await release.promise;
          return Response.json(claims);
        };
        return {
          entered: entered.promise,
          release: () => release.resolve(),
        };
      },
      handle,
      identity,
      mock,
      password,
      request,
      reviewer,
      reviewerBearer,
    };
  } catch (error) {
    mock.close();
    throw error;
  }
};

export const completeAuthenticatedLegacySsoCallback = async (
  handle: LegacySsoHttpHandler,
  bearer: string,
  returnTo: string
): Promise<{
  browserStart: LegacySsoBrowserTransaction;
  callback: Response;
}> => {
  const browserStart = await startLegacySsoBrowser(handle, returnTo, bearer);
  const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
    redirect: "manual",
  });
  const callbackLocation = oldBackendResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not return a callback");
  }
  const callback = await handle(
    new Request(callbackLocation, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  return { browserStart, callback };
};

export const claimLegacySsoBearer = async (
  handle: LegacySsoHttpHandler,
  callbackResponse: Response
): Promise<string> => {
  const claimResponse = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: cookiePairFrom(callbackResponse, "__Host-folio-sso-session"),
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(claimResponse.status).toBe(200);
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
  return claimBody.token;
};
