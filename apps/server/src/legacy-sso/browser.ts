import { randomBytes, createHash } from "node:crypto";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { auth } from "@onlyoffice/auth";
import { prisma } from "@onlyoffice/db";
import { env } from "@onlyoffice/env/server";

import type { Identity } from "../auth/identity";
import { identityFor } from "../auth/identity";
import { fail } from "../http/errors";
import { readJsonRecord } from "../http/input";
import { corsOrigin, originOf } from "../http/origins";
import { handoffCodeMaximumLength } from "../prefill/input";
import {
  nextLegacySsoGeneration,
  exchangeLegacyCode,
  resolveLegacySsoAccount,
} from "./accounts";
import {
  legacySsoSessionIdentifier,
  hostOnlySsoCookie,
  legacySsoSwitchCookieName,
  safeLegacyFormReturnPath,
  legacySsoTransactionLifetimeSeconds,
  legacySsoTransactionIdentifier,
  legacySsoSwitchIdentifier,
  legacySsoCookieName,
  legacySsoFailureResponse,
  cookieValueFor,
  legacySsoTransaction,
  legacySsoPendingResponse,
  legacySsoSwitchLifetimeSeconds,
  legacySsoSessionCookieName,
} from "./config";
import type {
  LegacySsoSwitchChallenge,
  LegacySsoConfig,
  LegacySsoTransaction,
} from "./config";

const legacySsoSessionTransferLifetimeSeconds = 60;

async function createSsoSessionTransfer(
  userId: string,
  now: Date
): Promise<string> {
  const authContext = await auth.$context;
  const session = await authContext.internalAdapter.createSession(
    userId,
    false,
    { isSso: true },
    true
  );
  const transferId = randomBytes(32).toString("base64url");
  try {
    await prisma.verification.create({
      data: {
        expiresAt: new Date(
          now.getTime() + legacySsoSessionTransferLifetimeSeconds * 1000
        ),
        id: transferId,
        identifier: legacySsoSessionIdentifier,
        value: session.token,
      },
    });
  } catch (error) {
    await authContext.internalAdapter.deleteSession(session.token);
    throw error;
  }
  return transferId;
}
function legacySsoSwitchChallenge(
  value: string
): LegacySsoSwitchChallenge | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return null;
  }
  const challenge = parsed as Record<string, unknown>;
  if (
    typeof challenge.initiatingSessionId !== "string" ||
    typeof challenge.initiatingUserId !== "string" ||
    typeof challenge.returnTo !== "string" ||
    typeof challenge.targetUserId !== "string"
  ) {
    return null;
  }
  return {
    initiatingSessionId: challenge.initiatingSessionId,
    initiatingUserId: challenge.initiatingUserId,
    returnTo: challenge.returnTo,
    targetUserId: challenge.targetUserId,
  };
}

function legacySsoSwitchCookie(value: string, maxAge: number): string {
  return hostOnlySsoCookie(legacySsoSwitchCookieName, value, maxAge, "None");
}

function legacySsoSwitchFailure(): globalThis.Response {
  return Response.json(
    { error: "unauthorized", message: "The account switch is unavailable" },
    { headers: { "Cache-Control": "no-store" }, status: 401 }
  );
}

export async function startLegacySso(
  request: Request,
  config: LegacySsoConfig,
  clock: () => Date,
  identity: Identity | null
): Promise<globalThis.Response> {
  if (identity && identity.role !== "user") {
    fail(403, "legacy_sso_user_required", "Only Users can switch accounts");
  }
  if (new URL(request.url).search) {
    fail(
      400,
      "invalid_return_path",
      "Return path must be sent in the JSON body"
    );
  }
  const input = await readJsonRecord(request, 8 * 1024);
  if (
    Object.keys(input).some((key) => key !== "returnTo") ||
    (input.returnTo !== undefined && typeof input.returnTo !== "string")
  ) {
    fail(400, "invalid_return_path", "Only a Form return path is accepted");
  }
  const returnTo =
    input.returnTo === undefined
      ? "/dashboard"
      : safeLegacyFormReturnPath(input.returnTo, config.callbackUrl);
  if (!returnTo) {
    fail(400, "invalid_return_path", "The Form return path is unsafe");
  }
  const state = randomBytes(32).toString("base64url");
  const codeVerifier = randomBytes(32).toString("base64url");
  const codeChallenge = createHash("sha256")
    .update(codeVerifier)
    .digest("base64url");
  const transactionId = randomBytes(32).toString("base64url");
  const now = clock();
  const expiresAt = new Date(
    now.getTime() + legacySsoTransactionLifetimeSeconds * 1000
  );
  await prisma.verification.deleteMany({
    where: {
      expiresAt: { lte: now },
      identifier: {
        in: [
          legacySsoTransactionIdentifier,
          legacySsoSessionIdentifier,
          legacySsoSwitchIdentifier,
        ],
      },
    },
  });
  await prisma.$transaction(async (tx) => {
    const generation = await nextLegacySsoGeneration(tx, config.providerId);
    await tx.verification.create({
      data: {
        expiresAt,
        id: transactionId,
        identifier: legacySsoTransactionIdentifier,
        value: JSON.stringify({
          callbackUrl: config.callbackUrl,
          clientId: config.clientId,
          codeVerifier,
          generation: generation.toString(),
          initiatingSessionId: identity?.sessionId,
          initiatingUserId: identity?.id,
          returnTo,
          state,
        } satisfies LegacySsoTransaction),
      },
    });
  });
  const authorizeUrl = new URL(config.authorizeUrl);
  authorizeUrl.searchParams.set("client_id", config.clientId);
  authorizeUrl.searchParams.set("redirect_uri", config.callbackUrl);
  authorizeUrl.searchParams.set("state", state);
  authorizeUrl.searchParams.set("code_challenge", codeChallenge);
  authorizeUrl.searchParams.set("code_challenge_method", "S256");
  const headers = new Headers({
    "Cache-Control": "no-store",
    "Referrer-Policy": "no-referrer",
  });
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(
      legacySsoCookieName,
      transactionId,
      legacySsoTransactionLifetimeSeconds,
      "None"
    )
  );
  headers.append("Set-Cookie", legacySsoSwitchCookie("", 0));
  return Response.json({ authorizationUrl: authorizeUrl.href }, { headers });
}

export async function completeLegacySso(
  request: Request,
  config: LegacySsoConfig,
  clock: () => Date
): Promise<globalThis.Response> {
  const requestUrl = new URL(request.url);
  const forwardedProtocol = request.headers.get("x-forwarded-proto")?.trim();
  const publicOrigin = forwardedProtocol
    ? `${forwardedProtocol}://${requestUrl.host}`
    : requestUrl.origin;
  const callbackUrl = new URL(config.callbackUrl);
  const queryKeys = [...requestUrl.searchParams.keys()];
  const codeValues = requestUrl.searchParams.getAll("code");
  const stateValues = requestUrl.searchParams.getAll("state");
  if (
    `${publicOrigin}${requestUrl.pathname}` !== callbackUrl.href ||
    queryKeys.length !== 2 ||
    !queryKeys.includes("code") ||
    !queryKeys.includes("state") ||
    codeValues.length !== 1 ||
    stateValues.length !== 1
  ) {
    return legacySsoFailureResponse();
  }
  const code = codeValues[0];
  const state = stateValues[0];
  const transactionId = cookieValueFor(request, legacySsoCookieName, 43);
  if (
    !code ||
    code.length > handoffCodeMaximumLength ||
    !state ||
    !/^[A-Za-z0-9_-]{43}$/u.test(state) ||
    !transactionId
  ) {
    return legacySsoFailureResponse();
  }
  const now = clock();
  const storedTransaction = await prisma.verification.findUnique({
    where: { id: transactionId },
  });
  const transaction = storedTransaction
    ? legacySsoTransaction(storedTransaction.value)
    : null;
  if (
    !storedTransaction ||
    storedTransaction.identifier !== legacySsoTransactionIdentifier ||
    storedTransaction.expiresAt.getTime() <= now.getTime() ||
    !transaction ||
    transaction.state !== state ||
    transaction.clientId !== config.clientId ||
    transaction.callbackUrl !== config.callbackUrl ||
    !safeLegacyFormReturnPath(transaction.returnTo, config.callbackUrl)
  ) {
    return legacySsoFailureResponse();
  }
  const sessionReturnTo =
    transaction.initiatingUserId && transaction.initiatingSessionId
      ? transaction.returnTo
      : undefined;
  const consumed = await prisma.verification.deleteMany({
    where: {
      expiresAt: { gt: now },
      id: transactionId,
      identifier: legacySsoTransactionIdentifier,
    },
  });
  if (consumed.count !== 1) {
    return legacySsoFailureResponse();
  }
  const identity = await exchangeLegacyCode(
    config,
    code,
    transaction.codeVerifier
  );
  if (!identity) {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  const accountResolution = await resolveLegacySsoAccount(
    config,
    identity,
    BigInt(transaction.generation)
  );
  if (!accountResolution) {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  if (accountResolution.kind === "pending") {
    return legacySsoPendingResponse(sessionReturnTo);
  }
  if (accountResolution.kind === "failed") {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  const user = accountResolution.user;
  if (!user.enabled || user.role !== "user") {
    return legacySsoFailureResponse(sessionReturnTo);
  }
  if (transaction.initiatingUserId && transaction.initiatingSessionId) {
    const initiatingSession = await prisma.session.findUnique({
      select: { expiresAt: true, userId: true },
      where: { id: transaction.initiatingSessionId },
    });
    if (
      !initiatingSession ||
      initiatingSession.userId !== transaction.initiatingUserId
    ) {
      return legacySsoFailureResponse(sessionReturnTo);
    }
    if (initiatingSession.expiresAt.getTime() > now.getTime()) {
      const challengeId = randomBytes(32).toString("base64url");
      await prisma.verification.create({
        data: {
          expiresAt: new Date(
            now.getTime() + legacySsoSwitchLifetimeSeconds * 1000
          ),
          id: challengeId,
          identifier: legacySsoSwitchIdentifier,
          value: JSON.stringify({
            initiatingSessionId: transaction.initiatingSessionId,
            initiatingUserId: transaction.initiatingUserId,
            returnTo: transaction.returnTo,
            targetUserId: user.id,
          } satisfies LegacySsoSwitchChallenge),
        },
      });
      const headers = new Headers({
        "Cache-Control": "no-store",
        Location: new URL("/legacy-sso/confirm", corsOrigin).href,
        "Referrer-Policy": "no-referrer",
      });
      headers.append(
        "Set-Cookie",
        hostOnlySsoCookie(legacySsoCookieName, "", 0)
      );
      headers.append(
        "Set-Cookie",
        legacySsoSwitchCookie(challengeId, legacySsoSwitchLifetimeSeconds)
      );
      return new Response(null, { headers, status: 303 });
    }
    if (user.id !== transaction.initiatingUserId) {
      return legacySsoFailureResponse(sessionReturnTo);
    }
  }
  const transferId = await createSsoSessionTransfer(user.id, now);
  const headers = new Headers({
    "Cache-Control": "no-store",
    Location: new URL(transaction.returnTo, corsOrigin).href,
    "Referrer-Policy": "no-referrer",
  });
  headers.append("Set-Cookie", hostOnlySsoCookie(legacySsoCookieName, "", 0));
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(
      legacySsoSessionCookieName,
      transferId,
      legacySsoSessionTransferLifetimeSeconds,
      "None"
    )
  );
  return new Response(null, { headers, status: 303 });
}

function legacySsoSessionUnavailableResponse(): globalThis.Response {
  const headers = new Headers({ "Cache-Control": "no-store" });
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSessionCookieName, "", 0, "None")
  );
  return Response.json(
    {
      error: "unauthorized",
      message: "The legacy SSO session handoff is unavailable",
    },
    { headers, status: 401 }
  );
}

export async function claimLegacySsoSession(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null
): Promise<globalThis.Response> {
  const requestOrigin = request.headers.get("origin");
  const allowedOrigins = [
    corsOrigin,
    originOf(env.BETTER_AUTH_URL),
    config ? originOf(config.callbackUrl) : null,
  ];
  if (!config || !requestOrigin || !allowedOrigins.includes(requestOrigin)) {
    return Response.json(
      {
        error: "forbidden",
        message: "The legacy SSO session handoff origin is invalid",
      },
      { headers: { "Cache-Control": "no-store" }, status: 403 }
    );
  }
  const transferId = cookieValueFor(request, legacySsoSessionCookieName, 43);
  if (!transferId) {
    return legacySsoSessionUnavailableResponse();
  }
  const now = clock();
  const transfer = await prisma.verification.findUnique({
    where: { id: transferId },
  });
  if (
    !transfer ||
    transfer.identifier !== legacySsoSessionIdentifier ||
    transfer.expiresAt.getTime() <= now.getTime()
  ) {
    return legacySsoSessionUnavailableResponse();
  }
  const consumed = await prisma.verification.deleteMany({
    where: {
      expiresAt: { gt: now },
      id: transferId,
      identifier: legacySsoSessionIdentifier,
    },
  });
  if (consumed.count !== 1) {
    return legacySsoSessionUnavailableResponse();
  }
  const session = await prisma.session.findUnique({
    select: {
      expiresAt: true,
      isSso: true,
      user: { select: { enabled: true, role: true } },
    },
    where: { token: transfer.value },
  });
  if (
    !session ||
    !session.isSso ||
    session.expiresAt.getTime() <= now.getTime() ||
    !session.user.enabled ||
    session.user.role !== "user"
  ) {
    return legacySsoSessionUnavailableResponse();
  }
  const headers = new Headers({ "Cache-Control": "no-store" });
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSessionCookieName, "", 0, "None")
  );
  return Response.json({ token: transfer.value }, { headers });
}
interface LegacySsoSwitchContext {
  challenge: LegacySsoSwitchChallenge;
  challengeId: string;
  confirmationFingerprint: string | null;
  identity: Identity;
  target: { email: string; id: string; name: string } | null;
}

function legacySsoConfirmationFingerprint(
  identity: Identity,
  challenge: LegacySsoSwitchChallenge,
  target: { email: string; id: string; name: string }
): string {
  return createHash("sha256")
    .update(
      JSON.stringify([
        identity.id,
        identity.sessionId,
        identity.email,
        identity.name,
        target.id,
        target.email,
        target.name,
        challenge.returnTo,
      ])
    )
    .digest("base64url");
}

async function legacySsoSwitchContext(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null
): Promise<LegacySsoSwitchContext | null> {
  const requestOrigin = request.headers.get("origin");
  const requestOriginAllowed =
    requestOrigin === null
      ? request.method === "GET"
      : config !== null &&
        [
          corsOrigin,
          originOf(env.BETTER_AUTH_URL),
          originOf(config.callbackUrl),
        ].includes(requestOrigin);
  const identity = await identityFor(request);
  const challengeId = cookieValueFor(request, legacySsoSwitchCookieName, 43);
  if (
    !config ||
    !requestOriginAllowed ||
    !identity ||
    identity.role !== "user" ||
    !challengeId
  ) {
    return null;
  }
  const now = clock();
  const stored = await prisma.verification.findUnique({
    where: { id: challengeId },
  });
  const challenge =
    stored?.identifier === legacySsoSwitchIdentifier &&
    stored.expiresAt.getTime() > now.getTime()
      ? legacySsoSwitchChallenge(stored.value)
      : null;
  if (
    !challenge ||
    challenge.initiatingUserId !== identity.id ||
    challenge.initiatingSessionId !== identity.sessionId ||
    !safeLegacyFormReturnPath(challenge.returnTo, config.callbackUrl)
  ) {
    return null;
  }
  const target = await prisma.user.findUnique({
    select: { email: true, enabled: true, id: true, name: true, role: true },
    where: { id: challenge.targetUserId },
  });
  const switchTarget =
    target?.enabled && target.role === "user"
      ? { email: target.email, id: target.id, name: target.name }
      : null;
  return {
    challenge,
    challengeId,
    confirmationFingerprint: switchTarget
      ? legacySsoConfirmationFingerprint(identity, challenge, switchTarget)
      : null,
    identity,
    target: switchTarget,
  };
}

export async function readLegacySsoSwitch(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null
): Promise<globalThis.Response> {
  const context = await legacySsoSwitchContext(request, clock, config);
  if (!context?.target || !context.confirmationFingerprint) {
    return legacySsoSwitchFailure();
  }
  const currentIdentity = context.identity;
  return Response.json(
    {
      confirmationFingerprint: context.confirmationFingerprint,
      current: { email: currentIdentity.email, name: currentIdentity.name },
      legacy: { email: context.target.email, name: context.target.name },
      returnTo: context.challenge.returnTo,
      sameUser: context.target.id === currentIdentity.id,
    },
    { headers: { "Cache-Control": "no-store" } }
  );
}

export async function consumeLegacySsoSwitch(
  request: Request,
  clock: () => Date,
  config: LegacySsoConfig | null,
  confirm: boolean
): Promise<globalThis.Response> {
  const context = await legacySsoSwitchContext(request, clock, config);
  if (!context?.target || !context.confirmationFingerprint) {
    return legacySsoSwitchFailure();
  }
  const input = await readJsonRecord(request, 1024);
  if (
    Object.keys(input).length !== 1 ||
    typeof input.confirmationFingerprint !== "string" ||
    input.confirmationFingerprint !== context.confirmationFingerprint
  ) {
    return legacySsoSwitchFailure();
  }
  const now = clock();
  const consumed = await prisma.verification.deleteMany({
    where: {
      expiresAt: { gt: now },
      id: context.challengeId,
      identifier: legacySsoSwitchIdentifier,
    },
  });
  if (consumed.count !== 1) {
    return legacySsoSwitchFailure();
  }
  const headers = new Headers({ "Cache-Control": "no-store" });
  headers.append("Set-Cookie", legacySsoSwitchCookie("", 0));
  if (!confirm) {
    return Response.json({ cancelled: true }, { headers });
  }
  const transferId = await createSsoSessionTransfer(context.target.id, now);
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(
      legacySsoSessionCookieName,
      transferId,
      legacySsoSessionTransferLifetimeSeconds,
      "None"
    )
  );
  return Response.json({ returnTo: context.challenge.returnTo }, { headers });
}
