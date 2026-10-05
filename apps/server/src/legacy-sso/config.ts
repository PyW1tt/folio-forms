// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { env } from "@onlyoffice/env/server";

import { idPattern } from "../http/input";
import { corsOrigin } from "../http/origins";

export const legacySsoTransactionLifetimeSeconds = 5 * 60;
export const legacySsoCookieName = "__Host-folio-sso";
export const legacySsoSessionCookieName = "__Host-folio-sso-session";
export const legacySsoSwitchCookieName = "__Host-folio-sso-switch";
export const legacySsoTransactionIdentifier = "legacy-sso-transaction";
export const legacySsoSessionIdentifier = "legacy-sso-session";
export const legacySsoSwitchIdentifier = "legacy-sso-switch";
export interface LegacySsoConfig {
  authorizeUrl: string;
  callbackUrl: string;
  clientId: string;
  clientSecret: string;
  exchangeUrl: string;
  providerId: string;
}
export interface LegacySsoTransaction {
  callbackUrl: string;
  clientId: string;
  codeVerifier: string;
  generation: string;
  initiatingSessionId?: string;
  initiatingUserId?: string;
  returnTo: string;
  state: string;
}
export interface LegacySsoSwitchChallenge {
  initiatingSessionId: string;
  initiatingUserId: string;
  returnTo: string;
  targetUserId: string;
}
export const legacySsoSwitchLifetimeSeconds = 5 * 60;

export function legacySsoConfigurationFromEnv(): LegacySsoConfig | null {
  const {
    LEGACY_SSO_AUTHORIZE_URL: authorizeUrl,
    LEGACY_SSO_CALLBACK_URL: callbackUrl,
    LEGACY_SSO_CLIENT_ID: clientId,
    LEGACY_SSO_CLIENT_SECRET: clientSecret,
    LEGACY_SSO_EXCHANGE_URL: exchangeUrl,
    LEGACY_SSO_PROVIDER_ID: providerId,
  } = env;
  const values = [
    authorizeUrl,
    callbackUrl,
    clientId,
    clientSecret,
    exchangeUrl,
    providerId,
  ];
  if (values.every((value) => value === undefined)) {
    return null;
  }
  if (
    typeof authorizeUrl !== "string" ||
    typeof callbackUrl !== "string" ||
    typeof clientId !== "string" ||
    typeof clientSecret !== "string" ||
    typeof exchangeUrl !== "string" ||
    typeof providerId !== "string"
  ) {
    throw new TypeError(
      "All six LEGACY_SSO settings must be configured together"
    );
  }
  return validatedLegacySsoConfiguration({
    authorizeUrl,
    callbackUrl,
    clientId,
    clientSecret,
    exchangeUrl,
    providerId,
  });
}

function parseLegacySsoEndpoint(value: string, label: string): URL {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`${label} must be an absolute URL`);
  }
  if (
    (url.protocol !== "https:" &&
      !(
        url.protocol === "http:" &&
        (url.hostname === "localhost" ||
          url.hostname === "127.0.0.1" ||
          url.hostname === "[::1]")
      )) ||
    url.username ||
    url.password ||
    url.hash
  ) {
    throw new Error(`${label} must use HTTPS without URL credentials`);
  }
  return url;
}

export function validatedLegacySsoConfiguration(
  config: LegacySsoConfig
): LegacySsoConfig {
  const authorize = parseLegacySsoEndpoint(
    config.authorizeUrl,
    "LEGACY_SSO_AUTHORIZE_URL"
  );
  const callback = parseLegacySsoEndpoint(
    config.callbackUrl,
    "LEGACY_SSO_CALLBACK_URL"
  );
  const exchange = parseLegacySsoEndpoint(
    config.exchangeUrl,
    "LEGACY_SSO_EXCHANGE_URL"
  );
  if (
    authorize.search ||
    exchange.search ||
    callback.search ||
    callback.pathname !== "/api/legacy-sso/callback" ||
    config.clientId.length === 0 ||
    config.clientId.length > 128 ||
    config.clientId.trim() !== config.clientId ||
    config.clientId.includes(":") ||
    config.clientSecret.length < 32 ||
    config.providerId.length === 0 ||
    config.providerId.length > 128 ||
    config.providerId.trim() !== config.providerId
  ) {
    throw new Error("Legacy SSO settings are invalid");
  }
  return {
    ...config,
    authorizeUrl: authorize.href,
    callbackUrl: callback.href,
    exchangeUrl: exchange.href,
  };
}

export function hasAsciiControlCharacters(value: string): boolean {
  for (const character of value) {
    const code = character.codePointAt(0);
    if (code !== undefined && (code <= 0x1f || code === 0x7f)) {
      return true;
    }
  }
  return false;
}

export function safeLegacyFormReturnPath(
  value: unknown,
  callbackUrl: string
): string | null {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > 2048 ||
    !value.startsWith("/") ||
    value.startsWith("//") ||
    value.includes("\\") ||
    hasAsciiControlCharacters(value)
  ) {
    return null;
  }
  let decoded: string;
  try {
    decoded = decodeURIComponent(value);
  } catch {
    return null;
  }
  if (
    decoded.startsWith("//") ||
    decoded.includes("\\") ||
    hasAsciiControlCharacters(decoded)
  ) {
    return null;
  }
  const callback = new URL(callbackUrl);
  let returnUrl: URL;
  try {
    returnUrl = new URL(value, callback.origin);
  } catch {
    return null;
  }
  const isFormPath = /^\/forms\/[0-9a-f]{32}\/fill$/u.test(returnUrl.pathname);
  const isDashboardPath =
    returnUrl.pathname === "/dashboard" && !returnUrl.search;
  if (
    returnUrl.origin !== callback.origin ||
    returnUrl.hash ||
    (!isFormPath && !isDashboardPath)
  ) {
    return null;
  }
  const queryKeys = [...returnUrl.searchParams.keys()];
  if (queryKeys.some((key) => key !== "responseId") || queryKeys.length > 1) {
    return null;
  }
  const responseIds = returnUrl.searchParams.getAll("responseId");
  if (
    responseIds.length > 1 ||
    (responseIds.length === 1 && !idPattern.test(responseIds[0] ?? ""))
  ) {
    return null;
  }
  return `${returnUrl.pathname}${returnUrl.search}`;
}

export function cookieValueFor(
  request: Request,
  name: string,
  maximumLength: number
): string | null {
  const cookieHeader = request.headers.get("cookie");
  if (!cookieHeader) {
    return null;
  }
  let match: string | null = null;
  for (const part of cookieHeader.split(";")) {
    const separator = part.indexOf("=");
    if (separator === -1 || part.slice(0, separator).trim() !== name) {
      continue;
    }
    const value = part.slice(separator + 1).trim();
    if (
      match !== null ||
      value.length === 0 ||
      value.length > maximumLength ||
      !/^[A-Za-z0-9_-]+$/u.test(value)
    ) {
      return null;
    }
    match = value;
  }
  return match;
}

export function hostOnlySsoCookie(
  name: string,
  value: string,
  maxAge: number,
  sameSite: "Lax" | "None" = "Lax"
): string {
  return `${name}=${value}; Path=/; Max-Age=${maxAge}; HttpOnly; Secure; SameSite=${sameSite}`;
}

function legacySsoRedirectResponse(
  result: "failed" | "pending",
  returnTo?: string
): globalThis.Response {
  const location = new URL(`/login?legacySso=${result}`, corsOrigin);
  if (returnTo) {
    location.searchParams.set("returnTo", returnTo);
  }
  const headers = new Headers({
    "Cache-Control": "no-store",
    Location: location.href,
    "Referrer-Policy": "no-referrer",
  });
  headers.append("Set-Cookie", hostOnlySsoCookie(legacySsoCookieName, "", 0));
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSessionCookieName, "", 0, "None")
  );
  headers.append(
    "Set-Cookie",
    hostOnlySsoCookie(legacySsoSwitchCookieName, "", 0, "None")
  );
  return new Response(null, { headers, status: 303 });
}
export function legacySsoFailureResponse(
  returnTo?: string
): globalThis.Response {
  return legacySsoRedirectResponse("failed", returnTo);
}
export function legacySsoPendingResponse(
  returnTo?: string
): globalThis.Response {
  return legacySsoRedirectResponse("pending", returnTo);
}

export function legacySsoTransaction(
  value: string
): LegacySsoTransaction | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return null;
  }
  const transaction = parsed as Record<string, unknown>;
  if (
    typeof transaction.callbackUrl !== "string" ||
    typeof transaction.clientId !== "string" ||
    typeof transaction.codeVerifier !== "string" ||
    typeof transaction.generation !== "string" ||
    !/^[1-9]\d*$/u.test(transaction.generation) ||
    typeof transaction.returnTo !== "string" ||
    typeof transaction.state !== "string" ||
    !/^[A-Za-z0-9._~-]{43,128}$/u.test(transaction.codeVerifier) ||
    !/^[A-Za-z0-9_-]{43}$/u.test(transaction.state) ||
    (transaction.initiatingUserId !== undefined &&
      typeof transaction.initiatingUserId !== "string") ||
    (transaction.initiatingSessionId !== undefined &&
      typeof transaction.initiatingSessionId !== "string") ||
    (typeof transaction.initiatingUserId === "string") !==
      (typeof transaction.initiatingSessionId === "string")
  ) {
    return null;
  }
  return {
    callbackUrl: transaction.callbackUrl,
    clientId: transaction.clientId,
    codeVerifier: transaction.codeVerifier,
    generation: transaction.generation,
    initiatingSessionId: transaction.initiatingSessionId as string | undefined,
    initiatingUserId: transaction.initiatingUserId as string | undefined,
    returnTo: transaction.returnTo,
    state: transaction.state,
  };
}
