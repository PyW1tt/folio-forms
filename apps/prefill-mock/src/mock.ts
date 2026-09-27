import {
  createHash,
  randomBytes,
  timingSafeEqual,
} from "node:crypto";

type JsonRecord = Record<string, unknown>;

export const deterministicExternalRecord = {
  account: {
    active: true,
    address: { city: "Bangkok", country: "TH", postalCode: "10110" },
    contact: { email: "person@example.com", phone: "+66000000000" },
    id: "account-1",
  },
  person: { birthDate: "1990-01-02", name: "Mock Person" },
} as const;

export const deterministicSchemaItems = [
  { pointer: "/account/active", type: "boolean" },
  { pointer: "/account/address/city", type: "string" },
  { pointer: "/account/address/country", type: "string" },
  { pointer: "/account/address/postalCode", type: "string" },
  { pointer: "/account/contact/email", type: "string" },
  { pointer: "/account/contact/phone", type: "string" },
  { pointer: "/account/id", type: "string" },
  { pointer: "/person/birthDate", type: "string" },
  { pointer: "/person/name", type: "string" },
] as const;

export interface ExternalPrefillHandoffInput {
  email: string;
  externalReference: string;
  publicId: string;
  values: JsonRecord;
}

export interface ExternalPrefillStatusInput {
  externalReference: string;
}

export interface FolioHandoffConnector {
  createHandoff: (input: ExternalPrefillHandoffInput) => Promise<Response>;
  getStatus: (input: ExternalPrefillStatusInput) => Promise<Response>;
}

export interface PrefillMockHandlerOptions {
  connector: FolioHandoffConnector;
  folioOrigin: string;
}

export interface PrefillMockServer {
  readonly url: string;
  close: () => void;
}

const htmlEscape = (value: string): string =>
  value.replaceAll(
    /[&<>"']/gu,
    (character) =>
      ({
        '"': "&quot;",
        "&": "&amp;",
        "'": "&#39;",
        "<": "&lt;",
        ">": "&gt;",
      })[character] ?? character
  );

const jsonObject = (value: unknown): JsonRecord | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as JsonRecord)
    : null;

const formCode = async (request: Request): Promise<string | null> => {
  const form = await request.formData();
  const entries = [...form.entries()];
  if (
    entries.length !== 1 ||
    entries[0]?.[0] !== "code" ||
    typeof entries[0][1] !== "string"
  ) {
    return null;
  }
  return entries[0][1];
};

const launchForm = (folioOrigin: string, code: string): Response => {
  const action = `${folioOrigin.replace(/\/$/u, "")}/prefill/handoff`;
  return new Response(
    `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Continue to Folio Forms</title></head><body><form method="post" action="${htmlEscape(action)}"><input type="hidden" name="code" value="${htmlEscape(code)}"><button type="submit">Continue to Folio Forms</button></form></body></html>`,
    { headers: { "Content-Type": "text/html; charset=utf-8" } }
  );
};

export const createPrefillMockHandler =
  (
    options: PrefillMockHandlerOptions
  ): ((request: Request) => Response | Promise<Response>) =>
  async (request) => {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/schema") {
      const query = url.searchParams.get("q")?.trim().toLowerCase() ?? "";
      const items = query
        ? deterministicSchemaItems.filter((item) =>
            item.pointer.toLowerCase().includes(query)
          )
        : deterministicSchemaItems;
      return Response.json({ items, nextCursor: null });
    }
    if (request.method === "POST" && url.pathname === "/handoffs") {
      const input = jsonObject(await request.json());
      if (!input) {
        return Response.json({ error: "invalid_request" }, { status: 400 });
      }
      return options.connector.createHandoff(
        input as unknown as ExternalPrefillHandoffInput
      );
    }
    if (request.method === "POST" && url.pathname === "/launch") {
      const code = await formCode(request);
      return code
        ? launchForm(options.folioOrigin, code)
        : Response.json({ error: "invalid_request" }, { status: 400 });
    }
    if (request.method === "POST" && url.pathname === "/status") {
      const input = jsonObject(await request.json());
      if (!input) {
        return Response.json({ error: "invalid_request" }, { status: 400 });
      }
      return options.connector.getStatus(
        input as unknown as ExternalPrefillStatusInput
      );
    }
    return Response.json({ error: "not_found" }, { status: 404 });
  };

export const createFetchFolioConnector = (
  folioOrigin: string,
  secret: string
): FolioHandoffConnector => {
  const origin = folioOrigin.replace(/\/$/u, "");
  const headers = {
    "Content-Type": "application/json",
    "X-Prefill-Handoff-Secret": secret,
  };
  return {
    createHandoff: (input) =>
      fetch(`${origin}/api/integrations/prefill/handoffs`, {
        body: JSON.stringify(input),
        headers,
        method: "POST",
      }),
    getStatus: (input) =>
      fetch(`${origin}/api/integrations/prefill/status`, {
        body: JSON.stringify(input),
        headers,
        method: "POST",
      }),
  };
};

export const createInProcessFolioConnector = (
  handle: (request: Request) => Response | Promise<Response>,
  secret: string
): FolioHandoffConnector => {
  const request = (path: string, input: object) =>
    Promise.resolve(
      handle(
        new Request(`http://folio.local${path}`, {
          body: JSON.stringify(input),
          headers: {
            "Content-Type": "application/json",
            "X-Prefill-Handoff-Secret": secret,
          },
          method: "POST",
        })
      )
    );
  return {
    createHandoff: (input) =>
      request("/api/integrations/prefill/handoffs", input),
    getStatus: (input) => request("/api/integrations/prefill/status", input),
  };
};

export const startPrefillMock = (
  options: PrefillMockHandlerOptions & { port?: number }
): PrefillMockServer => {
  const server = Bun.serve({
    fetch: createPrefillMockHandler(options),
    port: options.port ?? 0,
  });
  return {
    close: () => server.stop(true),
    url: server.url.origin,
  };
};

export interface LegacySsoMockIdentity {
  email: string;
  email_verified: boolean;
  sub: string;
}

export interface LegacySsoMockOptions {
  callbackUrl: string;
  clientId: string;
  clientSecret: string;
  identity: LegacySsoMockIdentity;
  authorizationCode?: string;
  codeLifetimeMs?: number;
  clock?: () => Date;
}

export interface LegacySsoMockServer {
  authorizeUrl: string;
  close: () => void;
  exchangeUrl: string;
}

interface LegacySsoCode {
  callbackUrl: string;
  challenge: string;
  clientId: string;
  expiresAt: number;
  used: boolean;
}

const invalidLegacySsoGrant = () =>
  Response.json({ error: "invalid_grant" }, { status: 400 });

export const createLegacySsoMockHandler = (
  options: LegacySsoMockOptions
): ((request: Request) => Promise<Response>) => {
  const codes = new Map<string, LegacySsoCode>();
  const clock = options.clock ?? (() => new Date());
  const expectedAuthorization = createHash("sha256")
    .update(
      `Basic ${Buffer.from(
        `${options.clientId}:${options.clientSecret}`
      ).toString("base64")}`
    )
    .digest();
  return async (request) => {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/authorize") {
      const query = url.searchParams;
      if (
        [...query.keys()].length !== 5 ||
        [...query.keys()].some(
          (key) =>
            key !== "client_id" &&
            key !== "redirect_uri" &&
            key !== "state" &&
            key !== "code_challenge" &&
            key !== "code_challenge_method"
        ) ||
        query.getAll("client_id").length !== 1 ||
        query.getAll("redirect_uri").length !== 1 ||
        query.getAll("state").length !== 1 ||
        query.getAll("code_challenge").length !== 1 ||
        query.get("client_id") !== options.clientId ||
        query.get("redirect_uri") !== options.callbackUrl ||
        !/^[A-Za-z0-9_-]{43}$/u.test(query.get("state") ?? "") ||
        !/^[A-Za-z0-9_-]{43}$/u.test(query.get("code_challenge") ?? "") ||
        query.get("code_challenge_method") !== "S256"
      ) {
        return Response.json({ error: "invalid_request" }, { status: 400 });
      }
      const code =
        options.authorizationCode ?? randomBytes(32).toString("base64url");
      codes.set(code, {
        callbackUrl: options.callbackUrl,
        challenge: query.get("code_challenge") ?? "",
        clientId: options.clientId,
        expiresAt:
          clock().getTime() + (options.codeLifetimeMs ?? 30_000),
        used: false,
      });
      const callback = new URL(options.callbackUrl);
      callback.searchParams.set("code", code);
      callback.searchParams.set("state", query.get("state") ?? "");
      return new Response(null, {
        headers: {
          "Cache-Control": "no-store",
          Location: callback.href,
        },
        status: 303,
      });
    }
    if (request.method !== "POST" || url.pathname !== "/exchange") {
      return Response.json({ error: "not_found" }, { status: 404 });
    }
    const authorization = request.headers.get("authorization") ?? "";
    const receivedAuthorization = createHash("sha256")
      .update(authorization)
      .digest();
    if (!timingSafeEqual(expectedAuthorization, receivedAuthorization)) {
      return invalidLegacySsoGrant();
    }
    if (
      request.headers.get("content-type")?.split(";", 1)[0] !==
      "application/x-www-form-urlencoded"
    ) {
      return invalidLegacySsoGrant();
    }
    const contentLength = Number(request.headers.get("content-length") ?? "");
    if (Number.isFinite(contentLength) && contentLength > 8192) {
      return invalidLegacySsoGrant();
    }
    const body = await request.text();
    if (new TextEncoder().encode(body).byteLength > 8192) {
      return invalidLegacySsoGrant();
    }
    const form = new URLSearchParams(body);
    const code = form.get("code") ?? "";
    const issued = codes.get(code);
    const verifier = form.get("code_verifier") ?? "";
    const challenge = createHash("sha256")
      .update(verifier)
      .digest("base64url");
    if (
      form.get("grant_type") !== "authorization_code" ||
      form.get("client_id") !== options.clientId ||
      form.get("redirect_uri") !== options.callbackUrl ||
      !/^[A-Za-z0-9._~-]{43,128}$/u.test(verifier) ||
      !issued ||
      issued.used ||
      issued.expiresAt <= clock().getTime() ||
      issued.clientId !== options.clientId ||
      issued.callbackUrl !== options.callbackUrl ||
      issued.challenge !== challenge
    ) {
      return invalidLegacySsoGrant();
    }
    issued.used = true;
    return Response.json(options.identity, {
      headers: { "Cache-Control": "no-store" },
    });
  };
};

export const startLegacySsoMock = (
  options: LegacySsoMockOptions & { port?: number }
): LegacySsoMockServer => {
  const server = Bun.serve({
    fetch: createLegacySsoMockHandler(options),
    port: options.port ?? 0,
  });
  return {
    authorizeUrl: new URL("/authorize", server.url).href,
    close: () => server.stop(true),
    exchangeUrl: new URL("/exchange", server.url).href,
  };
};
