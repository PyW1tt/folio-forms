import { expect, test } from "bun:test";
import { createHash } from "node:crypto";

import { createLegacySsoMockHandler } from "../src/mock";

test("exchanges a code before expiry in the body-read continuation and prevents replay", async () => {
  const callbackUrl = "http://mock.local/callback";
  const clientId = "body-read-client";
  const clientSecret = "body-read-client-secret";
  const verifier = "v".repeat(43);
  const identity = {
    email: "body-read@example.test",
    email_verified: true,
    sub: "body-read-subject",
  };
  let now = 0;
  const handler = createLegacySsoMockHandler({
    authorizationCode: "body-read-code",
    callbackUrl,
    clientId,
    clientSecret,
    clock: () => new Date(now),
    codeLifetimeMs: 1000,
    identity,
  });
  const query = new URLSearchParams({
    client_id: clientId,
    code_challenge: createHash("sha256").update(verifier).digest("base64url"),
    code_challenge_method: "S256",
    redirect_uri: callbackUrl,
    state: "s".repeat(43),
  });
  const authorizationResponse = await handler(
    new Request(`http://mock.local/authorize?${query}`)
  );
  expect(authorizationResponse.status).toBe(303);

  const body = Promise.withResolvers<string>();
  class ControlledBodyRequest extends Request {
    private readonly bodyText = body.promise;

    override text(): Promise<string> {
      return this.bodyText;
    }
  }
  const headers = {
    Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString("base64")}`,
    "Content-Type": "application/x-www-form-urlencoded",
  };
  const pending = handler(
    new ControlledBodyRequest("http://mock.local/exchange", {
      headers,
      method: "POST",
    })
  );
  const exchangeBody = new URLSearchParams({
    client_id: clientId,
    code: "body-read-code",
    code_verifier: verifier,
    grant_type: "authorization_code",
    redirect_uri: callbackUrl,
  }).toString();
  body.resolve(exchangeBody);
  queueMicrotask(() => {
    now = 1000;
  });
  const response = await pending;
  expect(response.status).toBe(200);
  expect(await response.json()).toEqual(identity);

  now = 0;
  const replay = await handler(
    new Request("http://mock.local/exchange", {
      body: exchangeBody,
      headers,
      method: "POST",
    })
  );
  expect(replay.status).toBe(400);
  expect(await replay.json()).toEqual({ error: "invalid_grant" });
});
