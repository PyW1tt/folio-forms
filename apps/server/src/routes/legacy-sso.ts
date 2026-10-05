// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { identityFor } from "../auth/identity";
import { fail } from "../http/errors";
import {
  startLegacySso,
  completeLegacySso,
  readLegacySsoSwitch,
  consumeLegacySsoSwitch,
  claimLegacySsoSession,
} from "../legacy-sso/browser";
import { legacySsoFailureResponse } from "../legacy-sso/config";
import type { RouteDependencies } from "./dependencies";

export function registerLegacySsoRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "handoffClock" | "legacySso">
): void {
  const { handoffClock, legacySso } = dependencies;
  app
    .get("/api/legacy-sso/status", () =>
      Response.json(
        { enabled: Boolean(legacySso) },
        { headers: { "Cache-Control": "no-store" } }
      )
    )
    .post(
      "/api/legacy-sso/start",
      async ({ request }) => {
        if (!legacySso) {
          fail(404, "legacy_sso_unavailable", "Legacy SSO is unavailable");
        }
        const identity = await identityFor(request);
        if (request.headers.has("authorization") && !identity) {
          fail(401, "unauthorized", "Authentication is required");
        }
        return startLegacySso(request, legacySso, handoffClock, identity);
      },
      { parse: "none" }
    )
    .get(
      "/api/legacy-sso/callback",
      ({ request }) =>
        legacySso
          ? completeLegacySso(request, legacySso, handoffClock)
          : legacySsoFailureResponse(),
      { parse: "none" }
    )
    .get(
      "/api/legacy-sso/switch",
      ({ request }) => readLegacySsoSwitch(request, handoffClock, legacySso),
      { parse: "none" }
    )
    .post(
      "/api/legacy-sso/switch/confirm",
      ({ request }) =>
        consumeLegacySsoSwitch(request, handoffClock, legacySso, true),
      { parse: "none" }
    )
    .post(
      "/api/legacy-sso/switch/cancel",
      ({ request }) =>
        consumeLegacySsoSwitch(request, handoffClock, legacySso, false),
      { parse: "none" }
    )
    .post(
      "/api/legacy-sso/session",
      ({ request }) => claimLegacySsoSession(request, handoffClock, legacySso),
      { parse: "none" }
    );
}
