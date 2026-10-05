// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { createFormFailureAudit } from "../audit/events";
import { HttpError, fail } from "../http/errors";
import { asRecord } from "../http/input";
import {
  createPrefillHandoff,
  launchPrefillHandoff,
} from "../prefill/create-launch";
import {
  prefillHandoffSecretMatches,
  handoffCreateInput,
  handoffStatusInput,
  requireTopLevelNavigation,
  readPrefillHandoffCode,
  pendingClaimCookie,
  pendingClaimLifetimeSeconds,
} from "../prefill/input";
import { pollPrefillHandoffStatus } from "../prefill/status";
import type { RouteDependencies } from "./dependencies";

export function registerPrefillEntryRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "prefillHandoffSecret" | "handoffClock">
): void {
  const { prefillHandoffSecret, handoffClock } = dependencies;
  app
    .post("/api/integrations/prefill/handoffs", async ({ body, request }) => {
      if (
        !prefillHandoffSecretMatches(
          prefillHandoffSecret,
          request.headers.get("x-prefill-handoff-secret")
        )
      ) {
        try {
          await createFormFailureAudit({
            action: "create_handoff",
            actorId: null,
            error: new HttpError(
              404,
              "handoff_unavailable",
              "The prefill handoff is unavailable"
            ),
            targetId: null,
          });
        } catch {
          // Preserve the non-enumerating response if the audit cannot be written.
        }
        fail(404, "handoff_unavailable", "The prefill handoff is unavailable");
      }
      try {
        return await createPrefillHandoff(handoffCreateInput(asRecord(body)));
      } catch (error) {
        if (error instanceof HttpError && error.code === "not_found") {
          fail(
            404,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          );
        }
        throw error;
      }
    })
    .post("/api/integrations/prefill/status", async ({ body, request }) => {
      if (
        !prefillHandoffSecretMatches(
          prefillHandoffSecret,
          request.headers.get("x-prefill-handoff-secret")
        )
      ) {
        fail(404, "handoff_unavailable", "The prefill handoff is unavailable");
      }
      try {
        return await pollPrefillHandoffStatus(
          handoffStatusInput(asRecord(body)).externalReference,
          handoffClock
        );
      } catch (error) {
        if (
          error instanceof HttpError &&
          error.code === "handoff_unavailable"
        ) {
          fail(
            404,
            "handoff_unavailable",
            "The prefill handoff is unavailable"
          );
        }
        throw error;
      }
    })
    .post(
      "/prefill/handoff",
      async ({ request }) => {
        try {
          requireTopLevelNavigation(request);
          const launch = await launchPrefillHandoff(
            await readPrefillHandoffCode(request),
            handoffClock
          );
          return new Response(null, {
            headers: {
              Location: `/forms/${launch.publicId}/fill`,
              "Set-Cookie": pendingClaimCookie(
                launch.claimToken,
                pendingClaimLifetimeSeconds
              ),
            },
            status: 303,
          });
        } catch (error) {
          try {
            await createFormFailureAudit({
              action: "launch_handoff",
              actorId: null,
              error,
              targetId: null,
            });
          } catch {
            // Preserve the retryable redirect if the audit cannot be written.
          }
          return new Response(null, {
            headers: {
              Location: "/handoff?error=handoff_unavailable",
              "Set-Cookie": pendingClaimCookie("", 0),
            },
            status: 303,
          });
        }
      },
      { parse: "none" }
    );
}

export function registerPrefillMethodRoute(
  app: Elysia,
  _dependencies: Pick<RouteDependencies, never>
): void {
  app.get("/prefill/handoff", () => {
    fail(
      405,
      "handoff_unavailable",
      "The prefill handoff must be launched with a top-level POST"
    );
  });
}
