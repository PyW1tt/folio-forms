// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { requireIdentity } from "../auth/identity";
import { renewEditorLease, releaseEditorLease } from "../editor/leases";
import { validateId } from "../http/input";
import type { RouteDependencies } from "./dependencies";

export function registerEditorLeaseRoutes(
  app: Elysia,
  _dependencies: Pick<RouteDependencies, never>
): void {
  app
    .post("/api/editor-leases/:id/renew", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Editor lease");
      const lease = await renewEditorLease(identity, params.id);
      return {
        lease: { expiresAt: lease.expiresAt.toISOString(), id: lease.id },
      };
    })
    .delete("/api/editor-leases/:id", async ({ request, params }) => {
      const identity = await requireIdentity(request);
      validateId(params.id, "Editor lease");
      await releaseEditorLease(identity, params.id);
      return { ok: true };
    });
}
