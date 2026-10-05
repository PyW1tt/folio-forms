// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { authoringDisclosure } from "../ai-authoring";
import { requireIdentity, requireAdmin } from "../auth/identity";
import { isAuthoringOwnerSessionCurrent } from "../auth/password-session";
import { validateTemplateControls } from "../documents/fields";
import {
  readAuthoringCreationInput,
  readAuthoringPromptInput,
} from "../http/input";
import { DOCX_CONTENT_TYPE } from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerAiAuthoringRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "aiAuthoring">
): void {
  const { aiAuthoring } = dependencies;
  app
    .get("/api/admin/ai-authoring", async ({ request }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      return { disclosure: authoringDisclosure, enabled: aiAuthoring.enabled };
    })
    .post(
      "/api/admin/ai-authoring/sessions",
      async ({ request }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const reservation = aiAuthoring.reserveRequest(identity.sessionId);
        let sourcePdf: Uint8Array | undefined;
        try {
          const input = await readAuthoringCreationInput(request);
          sourcePdf = input.pdfBytes;
          const session = await aiAuthoring.create(
            { authSessionId: identity.sessionId, userId: identity.id },
            input.prompt,
            input.consent,
            (document, expected) =>
              validateTemplateControls(document, expected),
            reservation,
            () => isAuthoringOwnerSessionCurrent(request, identity),
            input.pdfBytes
          );
          return Response.json(
            { session },
            { headers: { "Cache-Control": "private, no-store" } }
          );
        } finally {
          sourcePdf?.fill(0);
          aiAuthoring.releaseRequest(reservation);
        }
      },
      { parse: "none" }
    )
    .get("/api/admin/ai-authoring/sessions/current", async ({ request }) => {
      const identity = await requireIdentity(request);
      requireAdmin(identity);
      const session = await aiAuthoring.current({
        authSessionId: identity.sessionId,
        userId: identity.id,
      });
      return Response.json(
        { session },
        { headers: { "Cache-Control": "private, no-store" } }
      );
    })
    .post(
      "/api/admin/ai-authoring/sessions/:sessionId/revisions",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const reservation = aiAuthoring.reserveRequest(identity.sessionId);
        try {
          const input = await readAuthoringPromptInput(request);
          const session = await aiAuthoring.revise(
            params.sessionId,
            { authSessionId: identity.sessionId, userId: identity.id },
            input.prompt,
            input.consent,
            (document, expected) =>
              validateTemplateControls(document, expected),
            reservation,
            () => isAuthoringOwnerSessionCurrent(request, identity)
          );
          return Response.json(
            { session },
            { headers: { "Cache-Control": "private, no-store" } }
          );
        } finally {
          aiAuthoring.releaseRequest(reservation);
        }
      },
      { parse: "none" }
    )
    .get(
      "/api/admin/ai-authoring/sessions/:sessionId/docx",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        const { bytes } = await aiAuthoring.download(params.sessionId, {
          authSessionId: identity.sessionId,
          userId: identity.id,
        });
        return new Response(bytes, {
          headers: {
            "Cache-Control": "private, no-store",
            "Content-Disposition":
              'attachment; filename="ai-authored-template.docx"',
            "Content-Type": DOCX_CONTENT_TYPE,
            "X-Content-Type-Options": "nosniff",
          },
        });
      }
    )
    .delete(
      "/api/admin/ai-authoring/sessions/:sessionId",
      async ({ request, params }) => {
        const identity = await requireIdentity(request);
        requireAdmin(identity);
        await aiAuthoring.end(params.sessionId, {
          authSessionId: identity.sessionId,
          userId: identity.id,
        });
        return { ok: true };
      }
    );
}
