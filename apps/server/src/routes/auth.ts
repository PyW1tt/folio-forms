import { auth } from "@onlyoffice/auth";
import { prisma } from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { identityFor } from "../auth/identity";
import {
  handleEmailSignIn,
  passwordMinimumLength,
  passwordMaximumLength,
  endAiAuthoringSessions,
} from "../auth/password-session";
import { fail } from "../http/errors";
import { asRecord } from "../http/input";
import type { RouteDependencies } from "./dependencies";

export function registerSessionRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "aiAuthoring" | "requestIp">
): void {
  const { aiAuthoring, requestIp } = dependencies;
  app
    .post("/api/auth/sign-in/email", ({ body, request, server }) => {
      const sourceIp =
        requestIp?.(request) ??
        server?.requestIP(request)?.address ??
        "unknown";
      return handleEmailSignIn(request, body, sourceIp, aiAuthoring);
    })
    .get("/api/session", async ({ request }) => {
      const identity = await identityFor(request);
      if (!identity) {
        fail(401, "unauthorized", "Authentication is required");
      }
      return {
        session: { expiresAt: identity.expiresAt },
        user: {
          email: identity.email,
          id: identity.id,
          mustChangePassword: identity.mustChangePassword && !identity.isSso,
          name: identity.name,
          role: identity.role,
        },
      };
    });
}

export function registerPasswordRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "aiAuthoring">
): void {
  const { aiAuthoring } = dependencies;
  app.post("/api/account/password", async ({ body, request }) => {
    const identity = await identityFor(request);
    if (!identity) {
      fail(401, "unauthorized", "Authentication is required");
    }
    const input = asRecord(body);
    const { currentPassword, newPassword } = input;
    if (
      typeof currentPassword !== "string" ||
      typeof newPassword !== "string"
    ) {
      fail(
        400,
        "invalid_request",
        "currentPassword and newPassword are required"
      );
    }
    if (newPassword.length < passwordMinimumLength) {
      fail(
        400,
        "password_too_short",
        `Password must contain at least ${passwordMinimumLength} characters`
      );
    }
    if (newPassword.length > passwordMaximumLength) {
      fail(
        400,
        "password_too_long",
        `Password must contain at most ${passwordMaximumLength} characters`
      );
    }
    if (
      currentPassword.length < passwordMinimumLength ||
      currentPassword.length > passwordMaximumLength
    ) {
      fail(400, "invalid_current_password", "Current password is invalid");
    }
    const authContext = await auth.$context;

    const account = await prisma.account.findFirst({
      where: { providerId: "credential", userId: identity.id },
    });
    const currentHash = account?.password;
    if (
      !account ||
      !currentHash ||
      !(await authContext.password.verify({
        hash: currentHash,
        password: currentPassword,
      }))
    ) {
      fail(400, "invalid_current_password", "Current password is invalid");
    }

    const passwordHash = await authContext.password.hash(newPassword);
    const revokedOwnerSessionIds = await prisma.$transaction(async (tx) => {
      const update = await tx.account.updateMany({
        data: { password: passwordHash },
        where: { id: account.id, password: currentHash },
      });
      if (update.count !== 1) {
        fail(409, "credential_changed", "Credential changed concurrently");
      }
      await tx.user.update({
        data: { mustChangePassword: false },
        where: { id: identity.id },
      });
      const sessions = await tx.session.findMany({
        select: { id: true },
        where: { userId: identity.id },
      });
      await tx.session.deleteMany({ where: { userId: identity.id } });
      return sessions.map(({ id }) => id);
    });
    await endAiAuthoringSessions(aiAuthoring, revokedOwnerSessionIds);
    return { ok: true };
  });
}

export function registerSessionExitRoutes(
  app: Elysia,
  dependencies: Pick<RouteDependencies, "aiAuthoring">
): void {
  const { aiAuthoring } = dependencies;
  app

    .post("/api/auth/sign-out", async ({ request }) => {
      const identity = await identityFor(request);
      if (identity) {
        let sessionRevoked = false;
        try {
          await aiAuthoring.endForSession(identity.sessionId, async () => {
            await prisma.session.deleteMany({
              where: { id: identity.sessionId },
            });
            sessionRevoked = true;
          });
        } catch (error) {
          if (!sessionRevoked) {
            throw error;
          }
          console.error(error);
          return Response.json(
            {
              error: "sign_out_cleanup_failed",
              message: "Session was revoked, but AI authoring cleanup failed",
              sessionRevoked: true,
            },
            { status: 500 }
          );
        }
      }
      return { ok: true };
    })
    .all("/api/auth/*", () =>
      Response.json(
        { error: "not_found", message: "Authentication route was not found" },
        { status: 404 }
      )
    );
}
