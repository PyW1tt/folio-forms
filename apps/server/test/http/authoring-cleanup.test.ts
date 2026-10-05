import { test, expect, vi } from "bun:test";
import { mkdtemp, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import nodePath from "node:path";

import { AiAuthoringSessions } from "../../src/ai-authoring";
import { createApp } from "../../src/app";
import { createCredentialFixture, bearerFor } from "../fixtures/http";

// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
const { basename, join } = nodePath;

const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  throw new Error(
    "The HTTP application test requires DATABASE_URL for an isolated PostgreSQL database"
  );
}
const convertedDocumentKeys: string[] = [];
const app = createApp({
  legacySso: null,
  onlyOffice: {
    convertDocxToPdf: (documentKey) => {
      convertedDocumentKeys.push(documentKey);
      return Promise.resolve(new TextEncoder().encode("%PDF-test"));
    },
    forceSave: () => Promise.resolve(false),
  },
  prefillReturnUrl: "https://source.example.test/forms/return",
  requestIp: (request) => request.headers.get("x-test-ip"),
});

test("Ticket 17 sign-out always revokes and clears tombstones after failures", async () => {
  const authoring = new AiAuthoringSessions(null);
  const endedOwnerSessions = Reflect.get(
    authoring,
    "endedOwnerSessions"
  ) as Set<string>;
  const { promise: revocationGate, resolve: releaseRevocation } =
    Promise.withResolvers<undefined>();
  const { promise: revocationStarted, resolve: notifyRevocationStarted } =
    Promise.withResolvers<undefined>();

  const ending = authoring.endForSession("normal-sign-out", async () => {
    notifyRevocationStarted();
    await revocationGate;
  });
  await revocationStarted;
  expect(endedOwnerSessions.has("normal-sign-out")).toBe(true);
  releaseRevocation();
  await ending;
  expect(endedOwnerSessions.size).toBe(0);
  const firstRevocationError = new Error("first concurrent revoke failed");
  const {
    promise: firstConcurrentRevocationGate,
    resolve: releaseFirstConcurrentRevocation,
  } = Promise.withResolvers<undefined>();
  const {
    promise: firstConcurrentRevocationStarted,
    resolve: notifyFirstConcurrentRevocation,
  } = Promise.withResolvers<undefined>();
  const firstConcurrentEnd = authoring.endForSession(
    "concurrent-sign-out",
    async () => {
      notifyFirstConcurrentRevocation();
      await firstConcurrentRevocationGate;
      throw firstRevocationError;
    }
  );
  await firstConcurrentRevocationStarted;
  const {
    promise: secondConcurrentRevocationGate,
    resolve: releaseSecondConcurrentRevocation,
  } = Promise.withResolvers<undefined>();
  const {
    promise: secondConcurrentRevocationStarted,
    resolve: notifySecondConcurrentRevocation,
  } = Promise.withResolvers<undefined>();
  const secondConcurrentEnd = authoring.endForSession(
    "concurrent-sign-out",
    async () => {
      notifySecondConcurrentRevocation();
      await secondConcurrentRevocationGate;
    }
  );
  await secondConcurrentRevocationStarted;
  releaseFirstConcurrentRevocation();
  let firstConcurrentErrorPropagated = false;
  try {
    await firstConcurrentEnd;
  } catch (error) {
    firstConcurrentErrorPropagated = error === firstRevocationError;
  }
  expect(firstConcurrentErrorPropagated).toBe(true);
  expect(endedOwnerSessions.has("concurrent-sign-out")).toBe(true);
  expect(() => authoring.reserveRequest("concurrent-sign-out")).toThrow();
  releaseSecondConcurrentRevocation();
  await secondConcurrentEnd;
  expect(endedOwnerSessions.has("concurrent-sign-out")).toBe(false);

  for (const ownerSessionId of ["normal-sign-out-2", "normal-sign-out-3"]) {
    await authoring.endForSession(ownerSessionId, async () => {});
  }
  expect(endedOwnerSessions.size).toBe(0);
  const sessions = Reflect.get(authoring, "sessions") as Map<
    string,
    {
      document: { document: Uint8Array };
      expiryTimer?: unknown;
      lastActivity: Date;
      ownerSessionId: string;
      piSession: { dispose: () => void };
      tempDirectory: string;
      toolState: {
        pending: {
          cancelled: boolean;
          done: Promise<void>;
          inspectionAbort: AbortController;
          started: boolean;
        };
      };
      turns: { assistantMessage: string; prompt: string }[];
      userId: string;
    }
  >;
  const cleanupError = new Error("local AI authoring cleanup failed");
  const cleanupDirectory = await mkdtemp(
    join(tmpdir(), "folio-authoring-cleanup-")
  );
  const cleanupDocument = new Uint8Array([1]);
  const cleanupTurns = [
    { assistantMessage: "private reply", prompt: "private prompt" },
  ];
  let cleanupTimerFired = false;
  vi.useFakeTimers();
  const cleanupTimer = setTimeout(() => {
    cleanupTimerFired = true;
  }, 20);
  sessions.set("cleanup-failure", {
    document: { document: cleanupDocument },
    expiryTimer: cleanupTimer,
    lastActivity: new Date(),
    ownerSessionId: "cleanup-failure",
    piSession: {
      dispose: () => {
        throw cleanupError;
      },
    },
    tempDirectory: cleanupDirectory,
    toolState: {
      pending: {
        cancelled: false,
        done: Promise.resolve(),
        inspectionAbort: new AbortController(),
        started: false,
      },
    },
    turns: cleanupTurns,
    userId: "user",
  });
  let revocationAttempted = false;
  let cleanupErrorPropagated = false;
  try {
    try {
      await authoring.endForSession("cleanup-failure", () => {
        revocationAttempted = true;
        return Promise.resolve();
      });
    } catch (error) {
      cleanupErrorPropagated = error === cleanupError;
    }
    expect(revocationAttempted).toBe(true);
    expect(cleanupErrorPropagated).toBe(true);
    expect(cleanupDocument).toEqual(new Uint8Array([0]));
    expect(cleanupTurns).toEqual([]);
    expect(sessions.size).toBe(0);
    vi.advanceTimersByTime(21);
    expect(cleanupTimerFired).toBe(false);
    expect(await readdir(tmpdir())).not.toContain(basename(cleanupDirectory));
    expect(endedOwnerSessions.has("cleanup-failure")).toBe(false);
  } finally {
    vi.useRealTimers();
    await rm(cleanupDirectory, { force: true, recursive: true });
  }

  for (const ownerSessionId of ["failed-sign-out", "failed-sign-out"]) {
    const revocationError = new Error("session revocation failed");
    let revocationErrorPropagated = false;
    try {
      await authoring.endForSession(ownerSessionId, () =>
        Promise.reject(revocationError)
      );
    } catch (error) {
      revocationErrorPropagated = error === revocationError;
    }
    expect(revocationErrorPropagated).toBe(true);
    expect(endedOwnerSessions.has(ownerSessionId)).toBe(false);
  }
  expect(endedOwnerSessions.size).toBe(0);
  const signOutPassword = "Ticket17-sign-out-cleanup-password";
  const signOutUser = await createCredentialFixture({
    email: `ticket-17-sign-out-cleanup-${crypto.randomUUID()}@example.com`,
    name: "Ticket 17 Sign-out Cleanup",
    password: signOutPassword,
  });
  let revocationFailureToken: string | undefined;
  const endForSession = vi.spyOn(
    AiAuthoringSessions.prototype,
    "endForSession"
  );
  try {
    endForSession.mockImplementationOnce(
      async (_ownerSessionId, revokeOwnerSession) => {
        await revokeOwnerSession?.();
        throw new Error("AI authoring cleanup failed after revocation");
      }
    );
    const cleanupFailureToken = await bearerFor(
      app,
      signOutUser.email,
      signOutPassword
    );
    const cleanupFailure = await app.handle(
      new Request("http://test.local/api/auth/sign-out", {
        headers: { Authorization: `Bearer ${cleanupFailureToken}` },
        method: "POST",
      })
    );
    expect(cleanupFailure.status).toBe(500);
    expect(await cleanupFailure.json()).toMatchObject({
      error: "sign_out_cleanup_failed",
      sessionRevoked: true,
    });
    const revokedSession = await app.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${cleanupFailureToken}` },
      })
    );
    expect(revokedSession.status).toBe(401);

    revocationFailureToken = await bearerFor(
      app,
      signOutUser.email,
      signOutPassword
    );
    endForSession.mockRejectedValueOnce(new Error("Session revocation failed"));
    const revocationFailure = await app.handle(
      new Request("http://test.local/api/auth/sign-out", {
        headers: { Authorization: `Bearer ${revocationFailureToken}` },
        method: "POST",
      })
    );
    expect(revocationFailure.status).toBe(500);
    const revocationFailureBody = await revocationFailure.json();
    expect(revocationFailureBody).toMatchObject({ error: "internal_error" });
    expect(revocationFailureBody).not.toHaveProperty("sessionRevoked");
    const activeSession = await app.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${revocationFailureToken}` },
      })
    );
    expect(activeSession.status).toBe(200);
  } finally {
    endForSession.mockRestore();
  }
  if (revocationFailureToken) {
    await app.handle(
      new Request("http://test.local/api/auth/sign-out", {
        headers: { Authorization: `Bearer ${revocationFailureToken}` },
        method: "POST",
      })
    );
  }
});

test("Ticket 17 authoring shutdown cleans every active session after errors", async () => {
  const authoring = new AiAuthoringSessions(null);
  const sessions = Reflect.get(authoring, "sessions") as Map<
    string,
    {
      document: { document: Uint8Array };
      expiryTimer?: unknown;
      lastActivity: Date;
      ownerSessionId: string;
      piSession: { dispose: () => void };
      tempDirectory: string;
      toolState: {
        pending: {
          cancelled: boolean;
          done: Promise<void>;
          inspectionAbort: AbortController;
          started: boolean;
        };
      };
      turns: { assistantMessage: string; prompt: string }[];
      userId: string;
    }
  >;
  const pendingGenerations = Reflect.get(
    authoring,
    "pendingGenerations"
  ) as Map<
    string,
    Set<{
      cancelled: boolean;
      disposeError?: { cause: unknown };
      done: Promise<void>;
      finish: () => void;
      piSession?: { abort: () => Promise<void> };
      reservation: { ownerSessionId: string };
      started: boolean;
    }>
  >;
  const firstDirectory = await mkdtemp(
    join(tmpdir(), "folio-authoring-close-first-")
  );
  const secondDirectory = await mkdtemp(
    join(tmpdir(), "folio-authoring-close-second-")
  );
  const firstDocument = new Uint8Array([1, 2]);
  const secondDocument = new Uint8Array([3, 4]);
  const firstTurns = [
    { assistantMessage: "first reply", prompt: "first prompt" },
  ];
  const secondTurns = [
    { assistantMessage: "second reply", prompt: "second prompt" },
  ];
  const firstExpiry = vi.fn();
  const secondExpiry = vi.fn();
  const secondDispose = vi.fn();
  const disposeError = new Error("Pi disposal failed");
  const pendingDisposeError = new Error("Pending Pi disposal failed");
  vi.useFakeTimers();
  try {
    sessions.set("close-first", {
      document: { document: firstDocument },
      expiryTimer: setTimeout(firstExpiry, 20),
      lastActivity: new Date(),
      ownerSessionId: "close-owner-first",
      piSession: {
        dispose: () => {
          throw disposeError;
        },
      },
      tempDirectory: firstDirectory,
      toolState: {
        pending: {
          cancelled: false,
          done: Promise.resolve(),
          inspectionAbort: new AbortController(),
          started: false,
        },
      },
      turns: firstTurns,
      userId: "first-user",
    });
    sessions.set("close-second", {
      document: { document: secondDocument },
      expiryTimer: setTimeout(secondExpiry, 20),
      lastActivity: new Date(),
      ownerSessionId: "close-owner-second",
      piSession: { dispose: secondDispose },
      tempDirectory: secondDirectory,
      toolState: {
        pending: {
          cancelled: false,
          done: Promise.resolve(),
          inspectionAbort: new AbortController(),
          started: false,
        },
      },
      turns: secondTurns,
      userId: "second-user",
    });

    const pendingReservation = authoring.reserveRequest("close-owner-pending");
    const [pendingGeneration] = [
      ...(pendingGenerations.get("close-owner-pending") ?? []),
    ];
    if (!pendingGeneration) {
      throw new Error("The pending shutdown generation was not reserved");
    }
    pendingGeneration.started = true;
    pendingGeneration.disposeError = { cause: pendingDisposeError };
    pendingGeneration.piSession = {
      abort: async () =>
        await Promise.resolve(authoring.releaseRequest(pendingReservation)),
    };
    let closeError: unknown;
    try {
      await authoring.close();
    } catch (error) {
      closeError = error;
    }
    if (!(closeError instanceof AggregateError)) {
      throw new Error(
        "AI authoring shutdown did not aggregate cleanup failures"
      );
    }
    expect(closeError.errors).toEqual([pendingDisposeError, disposeError]);
    expect(firstDocument).toEqual(new Uint8Array([0, 0]));
    expect(secondDocument).toEqual(new Uint8Array([0, 0]));
    expect(firstTurns).toEqual([]);
    expect(secondTurns).toEqual([]);
    expect(secondDispose).toHaveBeenCalledTimes(1);
    expect(sessions.size).toBe(0);
    vi.advanceTimersByTime(21);
    expect(firstExpiry).not.toHaveBeenCalled();
    expect(secondExpiry).not.toHaveBeenCalled();
    const temporaryEntries = await readdir(tmpdir());
    expect(temporaryEntries).not.toContain(basename(firstDirectory));
    expect(temporaryEntries).not.toContain(basename(secondDirectory));
  } finally {
    vi.useRealTimers();
    await Promise.all([
      rm(firstDirectory, { force: true, recursive: true }),
      rm(secondDirectory, { force: true, recursive: true }),
    ]);
  }
});
