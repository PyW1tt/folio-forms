// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { expect } from "bun:test";
import { createHash } from "node:crypto";

import { prisma } from "@onlyoffice/db";
import {
  startPrefillMock,
  createInProcessFolioConnector,
  deterministicExternalRecord,
} from "prefill-mock/mock";
import type { PrefillMockServer } from "prefill-mock/mock";

import { createApp } from "../../../src/app";
import {
  jsonHeaders,
  createCredentialFixture,
  bearerFor,
} from "../../fixtures/http";
import {
  createLegacySsoTestApp,
  legacySsoCallbackUrl,
} from "../../fixtures/legacy-sso";
import type { FieldConfigurationOutput } from "./editor-access";
import type { BootstrapAndCreationOutput } from "./setup";

export interface PrefillEntryInput {
  app: ReturnType<typeof createApp>;
  userEmail: BootstrapAndCreationOutput["userEmail"];
  password: BootstrapAndCreationOutput["password"];
  formId: BootstrapAndCreationOutput["formId"];
  selectedPointer: FieldConfigurationOutput["selectedPointer"];
}

export interface PrefillEntryOutput {
  user: { email: string; id: string };
  userId: string;
  userBearer: string;
  formRecord: { publicId: string };
  prefillHandoffSecret: string;
  mock: PrefillMockServer;
  handoffCandidateValues: {
    account: {
      active: boolean;
      address: {
        readonly city: "Bangkok";
        readonly country: "TH";
        readonly postalCode: "10110";
      };
      contact: {
        readonly email: "person@example.com";
        readonly phone: "+66000000000";
      };
      id: "account-1";
    };
    person: { name: string; birthDate: "1990-01-02" };
  };
  handoffCreateBody: { code?: string; launchPath?: string } & { code: string };
  missingValueBearer: string;
}

export const runPrefillEntry = async (
  input: PrefillEntryInput
): Promise<PrefillEntryOutput> => {
  const { app, userEmail, password, formId, selectedPointer } = input;

  const user = await createCredentialFixture({
    email: userEmail,
    name: "Ticket 04 Workflow User",
    password,
  });
  const userId = user.id;
  const userBearer = await bearerFor(app, userEmail, password);

  const formRecord = await prisma.form.findUnique({
    select: { publicId: true },
    where: { id: formId },
  });
  if (!formRecord) {
    throw new Error("The published test form was not found");
  }
  const prefillHandoffSecret = process.env.PREFILL_HANDOFF_SECRET ?? "";
  const mock = startPrefillMock({
    connector: createInProcessFolioConnector(
      (request) => app.handle(request),
      prefillHandoffSecret
    ),
    folioOrigin: "http://test.local",
  });
  try {
    const schemaResponse = await fetch(`${mock.url}/schema?q=person`);
    expect(schemaResponse.status).toBe(200);
    const schemaBody = (await schemaResponse.json()) as {
      items?: { pointer: string; type: string }[];
    };
    expect(schemaBody.items).toContainEqual({
      pointer: "/person/name",
      type: "string",
    });
    const handoffExternalReference = `ticket-17-reference-${crypto.randomUUID()}`;
    const handoffCandidateValues = {
      ...deterministicExternalRecord,
      account: { ...deterministicExternalRecord.account, active: true },
      person: {
        ...deterministicExternalRecord.person,
        name: "Ticket 17 Prefill",
      },
      [selectedPointer]: selectedPointer.endsWith("/active")
        ? true
        : "Ticket 17 Prefill",
    };
    const ordinaryStartBeforeHandoff = await app.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
        headers: { Authorization: `Bearer ${userBearer}` },
        method: "POST",
      })
    );
    expect(ordinaryStartBeforeHandoff.status).toBe(409);
    expect(await ordinaryStartBeforeHandoff.json()).toMatchObject({
      error: "prefill_required",
    });
    const invalidHandoffSecret = await app.handle(
      new Request("http://test.local/api/integrations/prefill/handoffs", {
        body: JSON.stringify({
          email: userEmail,
          externalReference: handoffExternalReference,
          publicId: formRecord.publicId,
          values: handoffCandidateValues,
        }),
        headers: {
          ...jsonHeaders,
          "X-Prefill-Handoff-Secret": "wrong-secret",
        },
        method: "POST",
      })
    );
    expect(invalidHandoffSecret.status).toBe(404);
    expect(await invalidHandoffSecret.json()).toMatchObject({
      error: "handoff_unavailable",
    });
    const handoffCreateResponse = await fetch(`${mock.url}/handoffs`, {
      body: JSON.stringify({
        email: ` ${userEmail.toUpperCase()} `,
        externalReference: handoffExternalReference,
        publicId: formRecord.publicId,
        values: handoffCandidateValues,
      }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(handoffCreateResponse.status).toBe(200);
    const handoffCreateBody = (await handoffCreateResponse.json()) as {
      code?: string;
      launchPath?: string;
    };
    if (!handoffCreateBody.code) {
      throw new Error("The external mock did not return a handoff code");
    }
    expect(handoffCreateBody.code).toMatch(/^[A-Za-z0-9_-]{43}$/u);
    expect(handoffCreateBody.launchPath).toBe("/prefill/handoff");
    const duplicateHandoffCreateResponse = await fetch(`${mock.url}/handoffs`, {
      body: JSON.stringify({
        email: userEmail,
        externalReference: handoffExternalReference,
        publicId: formRecord.publicId,
        values: handoffCandidateValues,
      }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(duplicateHandoffCreateResponse.status).toBe(409);
    expect(await duplicateHandoffCreateResponse.json()).toMatchObject({
      error: "handoff_unavailable",
    });
    const handoffRecord = await prisma.handoff.findFirstOrThrow({
      orderBy: { createdAt: "desc" },
      where: {
        externalReferenceDigest: createHash("sha256")
          .update(handoffExternalReference)
          .digest("hex"),
      },
    });
    expect(handoffRecord.codeDigest).toBe(
      createHash("sha256").update(handoffCreateBody.code).digest("hex")
    );
    expect(handoffRecord.codeDigest).not.toBe(handoffCreateBody.code);
    expect(handoffRecord.codeDigest).toHaveLength(64);
    expect(handoffRecord.expiresAt.getTime()).toBeGreaterThan(Date.now());
    expect(handoffRecord.expiresAt.getTime()).toBeLessThanOrEqual(
      Date.now() + 120_000 + 1000
    );
    expect(handoffRecord.filteredValues).toEqual({
      full_name: selectedPointer.endsWith("/active")
        ? true
        : "Ticket 17 Prefill",
    });
    const pendingStatusResponse = await fetch(`${mock.url}/status`, {
      body: JSON.stringify({ externalReference: handoffExternalReference }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(pendingStatusResponse.status).toBe(200);
    const pendingStatusBody = (await pendingStatusResponse.json()) as Record<
      string,
      unknown
    >;
    expect(Object.keys(pendingStatusBody).toSorted()).toEqual([
      "consumedAt",
      "createdAt",
      "expiresAt",
      "latestCorrectionNumber",
      "reservedAt",
      "status",
      "submittedAt",
      "updatedAt",
    ]);
    expect(pendingStatusBody).toMatchObject({
      latestCorrectionNumber: null,
      reservedAt: null,
      status: "pending",
      submittedAt: null,
    });
    expect(JSON.stringify(pendingStatusBody)).not.toContain(userEmail);
    expect(JSON.stringify(pendingStatusBody)).not.toContain(
      "Ticket 17 Prefill"
    );
    const unknownStatusResponse = await fetch(`${mock.url}/status`, {
      body: JSON.stringify({
        externalReference: `ticket-19-unknown-${crypto.randomUUID()}`,
      }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(unknownStatusResponse.status).toBe(404);
    expect(await unknownStatusResponse.json()).toMatchObject({
      error: "handoff_unavailable",
    });
    const invalidStatusSecretResponse = await app.handle(
      new Request("http://test.local/api/integrations/prefill/status", {
        body: JSON.stringify({ externalReference: handoffExternalReference }),
        headers: {
          ...jsonHeaders,
          "X-Prefill-Handoff-Secret": "wrong-secret",
        },
        method: "POST",
      })
    );
    expect(invalidStatusSecretResponse.status).toBe(404);
    expect(await invalidStatusSecretResponse.json()).toMatchObject({
      error: "handoff_unavailable",
    });
    const handoffReturnOverride = await app.handle(
      new Request("http://test.local/api/integrations/prefill/handoffs", {
        body: JSON.stringify({
          email: userEmail,
          externalReference: `ticket-19-return-override-${crypto.randomUUID()}`,
          publicId: formRecord.publicId,
          returnUrl: "https://attacker.example.test",
          values: handoffCandidateValues,
        }),
        headers: {
          ...jsonHeaders,
          "X-Prefill-Handoff-Secret": prefillHandoffSecret,
        },
        method: "POST",
      })
    );
    expect(handoffReturnOverride.status).toBe(400);
    expect(await handoffReturnOverride.json()).toMatchObject({
      error: "invalid_request",
    });
    const getLaunchResponse = await app.handle(
      new Request(
        `http://test.local${handoffCreateBody.launchPath}?code=${handoffCreateBody.code}`
      )
    );
    expect(getLaunchResponse.status).toBe(405);
    const launchPageResponse = await fetch(`${mock.url}/launch`, {
      body: new URLSearchParams({ code: handoffCreateBody.code }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    });
    expect(launchPageResponse.status).toBe(200);
    const launchPage = await launchPageResponse.text();
    expect(launchPage).toContain('action="http://test.local/prefill/handoff"');
    expect(launchPage).not.toContain("?code=");
    const launchCode = /name="code" value="(?<code>[^"]+)"/u.exec(launchPage)
      ?.groups?.code;
    expect(launchCode).toBe(handoffCreateBody.code);
    if (!launchCode) {
      throw new Error("The external mock did not render a handoff code");
    }
    const launchResponse = await app.handle(
      new Request("http://test.local/prefill/handoff", {
        body: new URLSearchParams({ code: launchCode }),
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "Sec-Fetch-Dest": "document",
          "Sec-Fetch-Mode": "navigate",
        },
        method: "POST",
      })
    );
    expect(launchResponse.status).toBe(303);
    expect(launchResponse.headers.get("location")).toBe(
      `/forms/${formRecord.publicId}/fill`
    );
    const pendingCookie = launchResponse.headers
      .get("set-cookie")
      ?.split(";", 1)[0];
    if (!pendingCookie) {
      throw new Error("The handoff launch did not set a pending claim cookie");
    }
    const launchCookieHeader = launchResponse.headers.get("set-cookie") ?? "";
    expect(launchCookieHeader).toContain("__Host-folio-pending-claim=");
    expect(launchCookieHeader).toContain("Path=/");
    expect(launchCookieHeader).toContain("Max-Age=600");
    expect(launchCookieHeader).toContain("HttpOnly");
    expect(launchCookieHeader).toContain("Secure");
    expect(launchCookieHeader).toContain("SameSite=Lax");
    const prefillBoundarySsoApp = createLegacySsoTestApp({
      authorizeUrl: "https://legacy.example.test/authorize",
      callbackUrl: legacySsoCallbackUrl,
      clientId: "ticket-12-prefill-boundary",
      clientSecret: "ticket-12-prefill-boundary-secret",
      exchangeUrl: "https://legacy.example.test/token",
      providerId: "ticket-12-prefill-boundary",
    });
    const separatedHandoff = await prefillBoundarySsoApp.handle(
      new Request("https://folio.example.test/api/legacy-sso/session", {
        headers: {
          Cookie: pendingCookie,
          Origin: "https://folio.example.test",
        },
        method: "POST",
      })
    );
    expect(separatedHandoff.status).toBe(401);
    const redeemedStartResponse = await app.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
        headers: {
          Authorization: `Bearer ${userBearer}`,
          Cookie: pendingCookie,
        },
        method: "POST",
      })
    );
    expect(redeemedStartResponse.status).toBe(200);
    const redeemedStartBody = (await redeemedStartResponse.json()) as {
      editorConfigUrl?: string;
      prefill?: {
        data?: Record<string, unknown>;
        editableFields?: Record<string, unknown>;
      };
      response?: { id?: string };
    };
    expect(redeemedStartBody.editorConfigUrl).toContain("action=fill");
    expect(redeemedStartBody.prefill?.data).toEqual({
      full_name: selectedPointer.endsWith("/active")
        ? true
        : "Ticket 17 Prefill",
    });
    expect(redeemedStartBody.prefill?.editableFields).toEqual({
      full_name: false,
    });
    const handoffResponseId = redeemedStartBody.response?.id;
    if (!handoffResponseId) {
      throw new Error("The handoff did not create a Response");
    }
    const redeemedSnapshot = await prisma.prefillSnapshot.findUniqueOrThrow({
      where: { responseId: handoffResponseId },
    });
    expect(redeemedSnapshot.values).toEqual({
      full_name: selectedPointer.endsWith("/active")
        ? true
        : "Ticket 17 Prefill",
    });
    expect(redeemedSnapshot.lockedFields).toEqual({ full_name: true });
    expect(
      await prisma.response.findUniqueOrThrow({
        select: { externalReferenceDigest: true },
        where: { id: handoffResponseId },
      })
    ).toEqual({
      externalReferenceDigest: createHash("sha256")
        .update(handoffExternalReference)
        .digest("hex"),
    });
    const replayedHandoffStart = await app.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
        headers: {
          Authorization: `Bearer ${userBearer}`,
          Cookie: pendingCookie,
        },
        method: "POST",
      })
    );
    expect(replayedHandoffStart.status).toBe(409);
    expect(await replayedHandoffStart.json()).toMatchObject({
      error: "handoff_unavailable",
    });
    const redeemedHandoffRecord = await prisma.handoff.findUniqueOrThrow({
      where: { id: handoffRecord.id },
    });
    expect(redeemedHandoffRecord.status).toBe("consumed");
    expect(redeemedHandoffRecord.responseId).toBe(handoffResponseId);
    const draftStatusResponse = await fetch(`${mock.url}/status`, {
      body: JSON.stringify({ externalReference: handoffExternalReference }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(draftStatusResponse.status).toBe(200);
    expect(await draftStatusResponse.json()).toMatchObject({
      latestCorrectionNumber: null,
      status: "draft",
      submittedAt: null,
    });
    const reentryExternalReference = `ticket-18-reentry-${crypto.randomUUID()}`;
    const reentryCreateResponse = await fetch(`${mock.url}/handoffs`, {
      body: JSON.stringify({
        email: userEmail,
        externalReference: reentryExternalReference,
        publicId: formRecord.publicId,
        values: {
          ...handoffCandidateValues,
          [selectedPointer]: "A newer external value",
        },
      }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(reentryCreateResponse.status).toBe(200);
    const reentryCode = (
      (await reentryCreateResponse.json()) as { code: string }
    ).code;
    const reentryPage = await fetch(`${mock.url}/launch`, {
      body: new URLSearchParams({ code: reentryCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    });
    const reentryPageHtml = await reentryPage.text();
    const reentryLaunchCode = /name="code" value="(?<code>[^"]+)"/u.exec(
      reentryPageHtml
    )?.groups?.code;
    if (!reentryLaunchCode) {
      throw new Error("The reentry mock did not render a launch form");
    }
    const reentryLaunch = await app.handle(
      new Request("http://test.local/prefill/handoff", {
        body: new URLSearchParams({ code: reentryLaunchCode }),
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "Sec-Fetch-Dest": "document",
          "Sec-Fetch-Mode": "navigate",
        },
        method: "POST",
      })
    );
    const reentryCookie = reentryLaunch.headers
      .get("set-cookie")
      ?.split(";", 1)[0];
    if (!reentryCookie) {
      throw new Error("The reentry handoff did not set a pending claim cookie");
    }
    const reentryHandoff = await prisma.handoff.findFirstOrThrow({
      where: {
        externalReferenceDigest: createHash("sha256")
          .update(reentryExternalReference)
          .digest("hex"),
      },
    });
    const reentryReservedAt = reentryHandoff.reservedAt;
    if (!reentryReservedAt) {
      throw new Error("The reentry handoff was not reserved");
    }
    const lateApp = createApp({
      clock: () => new Date(reentryReservedAt.getTime() + 121_000),
      prefillHandoffSecret,
    });
    const stableReentryStart = await lateApp.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
        headers: {
          Authorization: `Bearer ${userBearer}`,
          Cookie: reentryCookie,
        },
        method: "POST",
      })
    );
    expect(stableReentryStart.status).toBe(200);
    expect(await stableReentryStart.json()).toMatchObject({
      prefill: {
        data: {
          full_name: selectedPointer.endsWith("/active")
            ? true
            : "Ticket 17 Prefill",
        },
        editableFields: { full_name: false },
      },
      response: { id: handoffResponseId },
    });
    const stableReentrySnapshot =
      await prisma.prefillSnapshot.findUniqueOrThrow({
        where: { responseId: handoffResponseId },
      });
    expect(stableReentrySnapshot).toMatchObject({
      lockedFields: { full_name: true },
      values: {
        full_name: selectedPointer.endsWith("/active")
          ? true
          : "Ticket 17 Prefill",
      },
    });
    expect(
      await prisma.response.findUniqueOrThrow({
        select: { externalReferenceDigest: true },
        where: { id: handoffResponseId },
      })
    ).toEqual({
      externalReferenceDigest: createHash("sha256")
        .update(handoffExternalReference)
        .digest("hex"),
    });
    const missingValueEmail = `ticket-18-missing-${crypto.randomUUID()}@example.com`;
    const missingValueUser = await createCredentialFixture({
      email: missingValueEmail,
      name: "Ticket 18 Missing Value User",
      password,
    });
    const missingValueBearer = await bearerFor(
      app,
      missingValueEmail,
      password
    );
    const missingValueReference = `ticket-18-missing-${crypto.randomUUID()}`;
    const missingValueCreate = await fetch(`${mock.url}/handoffs`, {
      body: JSON.stringify({
        email: missingValueEmail,
        externalReference: missingValueReference,
        publicId: formRecord.publicId,
        values: { account: { active: true } },
      }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(missingValueCreate.status).toBe(200);
    const missingValueCode = (
      (await missingValueCreate.json()) as { code: string }
    ).code;
    const missingValueLaunch = await app.handle(
      new Request("http://test.local/prefill/handoff", {
        body: new URLSearchParams({ code: missingValueCode }),
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "Sec-Fetch-Dest": "document",
          "Sec-Fetch-Mode": "navigate",
        },
        method: "POST",
      })
    );
    const missingValueCookie = missingValueLaunch.headers
      .get("set-cookie")
      ?.split(";", 1)[0];
    if (!missingValueCookie) {
      throw new Error("The missing-value handoff did not set a cookie");
    }
    const missingValueStart = await app.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
        headers: {
          Authorization: `Bearer ${missingValueBearer}`,
          Cookie: missingValueCookie,
        },
        method: "POST",
      })
    );
    expect(missingValueStart.status).toBe(200);
    expect(await missingValueStart.json()).toMatchObject({
      prefill: { data: {}, editableFields: {} },
    });
    const missingValueResponse = await prisma.response.findUniqueOrThrow({
      include: { prefillSnapshot: true },
      where: {
        formId_userId: { formId, userId: missingValueUser.id },
      },
    });
    expect(missingValueResponse.prefillSnapshot).toMatchObject({
      lockedFields: {},
      values: {},
    });
    const discardedMissingValue = await app.handle(
      new Request(
        `http://test.local/api/responses/${missingValueResponse.id}`,
        {
          headers: { Authorization: `Bearer ${missingValueBearer}` },
          method: "DELETE",
        }
      )
    );
    expect(discardedMissingValue.status).toBe(200);
    const deletedStatusResponse = await fetch(`${mock.url}/status`, {
      body: JSON.stringify({ externalReference: missingValueReference }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(deletedStatusResponse.status).toBe(200);
    expect(await deletedStatusResponse.json()).toMatchObject({
      status: "deleted",
      submittedAt: null,
    });
    const malformedValueCreate = await fetch(`${mock.url}/handoffs`, {
      body: JSON.stringify({
        email: ` ${missingValueEmail.toUpperCase()} `,
        externalReference: `ticket-18-malformed-${crypto.randomUUID()}`,
        publicId: formRecord.publicId,
        values: {
          account: {
            address: { city: 42 },
          },
        },
      }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(malformedValueCreate.status).toBe(409);
    expect(await malformedValueCreate.json()).toMatchObject({
      error: "handoff_unavailable",
    });
    const largePrefillEmail = `ticket-18-large-${crypto.randomUUID()}@example.com`;
    const handoffCountBeforeLargeValue = await prisma.handoff.count();
    const responseCountBeforeLargeValue = await prisma.response.count({
      where: { formId },
    });
    const largePrefillCreate = await fetch(`${mock.url}/handoffs`, {
      body: JSON.stringify({
        email: largePrefillEmail,
        externalReference: `ticket-18-large-${crypto.randomUUID()}`,
        publicId: formRecord.publicId,
        values: {
          ...handoffCandidateValues,
          [selectedPointer]: "x".repeat(10_001),
        },
      }),
      headers: jsonHeaders,
      method: "POST",
    });
    expect(largePrefillCreate.status).toBe(409);
    expect(await largePrefillCreate.json()).toMatchObject({
      error: "handoff_unavailable",
    });
    expect(await prisma.handoff.count()).toBe(handoffCountBeforeLargeValue);
    expect(
      await prisma.response.count({
        where: { formId },
      })
    ).toBe(responseCountBeforeLargeValue);
    return {
      formRecord,
      handoffCandidateValues,
      handoffCreateBody: { ...handoffCreateBody, code: handoffCreateBody.code },
      missingValueBearer,
      mock,
      prefillHandoffSecret,
      user,
      userBearer,
      userId,
    };
  } catch (error) {
    mock.close();
    throw error;
  }
};
