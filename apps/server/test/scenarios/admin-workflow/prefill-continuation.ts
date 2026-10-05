import { expect } from "bun:test";
import { createHash } from "node:crypto";

import { prisma } from "@onlyoffice/db";

import { createApp } from "../../../src/app";
import { reconcileRecoverableState } from "../../../src/operations/recovery";
import {
  jsonHeaders,
  createCredentialFixture,
  bearerFor,
} from "../../fixtures/http";
import type { EditorAccessOutput } from "./editor-access";
import type { PrefillEntryOutput } from "./prefill-entry";
import type { BootstrapAndCreationOutput } from "./setup";

export interface PrefillContinuationInput {
  app: ReturnType<typeof createApp>;
  password: BootstrapAndCreationOutput["password"];
  mock: PrefillEntryOutput["mock"];
  formRecord: PrefillEntryOutput["formRecord"];
  handoffCandidateValues: PrefillEntryOutput["handoffCandidateValues"];
  formId: BootstrapAndCreationOutput["formId"];
  userEmail: BootstrapAndCreationOutput["userEmail"];
  missingValueBearer: PrefillEntryOutput["missingValueBearer"];
  secondFormPublicId: EditorAccessOutput["secondFormPublicId"];
  userBearer: PrefillEntryOutput["userBearer"];
  prefillHandoffSecret: PrefillEntryOutput["prefillHandoffSecret"];
  handoffCreateBody: PrefillEntryOutput["handoffCreateBody"];
}

export const runPrefillContinuation = async (
  input: PrefillContinuationInput
): Promise<void> => {
  const {
    app,
    password,
    mock,
    formRecord,
    handoffCandidateValues,
    formId,
    userEmail,
    missingValueBearer,
    secondFormPublicId,
    userBearer,
    prefillHandoffSecret,
    handoffCreateBody,
  } = input;
  const mandatoryEmail = `ticket-17-password-${crypto.randomUUID()}@example.com`;
  const mandatoryUser = await createCredentialFixture({
    email: mandatoryEmail,
    mustChangePassword: true,
    name: "Ticket 17 Password User",
    password,
  });
  const mandatoryHandoffCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: mandatoryEmail,
      externalReference: `ticket-17-password-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const mandatoryHandoffCode = (
    (await mandatoryHandoffCreate.json()) as { code: string }
  ).code;
  const mandatoryHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: mandatoryHandoffCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(mandatoryHandoffLaunch.status).toBe(303);
  const mandatoryCookie = mandatoryHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!mandatoryCookie) {
    throw new Error("The mandatory-password handoff did not set a cookie");
  }
  const mandatoryBearer = await bearerFor(app, mandatoryEmail, password);
  const passwordChangeResponse = await app.handle(
    new Request("http://test.local/api/account/password", {
      body: JSON.stringify({
        currentPassword: password,
        newPassword: "Ticket17-new-password",
      }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${mandatoryBearer}`,
        Cookie: mandatoryCookie,
      },
      method: "POST",
    })
  );
  expect(passwordChangeResponse.status).toBe(200);
  const freshMandatoryBearer = await bearerFor(
    app,
    mandatoryEmail,
    "Ticket17-new-password"
  );
  const passwordContinuation = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${freshMandatoryBearer}`,
        Cookie: mandatoryCookie,
      },
      method: "POST",
    })
  );
  expect(passwordContinuation.status).toBe(200);
  const passwordContinuationBody = await passwordContinuation.json();
  expect(passwordContinuationBody).toMatchObject({
    response: { id: expect.any(String) },
  });
  const passwordResponseRecord = await prisma.response.findFirstOrThrow({
    select: { id: true, userId: true },
    where: { formId, userId: mandatoryUser.id },
  });
  expect(passwordResponseRecord.userId).toBe(mandatoryUser.id);
  const emailMismatchCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: `ticket-17-email-mismatch-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const emailMismatchCode = (
    (await emailMismatchCreate.json()) as {
      code: string;
    }
  ).code;
  const emailMismatchLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: emailMismatchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const emailMismatchCookie = emailMismatchLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!emailMismatchCookie) {
    throw new Error("The email-mismatch handoff did not set a cookie");
  }
  const emailMismatchStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${missingValueBearer}`,
        Cookie: emailMismatchCookie,
      },
      method: "POST",
    })
  );
  expect(emailMismatchStart.status).toBe(409);
  expect(await emailMismatchStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });

  const formMismatchCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: `ticket-17-form-mismatch-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const formMismatchCode = (
    (await formMismatchCreate.json()) as {
      code: string;
    }
  ).code;
  const formMismatchLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: formMismatchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const formMismatchCookie = formMismatchLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!formMismatchCookie) {
    throw new Error("The Form-mismatch handoff did not set a cookie");
  }
  const formMismatchStart = await app.handle(
    new Request(`http://test.local/api/forms/${secondFormPublicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: formMismatchCookie,
      },
      method: "POST",
    })
  );
  expect(formMismatchStart.status).toBe(409);
  expect(await formMismatchStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });

  const configMismatchCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: mandatoryEmail,
      externalReference: `ticket-17-config-mismatch-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const configMismatchCode = (
    (await configMismatchCreate.json()) as {
      code: string;
    }
  ).code;
  const configMismatchLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: configMismatchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const configMismatchCookie = configMismatchLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!configMismatchCookie) {
    throw new Error("The config-mismatch handoff did not set a cookie");
  }
  const configMismatchHandoff = await prisma.handoff.findFirstOrThrow({
    where: {
      codeDigest: createHash("sha256").update(configMismatchCode).digest("hex"),
    },
  });
  await prisma.handoff.update({
    data: { configurationHash: "0".repeat(64) },
    where: { id: configMismatchHandoff.id },
  });
  const configMismatchStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${freshMandatoryBearer}`,
        Cookie: configMismatchCookie,
      },
      method: "POST",
    })
  );
  expect(configMismatchStart.status).toBe(409);
  expect(await configMismatchStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const discardedMandatory = await app.handle(
    new Request(
      `http://test.local/api/responses/${passwordResponseRecord.id}`,
      {
        headers: { Authorization: `Bearer ${freshMandatoryBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(discardedMandatory.status).toBe(200);

  const expiredExternalReference = `ticket-17-expired-${crypto.randomUUID()}`;
  const expiredCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: expiredExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const expiredCode = ((await expiredCreate.json()) as { code: string }).code;
  const expiredRecord = await prisma.handoff.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      codeDigest: createHash("sha256").update(expiredCode).digest("hex"),
    },
  });
  await prisma.handoff.update({
    data: {
      expiresAt: new Date(expiredRecord.createdAt.getTime() + 1),
    },
    where: { id: expiredRecord.id },
  });
  const expiredApp = createApp({
    clock: () => new Date(expiredRecord.createdAt.getTime() + 2),
    prefillHandoffSecret,
  });
  const expiredLaunch = await expiredApp.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: expiredCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(expiredLaunch.status).toBe(303);
  expect(expiredLaunch.headers.get("location")).toBe(
    "/handoff?error=handoff_unavailable"
  );
  await reconcileRecoverableState();
  const purgedExpiredRecord = await prisma.handoff.findUniqueOrThrow({
    select: {
      configurationHash: true,
      consumedAt: true,
      createdAt: true,
      expiresAt: true,
      externalReferenceDigest: true,
      filteredValues: true,
      formId: true,
      id: true,
      normalizedEmail: true,
      reservedAt: true,
      responseId: true,
      status: true,
      updatedAt: true,
    },
    where: { id: expiredRecord.id },
  });
  expect(purgedExpiredRecord).toEqual({
    configurationHash: null,
    consumedAt: null,
    createdAt: expiredRecord.createdAt,
    expiresAt: new Date(expiredRecord.createdAt.getTime() + 1),
    externalReferenceDigest: createHash("sha256")
      .update(expiredExternalReference)
      .digest("hex"),
    filteredValues: null,
    formId: null,
    id: expiredRecord.id,
    normalizedEmail: null,
    reservedAt: null,
    responseId: null,
    status: "expired",
    updatedAt: expect.any(Date),
  });
  expect(
    await prisma.pendingClaim.count({ where: { handoffId: expiredRecord.id } })
  ).toBe(0);
  const expiredStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: expiredExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(expiredStatusResponse.status).toBe(200);
  expect(await expiredStatusResponse.json()).toMatchObject({
    status: "expired",
    submittedAt: null,
  });
  const raceExternalReference = `ticket-19-expiry-race-${crypto.randomUUID()}`;
  const raceCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: raceExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(raceCreate.status).toBe(200);
  const raceCode = ((await raceCreate.json()) as { code: string }).code;
  const raceLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: raceCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const raceCookie = raceLaunch.headers.get("set-cookie")?.split(";", 1)[0];
  if (!raceCookie) {
    throw new Error("The expiry-race handoff did not set a cookie");
  }
  const raceHandoff = await prisma.handoff.findFirstOrThrow({
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(raceExternalReference)
        .digest("hex"),
    },
  });
  await prisma.handoff.update({
    data: { expiresAt: new Date(Date.now() - 1) },
    where: { id: raceHandoff.id },
  });
  const [raceStatus, , raceRedeem] = await Promise.all([
    fetch(`${mock.url}/status`, {
      body: JSON.stringify({ externalReference: raceExternalReference }),
      headers: jsonHeaders,
      method: "POST",
    }),
    reconcileRecoverableState(),
    app.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
        headers: {
          Authorization: `Bearer ${userBearer}`,
          Cookie: raceCookie,
        },
        method: "POST",
      })
    ),
  ]);
  expect(raceStatus.status).toBe(200);
  expect(await raceStatus.json()).toMatchObject({ status: "expired" });
  expect(raceRedeem.status).toBe(409);
  const raceRedeemError = ((await raceRedeem.json()) as { error?: unknown })
    .error;
  if (typeof raceRedeemError !== "string") {
    throw new TypeError("The expiry-race redeem did not return an error code");
  }
  expect(["handoff_unavailable", "prefill_required"]).toContain(
    raceRedeemError
  );
  expect(
    await prisma.pendingClaim.count({ where: { handoffId: raceHandoff.id } })
  ).toBe(0);
  expect(
    await prisma.handoff.findUniqueOrThrow({ where: { id: raceHandoff.id } })
  ).toMatchObject({
    codeDigest: null,
    filteredValues: null,
    normalizedEmail: null,
    status: "expired",
  });
  const malformedLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: JSON.stringify({ code: handoffCreateBody.code }),
      headers: {
        "Content-Type": "application/json",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(malformedLaunch.status).toBe(303);
  expect(malformedLaunch.headers.get("location")).toBe(
    "/handoff?error=handoff_unavailable"
  );
  const nonNavigationLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: handoffCreateBody.code }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Mode": "cors",
      },
      method: "POST",
    })
  );
  expect(nonNavigationLaunch.status).toBe(303);
  expect(nonNavigationLaunch.headers.get("location")).toBe(
    "/handoff?error=handoff_unavailable"
  );
  const raceUserEmail = `ticket-17-race-${crypto.randomUUID()}@example.com`;
  await createCredentialFixture({
    email: raceUserEmail,
    name: "Ticket 17 Race User",
    password,
  });
  const raceUserBearer = await bearerFor(app, raceUserEmail, password);
  const raceHandoffReference = `ticket-17-race-${crypto.randomUUID()}`;
  const raceHandoffCreate = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: raceUserEmail,
        externalReference: raceHandoffReference,
        publicId: formRecord.publicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(raceHandoffCreate.status).toBe(200);
  const raceHandoffCode = ((await raceHandoffCreate.json()) as { code: string })
    .code;
  const raceHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: raceHandoffCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    })
  );
  expect(raceHandoffLaunch.status).toBe(303);
  const racePendingCookie = raceHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!racePendingCookie) {
    throw new Error("The race handoff did not set a pending claim cookie");
  }
  const raceStarts = await Promise.all(
    [0, 1].map(() =>
      app.handle(
        new Request(
          `http://test.local/api/forms/${formRecord.publicId}/start`,
          {
            headers: {
              Authorization: `Bearer ${raceUserBearer}`,
              Cookie: racePendingCookie,
            },
            method: "POST",
          }
        )
      )
    )
  );
  expect(raceStarts.map((response) => response.status).toSorted()).toEqual([
    200, 409,
  ]);
  const handoffRaceBodies = (await Promise.all(
    raceStarts.map((response) => response.json())
  )) as { response?: { id?: string } }[];
  const raceResponseId = handoffRaceBodies.find((body) => body.response?.id)
    ?.response?.id;
  if (!raceResponseId) {
    throw new Error("The handoff race did not create a Response");
  }
  const raceHandoffRecord = await prisma.handoff.findFirstOrThrow({
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(raceHandoffReference)
        .digest("hex"),
    },
  });
  expect(raceHandoffRecord.status).toBe("consumed");
  expect(raceHandoffRecord.responseId).toBe(raceResponseId);
  const raceDiscard = await app.handle(
    new Request(`http://test.local/api/responses/${raceResponseId}`, {
      headers: { Authorization: `Bearer ${raceUserBearer}` },
      method: "DELETE",
    })
  );
  expect(raceDiscard.status).toBe(200);
  const deletedRaceHandoff = await prisma.handoff.findUniqueOrThrow({
    where: { id: raceHandoffRecord.id },
  });
  expect(deletedRaceHandoff.status).toBe("deleted");
};
