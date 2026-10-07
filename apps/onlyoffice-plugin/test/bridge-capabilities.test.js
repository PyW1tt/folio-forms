// oxlint-disable prefer-await-to-callbacks node/callback-return unicorn/prefer-dom-node-remove prefer-destructuring unicorn/prefer-response-static-json no-await-in-loop unicorn/consistent-function-scoping no-plusplus unicorn/prefer-array-find -- VM harness intentionally mirrors browser callbacks and DOM shims.
import { expect, test } from "bun:test";

import {
  acknowledgeBridge,
  completedOperation,
  createHarness,
  flushPlugin,
} from "./harness.js";

test("uses fresh capabilities and an exact acknowledged bridge", async () => {
  const harness = createHarness({
    capabilityResponses: [
      "fresh-save-draft-capability",
      "fresh-submit-capability",
    ],
    responses: [
      {
        operationCapability: "save-operation-capability",
        operationId: "save-operation",
      },
      { operation: { status: "processing" } },
      completedOperation({ saved: true }),
      {
        operationCapability: "submit-operation-capability",
        operationId: "submit-operation",
      },
      { operation: { status: "processing" } },
      completedOperation({ submissionId: "submission-1" }),
    ],
  });

  expect(harness.messages).toHaveLength(1);
  expect(harness.messages[0]).toEqual({
    message: {
      bridgeId: harness.bridgeId,
      source: "form-bridge",
      type: "bridge-ready",
    },
    targetOrigin: harness.parentOrigin,
  });

  acknowledgeBridge(harness);

  const firstAction = harness.window.FormBridge.runAction("save-draft");
  await expect(harness.window.FormBridge.runAction("submit")).resolves.toEqual({
    ignored: true,
    ok: false,
  });
  await expect(firstAction).resolves.toMatchObject({ ok: true });
  await expect(
    harness.window.FormBridge.runAction("submit")
  ).resolves.toMatchObject({ ok: true });

  const capabilityRequests = harness.messages.filter(
    ({ message }) => message.type === "capability-request"
  );
  expect(capabilityRequests).toEqual([
    {
      message: {
        action: "save-draft",
        bridgeId: harness.bridgeId,
        requestId: "capability-1",
        source: "form-bridge",
        type: "capability-request",
      },
      targetOrigin: harness.parentOrigin,
    },
    {
      message: {
        action: "submit",
        bridgeId: harness.bridgeId,
        requestId: "capability-2",
        source: "form-bridge",
        type: "capability-request",
      },
      targetOrigin: harness.parentOrigin,
    },
  ]);

  expect(
    harness.requests.map(({ headers, method, url }) => [
      method,
      url,
      headers.get("x-editor-capability"),
      headers.has("authorization"),
    ])
  ).toEqual([
    [
      "POST",
      "https://api.example.test/api/forms/public-id/draft",
      "fresh-save-draft-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/save-operation",
      "save-operation-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/save-operation",
      "save-operation-capability",
      false,
    ],
    [
      "POST",
      "https://api.example.test/api/forms/public-id/submit",
      "fresh-submit-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/submit-operation",
      "submit-operation-capability",
      false,
    ],
    [
      "GET",
      "https://api.example.test/api/operations/submit-operation",
      "submit-operation-capability",
      false,
    ],
  ]);
  expect(
    harness.requests.every(({ credentials }) => credentials === "omit")
  ).toBe(true);

  const operationMessages = harness.messages.filter(
    ({ message }) => message.type === "operation"
  );
  expect(operationMessages).toHaveLength(4);
  expect(
    operationMessages.map(({ message, targetOrigin }) => [
      message.action,
      message.status,
      message.bridgeId,
      targetOrigin,
    ])
  ).toEqual([
    ["save-draft", "pending", harness.bridgeId, harness.parentOrigin],
    ["save-draft", "completed", harness.bridgeId, harness.parentOrigin],
    ["submit", "pending", harness.bridgeId, harness.parentOrigin],
    ["submit", "completed", harness.bridgeId, harness.parentOrigin],
  ]);
});

test("rejects parent acknowledgements with the wrong source, origin, or bridge id", async () => {
  const invalidEvents = [
    (harness) => ({
      data: {
        bridgeId: harness.bridgeId,
        source: "folio-parent",
        type: "bridge-ack",
      },
      origin: harness.parentOrigin,
      source: {},
    }),
    (harness) => ({
      data: {
        bridgeId: harness.bridgeId,
        source: "folio-parent",
        type: "bridge-ack",
      },
      origin: "https://wrong.example.test",
      source: harness.parentWindow,
    }),
    (harness) => ({
      data: {
        bridgeId: "wrong-bridge",
        source: "folio-parent",
        type: "bridge-ack",
      },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    }),
  ];

  await Promise.all(
    invalidEvents.map(async (createEvent) => {
      const harness = createHarness({
        action: "template-edit",
      });
      harness.dispatch(createEvent(harness));
      await expect(
        harness.window.FormBridge.runAction("save-template")
      ).resolves.toMatchObject({ ok: false });
      expect(harness.requests).toHaveLength(0);
      expect(
        harness.messages.filter(({ message }) => message.type === "operation")
      ).toHaveLength(0);
    })
  );
});

test("fails safely when capability renewal times out or returns an error", async () => {
  const timeoutHarness = createHarness({ action: "template-edit" });
  acknowledgeBridge(timeoutHarness);
  const timedOutAction =
    timeoutHarness.window.FormBridge.runAction("save-template");
  await flushPlugin();

  expect(timeoutHarness.requests).toHaveLength(0);
  timeoutHarness.expireCapabilityRequests();
  await expect(timedOutAction).resolves.toMatchObject({
    error: "Timed out waiting for save-template capability",
    ok: false,
  });

  const errorHarness = createHarness({
    action: "template-edit",
    capabilityResponses: [{ error: "capability denied" }],
  });
  acknowledgeBridge(errorHarness);

  await expect(
    errorHarness.window.FormBridge.runAction("save-template")
  ).resolves.toMatchObject({
    error: "capability denied",
    ok: false,
  });
  expect(errorHarness.requests).toHaveLength(0);
});

test("ignores forged or malformed capability responses", async () => {
  const harness = createHarness({
    action: "template-edit",
    responses: [{ ok: true }],
  });
  acknowledgeBridge(harness);

  const action = harness.window.FormBridge.runAction("save-template");
  await flushPlugin();
  const request = harness.messages.find(
    ({ message }) => message.type === "capability-request"
  );
  const validResponse = {
    ...request.message,
    capability: "forged-capability",
    source: "folio-parent",
    type: "capability-response",
  };
  const invalidEvents = [
    {
      data: validResponse,
      origin: harness.parentOrigin,
      source: {},
    },
    {
      data: validResponse,
      origin: "https://wrong.example.test",
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, bridgeId: "wrong-bridge" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, requestId: "unknown-request" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, action: "submit" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, action: "unknown-action" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, source: "attacker" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, type: "unknown-response" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, capability: "" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, capability: {} },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: {
        ...harness.messages.find(
          ({ message }) => message.type === "capability-request"
        ).message,
        source: "folio-parent",
        type: "capability-response",
      },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
    {
      data: { ...validResponse, capability: "forged", error: "also forged" },
      origin: harness.parentOrigin,
      source: harness.parentWindow,
    },
  ];

  let settled = false;
  void action.then(() => {
    settled = true;
  });
  for (const event of invalidEvents) {
    harness.dispatch(event);
  }
  await Promise.resolve();

  expect(settled).toBe(false);
  expect(harness.requests).toHaveLength(0);
  harness.expireCapabilityRequests();
  await expect(action).resolves.toMatchObject({ ok: false });
});

test("uses the public Form identifier for Admin template actions", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["action-capability"],
    responses: [{ ok: true }],
  });
  acknowledgeBridge(harness);

  await expect(
    harness.window.FormBridge.runAction("save-template")
  ).resolves.toMatchObject({ ok: true });
  expect(harness.requests[0]?.url).toBe(
    "https://api.example.test/api/admin/forms/public-id/save"
  );
  expect(harness.requests[0]?.url).not.toContain("form-id");
});
