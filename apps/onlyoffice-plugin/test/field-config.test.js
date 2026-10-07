// oxlint-disable prefer-await-to-callbacks node/callback-return unicorn/prefer-dom-node-remove prefer-destructuring unicorn/prefer-response-static-json no-await-in-loop unicorn/consistent-function-scoping no-plusplus unicorn/prefer-array-find -- VM harness intentionally mirrors browser callbacks and DOM shims.
import { expect, test } from "bun:test";

import {
  acknowledgeBridge,
  createHarness,
  flushPlugin,
  selectedControl,
} from "./harness.js";

test("configures the selected field with exact schema pointers and policy", async () => {
  let copiedPointer = null;
  const pointer = "/account/address/city";
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: [
      "field-capability-1",
      "field-capability-2",
      "field-capability-3",
      "field-capability-4",
    ],
    clipboard: {
      writeText(value) {
        copiedPointer = value;
        return Promise.resolve();
      },
    },
    responses: [
      { rules: [] },
      {
        items: [{ pointer, type: "string" }],
        nextCursor: "schema-cursor-1",
      },
      {
        items: [{ pointer: "/account/address/country", type: "string" }],
        nextCursor: null,
      },
      {
        rule: {
          prefillPointer: pointer,
          prefillPolicy: "lock-when-available",
          required: true,
          tag: pointer,
        },
      },
    ],
    selection: selectedControl("title", "text", "control-1"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.element("field-panel").hidden).toBe(false);
  expect(harness.element("field-selection-tag").textContent).toBe("title");
  expect(harness.requests[0]).toMatchObject({
    credentials: "omit",
    method: "GET",
    url: "https://api.example.test/api/admin/forms/public-id/field-rules",
  });

  const query = harness.element("field-schema-query");
  query.value = "address";
  query.dispatchEvent({ target: query, type: "input" });
  harness.element("field-schema-search").dispatchEvent({ type: "click" });
  await flushPlugin();

  expect(harness.requests[1]).toMatchObject({
    credentials: "omit",
    method: "GET",
    url: "https://api.example.test/api/admin/forms/public-id/schema?q=address",
  });
  expect(harness.requests[1].headers.get("x-editor-capability")).toBe(
    "field-capability-2"
  );
  const list = harness.element("field-schema-list");
  expect(list.children).toHaveLength(1);
  const firstRow = list.children[0];
  const firstActions = firstRow.children[2];
  firstActions.children[0].dispatchEvent({ type: "click" });
  await flushPlugin();
  expect(copiedPointer).toBe(pointer);

  firstRow.children[0].dispatchEvent({ type: "click" });
  harness.element("field-schema-next").dispatchEvent({ type: "click" });
  await flushPlugin();
  expect(harness.requests[2]).toMatchObject({
    credentials: "omit",
    method: "GET",
    url: "https://api.example.test/api/admin/forms/public-id/schema?q=address&cursor=schema-cursor-1",
  });
  expect(list.children).toHaveLength(2);

  harness.element("field-apply-pointer").dispatchEvent({ type: "click" });
  await flushPlugin();
  const required = harness.element("field-required");
  required.checked = true;
  required.dispatchEvent({ target: required, type: "change" });
  const policy = harness.element("field-prefill-policy");
  policy.value = "lock-when-available";
  policy.dispatchEvent({ target: policy, type: "change" });
  harness.element("field-policy-form").dispatchEvent({ type: "submit" });
  await flushPlugin();

  expect(harness.requests[3]).toMatchObject({
    credentials: "omit",
    method: "PATCH",
    url: "https://api.example.test/api/admin/forms/public-id/field-rules",
  });
  expect(JSON.parse(harness.requests[3].body)).toEqual({
    documentKey: "document-key",
    prefillPointer: pointer,
    prefillPolicy: "lock-when-available",
    previousTag: null,
    required: true,
    tag: pointer,
  });
});
test("keeps native picture fields embedded and outside prefill configuration", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["picture-field-capability"],
    responses: [
      {
        rules: [
          {
            prefillPointer: "/photo",
            prefillPolicy: "lock-when-available",
            required: true,
            tag: "photo",
          },
        ],
      },
    ],
    selection: selectedControl("photo", "picture", "picture-control"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.element("field-picture-help").hidden).toBe(false);
  expect(harness.element("field-prefill-policy").disabled).toBe(true);
  expect(harness.element("field-schema-query").disabled).toBe(true);
  expect(harness.window.FormBridge.getPanelState()).toMatchObject({
    currentPointer: null,
    prefillPolicy: "editable",
    required: true,
    selection: { controlType: "picture", tag: "photo" },
  });
});

test("loads persisted field rules into the selected panel", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["field-capability"],
    responses: [
      {
        rules: [
          {
            prefillPointer: "/account/address/city",
            prefillPolicy: "lock-when-available",
            required: true,
            tag: "title",
          },
        ],
      },
    ],
    selection: selectedControl("title", "text", "control-1"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.element("field-required").checked).toBe(true);
  expect(harness.element("field-prefill-policy").value).toBe(
    "lock-when-available"
  );
  expect(harness.window.FormBridge.getPanelState()).toMatchObject({
    currentPointer: "/account/address/city",
    selection: { tag: "title" },
  });
});

test("keeps field configuration available only in template-edit mode", async () => {
  const templateHarness = createHarness({ action: "template-edit" });
  acknowledgeBridge(templateHarness);
  await flushPlugin();
  expect(templateHarness.element("field-panel").hidden).toBe(false);
  expect(typeof templateHarness.window.FormBridge.saveFieldRule).toBe(
    "function"
  );
  expect(
    templateHarness.messages.find(
      ({ message }) => message.type === "field-selection"
    )?.message
  ).toMatchObject({
    controlType: "unsupported",
    selected: false,
    tag: null,
    type: "field-selection",
  });

  const fillHarness = createHarness({ action: "fill" });
  expect(fillHarness.element("field-panel").hidden).toBe(true);
  expect(fillHarness.window.FormBridge.applySchemaPointer).toBeUndefined();
  expect(fillHarness.window.FormBridge.getFieldRules).toBeUndefined();
  expect(fillHarness.window.FormBridge.saveFieldRule).toBeUndefined();
  for (const action of ["configure-fields", "publish", "save-template"]) {
    await expect(
      fillHarness.window.FormBridge.runAction(action)
    ).resolves.toMatchObject({ ok: false });
  }
  expect(fillHarness.requests).toHaveLength(0);
  expect(
    fillHarness.messages.filter(
      ({ message }) => message.type === "capability-request"
    )
  ).toHaveLength(0);
});

test("keeps field-rule API failures outside selection value payloads", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["field-capability"],
    responses: [{ body: { error: "unavailable" }, httpStatus: 503 }],
    selection: selectedControl("title", "text", "control-1"),
  });
  acknowledgeBridge(harness);
  await flushPlugin();

  expect(harness.requests[0]?.body).toBeUndefined();
  expect(
    harness.messages
      .filter(({ message }) => message.type === "field-selection")
      .every(({ message }) => !Object.hasOwn(message, "value"))
  ).toBe(true);
});

test("retries persisted field rules when bridge acknowledgement is delayed", async () => {
  const harness = createHarness({
    action: "template-edit",
    capabilityResponses: ["field-capability"],
    responses: [{ rules: [] }],
    selection: selectedControl("title", "text", "control-1"),
  });
  await flushPlugin();
  expect(harness.requests).toHaveLength(0);

  acknowledgeBridge(harness);
  await flushPlugin();
  expect(harness.requests).toHaveLength(1);
  expect(harness.requests[0]?.url).toContain("/field-rules");
});
