// oxlint-disable prefer-await-to-callbacks node/callback-return unicorn/prefer-dom-node-remove prefer-destructuring unicorn/prefer-response-static-json no-await-in-loop unicorn/consistent-function-scoping no-plusplus unicorn/prefer-array-find -- VM harness intentionally mirrors browser callbacks and DOM shims.
import { expect, test } from "bun:test";

import { acknowledgeBridge, createHarness, flushPlugin } from "./harness.js";

test("maps real ONLYOFFICE form-type snapshots to supported controls", async () => {
  const cases = [
    ["textForm", "text"],
    ["comboBoxForm", "combo"],
    ["dropDownForm", "dropdown"],
    ["checkBoxForm", "checkbox"],
    ["pictureForm", "picture"],
    ["dateForm", "date"],
    ["radioForm", "unsupported"],
  ];

  for (const [formType, expectedType] of cases) {
    const harness = createHarness({
      action: "template-edit",
      capabilityResponses: ["field-capability"],
      responses: [{ rules: [] }],
      selection: {
        currentControl: "control-1",
        formType,
        properties: {
          Id: "control-1",
          InternalId: "control-1",
          Tag: "title",
        },
      },
    });
    acknowledgeBridge(harness);
    await flushPlugin();

    const selectionMessage = harness.messages
      .filter(({ message }) => message.type === "field-selection")
      .at(-1)?.message;
    expect(selectionMessage).toMatchObject({
      controlType: expectedType,
      selected: true,
      tag: "title",
      type: "field-selection",
    });
  }
});
