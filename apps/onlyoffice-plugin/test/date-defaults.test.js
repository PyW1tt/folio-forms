// oxlint-disable prefer-await-to-callbacks node/callback-return -- Simulate ONLYOFFICE editor event callbacks.
import { expect, test } from "bun:test";

import { createHarness, flushPlugin } from "./harness.js";

const dateControl = (id, format) => ({
  GetInternalId() {
    return id;
  },
  IsDatePicker() {
    return true;
  },
  SetDateFormat(value) {
    this.format = value;
    return true;
  },
  format,
});

test("new template date fields use day/month/year without changing existing formats", async () => {
  const existing = dateControl("existing", "yyyy-MM-dd");
  const controls = [existing];
  const harness = createHarness({
    action: "template-edit",
    controls,
    serializeCommands: true,
  });
  for (const callback of harness.editorEvents.get("onDocumentContentReady")) {
    callback();
  }
  await flushPlugin();
  const inserted = dateControl("new", "mm/dd/yyyy");
  controls.push(inserted);
  for (const callback of harness.editorEvents.get("onChangeContentControl")) {
    callback();
  }
  await flushPlugin();
  expect(inserted.format).toBe("dd/MM/yyyy");
  expect(existing.format).toBe("yyyy-MM-dd");

  inserted.format = "yyyy-MM-dd";
  for (const callback of harness.editorEvents.get("onDocumentContentChanged")) {
    callback();
  }
  await flushPlugin();
  expect(inserted.format).toBe("yyyy-MM-dd");
});

test("response editors preserve date formats", async () => {
  const controls = [];
  const harness = createHarness({ controls, serializeCommands: true });
  for (const callback of harness.editorEvents.get("onDocumentContentReady")) {
    callback();
  }
  await flushPlugin();
  const inserted = dateControl("response", "mm/dd/yyyy");
  controls.push(inserted);
  for (const callback of harness.editorEvents.get("onChangeContentControl")) {
    callback();
  }
  await flushPlugin();
  expect(inserted.format).toBe("mm/dd/yyyy");
});
