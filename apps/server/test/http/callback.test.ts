import { test, expect } from "bun:test";

import { createApp } from "../../src/app";
import { resolveCallbackDocumentUrl } from "../../src/operations/callback";

// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  throw new Error(
    "The HTTP application test requires DATABASE_URL for an isolated PostgreSQL database"
  );
}
const convertedDocumentKeys: string[] = [];
createApp({
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

test("rewrites public ONLYOFFICE callback paths to the internal base", () => {
  expect(
    resolveCallbackDocumentUrl(
      "http://localhost:8080/office/cache/files/data/example/output.docx?md5=x",
      new Set(["http://localhost:8080"]),
      "http://localhost:8080/office",
      "http://localhost:8081"
    )
  ).toBe("http://localhost:8081/cache/files/data/example/output.docx?md5=x");
});
