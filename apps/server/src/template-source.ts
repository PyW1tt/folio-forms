// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import path from "node:path";

import { env } from "@onlyoffice/env/server";

import { validateTemplatePackage } from "./documents/package";
import { fail } from "./http/errors";
import { maxTemplateUploadBytes } from "./http/input";

const fallbackTemplatePath = path.resolve(
  import.meta.dirname,
  "../../../onlyoffice-templates/template.docx"
);
export async function findTemplateSource(): Promise<string | null> {
  const configuredSource = Bun.file(env.TEMPLATE_PATH);
  if (await configuredSource.exists()) {
    return env.TEMPLATE_PATH;
  }
  const fallbackSource = Bun.file(fallbackTemplatePath);
  return (await fallbackSource.exists()) ? fallbackTemplatePath : null;
}

export async function readTemplateSourceBytes(
  absolutePath: string
): Promise<Uint8Array> {
  const file = Bun.file(absolutePath);
  if (!(await file.exists())) {
    fail(404, "not_found", "Template source was not found");
  }
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes.byteLength > maxTemplateUploadBytes) {
    fail(413, "payload_too_large", "Template source is too large");
  }
  validateTemplatePackage(bytes);
  return bytes;
}
