import path from "node:path";

import type { BuildArtifact } from "bun";

export const buildPlugin = async (
  pluginDirectory: string
): Promise<BuildArtifact> => {
  const result = await Bun.build({
    entrypoints: [path.join(pluginDirectory, "src/plugin-entry.js")],
    format: "iife",
    minify: false,
    target: "browser",
    throw: false,
  });
  if (!result.success) {
    throw new AggregateError(result.logs, "ONLYOFFICE plugin build failed");
  }
  const [artifact] = result.outputs;
  if (
    result.outputs.length !== 1 ||
    !artifact ||
    artifact.kind !== "entry-point" ||
    !artifact.path.endsWith(".js")
  ) {
    throw new Error(
      "ONLYOFFICE plugin build must produce one JavaScript entry"
    );
  }
  return artifact;
};
