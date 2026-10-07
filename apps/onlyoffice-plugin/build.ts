import path from "node:path";

import { buildPlugin } from "../server/src/plugin-bundle.ts";

const pluginDirectory = import.meta.dirname;
await Bun.write(
  path.join(pluginDirectory, "dist/plugin.js"),
  await buildPlugin(pluginDirectory)
);
