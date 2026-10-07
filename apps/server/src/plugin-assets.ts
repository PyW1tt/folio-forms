import path from "node:path";

import { env } from "@onlyoffice/env/server";

import { originOf } from "./http/origins";
import { htmlEscape } from "./markup";
import { buildPlugin } from "./plugin-bundle";

const onlyOfficePluginSdkUrlPlaceholder = "__ONLYOFFICE_PLUGIN_SDK_URL__";

const pluginDir = path.resolve(import.meta.dirname, "../../onlyoffice-plugin");

export const pluginIndexResponse = async () => {
  const html = await Bun.file(path.resolve(pluginDir, "index.html")).text();
  const onlyOfficeBaseUrl = env.ONLYOFFICE_URL.replace(/\/+$/u, "");
  const sdkUrl = `${onlyOfficeBaseUrl}/sdkjs-plugins/v1/plugins.js`;
  return new Response(
    html.replace(onlyOfficePluginSdkUrlPlaceholder, htmlEscape(sdkUrl)),
    { headers: { "content-type": "text/html; charset=utf-8" } }
  );
};

export const pluginOrigins = new Set(
  [env.API_BASE, env.ONLYOFFICE_URL, env.CORS_ORIGIN]
    .map(originOf)
    .filter((origin): origin is string => Boolean(origin))
);

export const pluginScriptResponse = async () =>
  new Response(
    env.NODE_ENV === "production"
      ? Bun.file(path.resolve(pluginDir, "dist/plugin.js"))
      : await buildPlugin(pluginDir),
    { headers: { "content-type": "text/javascript; charset=utf-8" } }
  );
