import { cors } from "@elysiajs/cors";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { env } from "@onlyoffice/env/server";
import { Elysia } from "elysia";

import { AiAuthoringSessions } from "./ai-authoring";
import { AiAuthoringError } from "./ai-authoring-error";
import type { OmniRouteConfig } from "./ai-authoring/agent";
import { databaseErrorCode } from "./db-errors";
import { HttpError } from "./http/errors";
import type { LegacySsoConfig } from "./legacy-sso/config";
import {
  legacySsoConfigurationFromEnv,
  validatedLegacySsoConfiguration,
} from "./legacy-sso/config";
import type { OnlyOfficeClient } from "./onlyoffice";
import { createOnlyOfficeClient } from "./onlyoffice";
import {
  callbackOrigins,
  maxCallbackDocumentBytes,
} from "./operations/callback";
import { configuredPrefillReturnUrl } from "./prefill/input";
import { readinessStatus } from "./readiness";
import { registerAccountRoutes } from "./routes/accounts";
import { registerAiAuthoringRoutes } from "./routes/ai-authoring";
import { registerAuditRoutes } from "./routes/audit";
import {
  registerSessionRoutes,
  registerPasswordRoutes,
  registerSessionExitRoutes,
} from "./routes/auth";
import type { RouteDependencies } from "./routes/dependencies";
import { registerEditorLeaseRoutes } from "./routes/editor-leases";
import { registerFormAuthoringRoutes } from "./routes/form-authoring";
import { registerFormCatalogRoutes } from "./routes/form-catalog";
import { registerLegacySsoRoutes } from "./routes/legacy-sso";
import { registerOnlyOfficeRoutes } from "./routes/onlyoffice";
import { registerOperationRoutes } from "./routes/operations";
import {
  registerPrefillEntryRoutes,
  registerPrefillMethodRoute,
} from "./routes/prefill";
import { registerPublicFormRoutes } from "./routes/public-forms";
import { registerResponseActionRoutes } from "./routes/response-actions";
import { registerResponseRoutes } from "./routes/responses";
import { registerResultRoutes } from "./routes/results";
import { registerSubmissionRoutes } from "./routes/submissions";
import { putObject, deleteObject } from "./storage";

export interface AppOptions {
  clock?: () => Date;
  deleteObject?: (key: string) => Promise<void>;
  putObject?: typeof putObject;
  legacySso?: LegacySsoConfig | null;
  omniRoute?: OmniRouteConfig | null;
  onlyOffice?: OnlyOfficeClient;
  onlyOfficeCallbackOrigins?: readonly string[];
  onlyOfficeCallbackMaxBytes?: number;
  prefillHandoffSecret?: string;
  prefillReturnUrl?: string;
  requestIp?: (request: Request) => string | null | undefined;
}

export function createApp(options: AppOptions = {}) {
  const onlyOffice = options.onlyOffice ?? createOnlyOfficeClient();
  const removeObject = options.deleteObject ?? deleteObject;
  const storeObject = options.putObject ?? putObject;
  const allowedCallbackOrigins = options.onlyOfficeCallbackOrigins
    ? new Set(options.onlyOfficeCallbackOrigins)
    : callbackOrigins;
  const callbackMaximumBytes =
    options.onlyOfficeCallbackMaxBytes ?? maxCallbackDocumentBytes;
  const prefillHandoffSecret =
    options.prefillHandoffSecret ?? env.PREFILL_HANDOFF_SECRET;
  const prefillReturnUrl = configuredPrefillReturnUrl(
    options.prefillReturnUrl ?? env.PREFILL_RETURN_URL
  );
  const handoffClock = options.clock ?? (() => new Date());
  const legacySso =
    options.legacySso === undefined
      ? legacySsoConfigurationFromEnv()
      : options.legacySso
        ? validatedLegacySsoConfiguration(options.legacySso)
        : null;
  let omniRoute = options.omniRoute;
  if (omniRoute === undefined) {
    const apiKey = env.OMNIROUTE_API_KEY;
    const baseUrl = env.OMNIROUTE_BASE_URL;
    const model = env.OMNIROUTE_MODEL;
    omniRoute = apiKey && baseUrl && model ? { apiKey, baseUrl, model } : null;
  }
  const aiAuthoring = new AiAuthoringSessions(omniRoute, options.clock);
  const dependencies: RouteDependencies = {
    aiAuthoring,
    allowedCallbackOrigins,
    callbackMaximumBytes,
    handoffClock,
    legacySso,
    onlyOffice,
    prefillHandoffSecret,
    prefillReturnUrl,
    removeObject,
    requestIp: options.requestIp,
    storeObject,
  };
  const app = new Elysia();
  app
    .onError(({ error, set }) => {
      if (error instanceof AiAuthoringError) {
        set.status = error.status;
        return Response.json({
          error: error.code,
          message: error.message,
        });
      }
      if (error instanceof HttpError) {
        set.status = error.httpStatus;
        return Response.json({
          error: error.code,
          message: error.message,
        });
      }
      if (databaseErrorCode(error) === "P2002") {
        set.status = 409;
        return Response.json({
          error: "operation_conflict",
          message: "Another operation is already in progress",
        });
      }
      console.error(error);
      set.status = 500;
      return Response.json({
        error: "internal_error",
        message: "An unexpected server error occurred",
      });
    })
    .use(
      cors({
        allowedHeaders: [
          "Authorization",
          "Content-Type",
          "X-Editor-Capability",
          "X-Prefill-Handoff-Secret",
        ],
        credentials: true,
        methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
        origin: env.CORS_ORIGIN,
      })
    );

  registerPrefillEntryRoutes(app, dependencies);
  registerLegacySsoRoutes(app, dependencies);
  registerPrefillMethodRoute(app, dependencies);
  registerSessionRoutes(app, dependencies);
  registerEditorLeaseRoutes(app, dependencies);
  registerPasswordRoutes(app, dependencies);
  registerAccountRoutes(app, dependencies);
  registerSessionExitRoutes(app, dependencies);
  registerAiAuthoringRoutes(app, dependencies);
  app
    .get("/health", () => ({ ok: true }))
    .get("/ready", async ({ set }) => {
      if (!(await readinessStatus())) {
        set.status = 503;
        return { ok: false };
      }
      return { ok: true };
    });
  registerFormCatalogRoutes(app, dependencies);
  registerFormAuthoringRoutes(app, dependencies);
  registerAuditRoutes(app, dependencies);
  registerResultRoutes(app, dependencies);
  registerPublicFormRoutes(app, dependencies);
  registerResponseRoutes(app, dependencies);
  registerResponseActionRoutes(app, dependencies);
  registerOperationRoutes(app, dependencies);
  registerSubmissionRoutes(app, dependencies);
  registerOnlyOfficeRoutes(app, dependencies);
  app.onStop(() => aiAuthoring.close());
  return app;
}
