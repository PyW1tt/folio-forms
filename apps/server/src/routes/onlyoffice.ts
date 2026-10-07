import { prisma, OperationStatus } from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { HttpError, fail } from "../http/errors";
import { readJsonRecord } from "../http/input";
import { corsOrigin } from "../http/origins";
import {
  verifyDocumentAccessToken,
  verifyOnlyOfficeAuthorization,
  pluginGuid,
  callbackClaim,
} from "../onlyoffice";
import { maxCallbackBodyBytes, finalizeCallback } from "../operations/callback";
import {
  updateOperationFailed,
  operationDocumentKey,
  consumeCallbackClaim,
} from "../operations/lifecycle";
import type { CallbackPayload } from "../operations/model";
import { cleanupTerminalOperationObjects } from "../operations/object-cleanup";
import {
  pluginIndexResponse,
  pluginOrigins,
  pluginScriptResponse,
} from "../plugin-assets";
import { nativePdfDocument } from "../responses/native-document";
import { DOCX_CONTENT_TYPE, objectExists, streamObject } from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerOnlyOfficeRoutes(
  app: Elysia,
  dependencies: Pick<
    RouteDependencies,
    "allowedCallbackOrigins" | "callbackMaximumBytes"
  >
): void {
  const { allowedCallbackOrigins, callbackMaximumBytes } = dependencies;
  app
    .get("/onlyoffice/document/:key", async ({ request, params, query }) => {
      const { key } = params;
      const token = typeof query.token === "string" ? query.token : "";
      if (
        !verifyDocumentAccessToken(token, key) ||
        !verifyOnlyOfficeAuthorization(request.headers.get("authorization"), {
          url: request.url,
        })
      ) {
        fail(
          401,
          "unauthorized",
          "Document access token is invalid or expired"
        );
      }
      const objectKey = await operationDocumentKey(key);
      if (!objectKey || !(await objectExists(objectKey))) {
        fail(404, "not_found", "Document was not found");
      }
      if (query.pdf === "1") {
        return new Response(await nativePdfDocument(key), {
          headers: { "Content-Type": DOCX_CONTENT_TYPE },
        });
      }
      return new Response(streamObject(objectKey), {
        headers: { "Content-Type": DOCX_CONTENT_TYPE },
      });
    })
    .get("/onlyoffice-plugin/config.json", ({ request, set }) => {
      const origin = request.headers.get("origin");
      if (origin && !pluginOrigins.has(origin)) {
        fail(403, "forbidden_origin", "Origin is not allowed");
      }
      if (origin && origin !== corsOrigin) {
        set.headers["Access-Control-Allow-Origin"] = origin;
        set.headers.Vary = "Origin";
      }
      return {
        guid: pluginGuid,
        name: "ตั้งค่าฟิลด์",
        variations: [
          {
            EditorsSupport: ["word"],
            buttons: [],
            description: "แผงตั้งค่าฟิลด์สำหรับแบบฟอร์ม",
            events: [
              "onToolbarMenuClick",
              "onDocumentContentReady",
              "onChangeContentControl",
              "onTargetPositionChanged",
            ],
            initData: "",
            initDataType: "none",
            initOnSelectionChanged: true,
            isActivated: true,
            isInsideMode: false,
            isModal: false,
            isViewer: false,
            isVisual: true,
            type: "panelRight",
            url: "index.html",
          },
        ],
        version: "2.1.0",
      };
    })
    .get("/onlyoffice-plugin", pluginIndexResponse)
    .get("/onlyoffice-plugin/", pluginIndexResponse)
    .get("/onlyoffice-plugin/index.html", pluginIndexResponse)
    .get("/onlyoffice-plugin/plugin.js", pluginScriptResponse)
    .post(
      "/onlyoffice/callback",
      async ({ request }) => {
        const payload = await readJsonRecord(request, maxCallbackBodyBytes);
        if (
          !verifyOnlyOfficeAuthorization(
            request.headers.get("authorization"),
            payload
          )
        ) {
          fail(401, "invalid_onlyoffice_token", "OnlyOffice token is invalid");
        }
        const status =
          typeof payload.status === "number"
            ? payload.status
            : Number(payload.status);
        if (status !== 6 && status !== 7) {
          return { error: 0 };
        }
        if (typeof payload.userdata !== "string") {
          return { error: 1 };
        }
        const claim = callbackClaim(payload.userdata);
        if (!claim) {
          return { error: 1 };
        }
        const operation = await prisma.operation.findUnique({
          where: { id: claim.operationId },
        });
        if (
          !operation ||
          claim.documentKey !== operation.documentKey ||
          claim.operationType !== operation.type ||
          typeof payload.key !== "string" ||
          payload.key !== operation.documentKey
        ) {
          return { error: 1 };
        }
        const consumption = await consumeCallbackClaim(
          operation.id,
          payload.userdata
        );
        if (consumption === "invalid") {
          return { error: 1 };
        }
        if (
          consumption === "replayed" ||
          operation.status === OperationStatus.completed ||
          operation.status === OperationStatus.failed
        ) {
          if (
            operation.status === OperationStatus.completed ||
            operation.status === OperationStatus.failed
          ) {
            await cleanupTerminalOperationObjects(operation);
          }
          return { error: 0 };
        }
        if (status === 7) {
          await updateOperationFailed(
            operation.id,
            "onlyoffice_document_error"
          );
          return { error: 0 };
        }
        try {
          await finalizeCallback(
            operation.id,
            payload as unknown as CallbackPayload,
            undefined,
            allowedCallbackOrigins,
            callbackMaximumBytes
          );
          return { error: 0 };
        } catch (error) {
          await updateOperationFailed(
            operation.id,
            error instanceof HttpError && error.code === "invalid_template"
              ? "invalid_template"
              : "callback_processing_failed"
          );
          return { error: 1 };
        }
      },
      { parse: "none" }
    );
}
