import { prisma, OperationStatus, OperationTargetType } from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import {
  requireEditorScope,
  editorAuthorization,
} from "../editor/authorization";
import { fail } from "../http/errors";
import { validateId } from "../http/input";
import { expireOperationIfNeeded } from "../operations/lifecycle";
import { cleanupTerminalOperationObjects } from "../operations/object-cleanup";
import type { RouteDependencies } from "./dependencies";

export function registerOperationRoutes(
  app: Elysia,
  _dependencies: Pick<RouteDependencies, never>
): void {
  app.get("/api/operations/:id", async ({ request, params }) => {
    const authorization = await editorAuthorization(request);
    const { actor: identity } = authorization;
    validateId(params.id, "Operation");
    let operation = await prisma.operation.findUnique({
      where: { id: params.id },
    });
    if (!operation) {
      fail(404, "not_found", "Operation was not found");
    }
    if (authorization.capability) {
      const targetType =
        operation.targetType === OperationTargetType.template_draft
          ? "template-draft"
          : operation.targetType === OperationTargetType.response
            ? "response"
            : operation.targetType === OperationTargetType.correction
              ? "correction"
              : null;
      if (
        !targetType ||
        !operation.documentKey ||
        operation.actorId !== identity.id
      ) {
        fail(
          403,
          "editor_capability_scope_mismatch",
          "Editor capability does not permit this operation"
        );
      }
      requireEditorScope(authorization, {
        action: "poll-operation",
        documentKey: operation.documentKey,
        formId: operation.formId,
        operationId: operation.id,
        targetId: operation.targetId,
        targetType,
      });
    } else if (identity.role !== "admin") {
      if (!operation.responseId) {
        fail(403, "forbidden", "You may not access this operation");
      }
      const response = await prisma.response.findUnique({
        select: { userId: true },
        where: { id: operation.responseId },
      });
      if (response?.userId !== identity.id) {
        fail(403, "forbidden", "You may not access this operation");
      }
    }
    operation = await expireOperationIfNeeded(operation);
    await cleanupTerminalOperationObjects(operation);
    return {
      operation: {
        createdAt: operation.createdAt,
        error: operation.errorCode,
        id: operation.id,
        responseId: operation.responseId,
        result:
          operation.status === OperationStatus.completed
            ? operation.result
            : undefined,
        status: operation.status,
        submissionId: operation.submissionId,
        type: operation.type,
        updatedAt: operation.updatedAt,
      },
    };
  });
}
