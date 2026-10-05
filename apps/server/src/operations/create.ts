import type {
  FillMethod,
  OperationTargetType,
  OperationType,
} from "@onlyoffice/db";
import { prisma, Prisma, OperationStatus } from "@onlyoffice/db";

import { databaseErrorCode } from "../db-errors";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type {
  ActionEditorAuthorization,
  EditorCapabilityScope,
} from "../editor/authorization";
import { lockActiveEditorLease } from "../editor/leases";
import { fail } from "../http/errors";
import type { Operation } from "../model-types";
import { jsonValue } from "../responses/data";
import type { OperationMetadata } from "./model";

export async function createOperation(input: {
  actorId: string;
  authorization: ActionEditorAuthorization;
  capabilityScope: Omit<EditorCapabilityScope, "action" | "operationId">;
  documentKey: string;
  formId: string;
  expectedFillMethod?: FillMethod;
  metadata: OperationMetadata;
  ownerUserId: string;
  responseId?: string;
  stagingObjectKey: string;
  submissionId?: string;
  targetId: string;
  targetType: OperationTargetType;
  type: OperationType;
}): Promise<Operation> {
  try {
    return await prisma.$transaction(async (tx) => {
      const [lockedForm] = await tx.$queryRaw<
        {
          fillMethod: FillMethod;
          id: string;
        }[]
      >(
        Prisma.sql`
          SELECT "id", "fill_method" AS "fillMethod"
          FROM "forms"
          WHERE "id" = ${input.formId}::uuid
          FOR UPDATE
        `
      );
      if (!lockedForm) {
        fail(404, "not_found", "Form was not found");
      }
      if (
        input.expectedFillMethod !== undefined &&
        lockedForm.fillMethod !== input.expectedFillMethod
      ) {
        fail(409, "fill_method_changed", "The form Fill Method changed");
      }
      await lockActiveEditorLease(
        tx,
        input.authorization,
        input.capabilityScope
      );
      const activeOperation = await tx.operation.findFirst({
        select: { id: true },
        where: {
          status: { in: [OperationStatus.pending, OperationStatus.processing] },
          targetId: input.targetId,
          targetType: input.targetType,
        },
      });
      if (activeOperation) {
        fail(
          409,
          "operation_in_progress",
          "Another document operation is already in progress"
        );
      }
      return tx.operation.create({
        data: {
          actorId: input.actorId,
          documentKey: input.documentKey,
          errorCode: null,
          formId: input.formId,
          metadata: jsonValue(input.metadata),
          ownerUserId: input.ownerUserId,
          responseId: input.responseId,
          stagingObjectKey: input.stagingObjectKey,
          status: OperationStatus.pending,
          submissionId: input.submissionId,
          targetId: input.targetId,
          targetType: input.targetType,
          type: input.type,
        },
      });
    });
  } catch (error) {
    if (databaseErrorCode(error) === "P2002") {
      fail(
        409,
        "operation_in_progress",
        "Another document operation is already in progress"
      );
    }
    throw error;
  }
}
