import {
  prisma,
  Prisma,
  ResponseStatus,
  OperationStatus,
  FillMethod,
  OperationTargetType,
} from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { databaseErrorCode } from "../db-errors";
import { validateResponseDocument } from "../documents/native-eligibility";
import { callbackFieldMetadata } from "../documents/office-values";
import {
  requireActionEditorAuthorization,
  requireEditorScope,
  operationEditorCapability,
} from "../editor/authorization";
import {
  lockActiveEditorLease,
  requireActiveEditorLease,
} from "../editor/leases";
import {
  findFormByPublicId,
  nativePictureManifestFields,
} from "../forms/query";
import { fail } from "../http/errors";
import { requiredString } from "../http/input";
import type { Operation } from "../model-types";
import { createOperation } from "../operations/create";
import { launchForceSave } from "../operations/force-save";
import {
  activeOperationForResponse,
  updateOperationFailed,
} from "../operations/lifecycle";
import type { OperationMetadata } from "../operations/model";
import { operationTypeForAction } from "../operations/model";
import { processNativeResponseOperation } from "../operations/response-completion";
import {
  normalizeResponseData,
  jsonValue,
  requireCurrentFillMethod,
} from "../responses/data";
import { nativeResponseDocument } from "../responses/native-document";
import { nativeResponseRequest } from "../responses/native-request";
import { findOwnedResponse } from "../responses/query";
import { DOCX_CONTENT_TYPE, readObject, objectKey } from "../storage";
import type { RouteDependencies } from "./dependencies";

export function registerResponseActionRoutes(
  app: Elysia,
  dependencies: Pick<
    RouteDependencies,
    "onlyOffice" | "storeObject" | "allowedCallbackOrigins"
  >
): void {
  const { onlyOffice, storeObject, allowedCallbackOrigins } = dependencies;
  app
    .post(
      "/api/forms/:publicId/draft",
      async ({ request, params, set }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        const form = await findFormByPublicId(params.publicId);
        const pictureFields =
          form.fillMethod === FillMethod.native && form.publishedTemplate
            ? await nativePictureManifestFields(form.publishedTemplate.id)
            : [];
        const nativeRequest = await nativeResponseRequest(
          request,
          pictureFields
        );
        const input = nativeRequest.input;
        requireCurrentFillMethod(form, input);
        const responseId = requiredString(input, "responseId");
        const documentKey = requiredString(input, "documentKey");
        const response = await findOwnedResponse(
          responseId,
          form.id,
          identity.id
        );
        if (
          response.status !== ResponseStatus.draft ||
          response.draftDocumentKey !== documentKey ||
          response.publishedVersion !== form.version
        ) {
          fail(409, "stale_response", "The response is no longer editable");
        }
        const capabilityScope = {
          documentKey,
          formId: form.id,
          targetId: response.id,
          targetType: "response",
        } as const;
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "save-draft",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const data = await normalizeResponseData(
          form,
          response,
          input.data,
          false,
          form.fillMethod === FillMethod.onlyoffice,
          form.fillMethod === FillMethod.onlyoffice
            ? input.canonicalDateFields
            : undefined
        );
        if (await activeOperationForResponse(response.id)) {
          fail(
            409,
            "operation_in_progress",
            "Another response operation is already in progress"
          );
        }
        const nativeDocumentBytes =
          pictureFields.length === 0
            ? null
            : response.draftObjectKey
              ? await nativeResponseDocument(
                  response.publishedTemplateId,
                  data,
                  await readObject(response.draftObjectKey),
                  nativeRequest.pictures
                )
              : fail(
                  409,
                  "document_unavailable",
                  "Response document is unavailable"
                );
        if (nativeDocumentBytes) {
          await validateResponseDocument(
            response.publishedTemplateId,
            nativeDocumentBytes,
            false
          );
        }
        const operationId = crypto.randomUUID();
        const stagedObjectKey = objectKey(
          "operations",
          operationId,
          "draft",
          crypto.randomUUID(),
          "docx"
        );
        const metadata: OperationMetadata = {
          action: "save-draft",
          data,
          finalObjectKey: objectKey(
            "responses",
            response.id,
            "draft",
            crypto.randomUUID(),
            "docx"
          ),
          formId: form.id,
          nextDocumentKey: `response-${response.id}-${crypto.randomUUID()}`,
          publicId: form.publicId,
          responseId: response.id,
          ...(form.fillMethod === FillMethod.onlyoffice
            ? callbackFieldMetadata(data, input.data, response.prefillSnapshot)
            : {}),
          ...(nativeDocumentBytes ? { nativeDocumentStaged: true } : {}),
          stagedObjectKey,
        };

        const operation = await createOperation({
          actorId: identity.id,
          authorization,
          capabilityScope,
          documentKey,
          expectedFillMethod: form.fillMethod,
          formId: form.id,
          metadata,
          ownerUserId: identity.id,
          responseId: response.id,
          stagingObjectKey: stagedObjectKey,
          targetId: response.id,
          targetType: OperationTargetType.response,
          type: operationTypeForAction["save-draft"],
        });
        if (nativeDocumentBytes) {
          try {
            await storeObject(
              stagedObjectKey,
              nativeDocumentBytes,
              DOCX_CONTENT_TYPE
            );
          } catch (error) {
            await updateOperationFailed(operation.id, "document_save_failed");
            throw error;
          }
        }
        set.status = 202;
        if (form.fillMethod === FillMethod.native) {
          void processNativeResponseOperation(operation, storeObject);
        } else {
          launchForceSave(operation, onlyOffice, allowedCallbackOrigins);
        }
        return {
          operationCapability: operationEditorCapability(
            identity,
            capabilityScope,
            operation.id
          ),
          operationId: operation.id,
          responseId: response.id,
          status: operation.status,
        };
      },
      { parse: "none" }
    )
    .post(
      "/api/forms/:publicId/submit",
      async ({ request, params, set }) => {
        const authorization = await requireActionEditorAuthorization(request);
        const { actor: identity } = authorization;
        const form = await findFormByPublicId(params.publicId);
        const pictureFields =
          form.fillMethod === FillMethod.native && form.publishedTemplate
            ? await nativePictureManifestFields(form.publishedTemplate.id)
            : [];
        const nativeRequest = await nativeResponseRequest(
          request,
          pictureFields
        );
        const input = nativeRequest.input;
        requireCurrentFillMethod(form, input);
        const responseId = requiredString(input, "responseId");
        const documentKey = requiredString(input, "documentKey");
        const response = await findOwnedResponse(
          responseId,
          form.id,
          identity.id
        );
        if (
          response.status !== ResponseStatus.draft ||
          response.draftDocumentKey !== documentKey ||
          response.publishedVersion !== form.version
        ) {
          fail(409, "stale_response", "The response is no longer editable");
        }
        const capabilityScope = {
          documentKey,
          formId: form.id,
          targetId: response.id,
          targetType: "response",
        } as const;
        requireEditorScope(authorization, {
          ...capabilityScope,
          action: "submit",
        });
        await requireActiveEditorLease(authorization, capabilityScope);
        const data = await normalizeResponseData(
          form,
          response,
          input.data,
          true,
          form.fillMethod === FillMethod.onlyoffice,
          form.fillMethod === FillMethod.onlyoffice
            ? input.canonicalDateFields
            : undefined
        );
        if (await activeOperationForResponse(response.id)) {
          fail(
            409,
            "operation_in_progress",
            "Another response operation is already in progress"
          );
        }
        const nativeDocumentBytes =
          pictureFields.length === 0
            ? null
            : response.draftObjectKey
              ? await nativeResponseDocument(
                  response.publishedTemplateId,
                  data,
                  await readObject(response.draftObjectKey),
                  nativeRequest.pictures
                )
              : fail(
                  409,
                  "document_unavailable",
                  "Response document is unavailable"
                );
        if (nativeDocumentBytes) {
          await validateResponseDocument(
            response.publishedTemplateId,
            nativeDocumentBytes,
            true
          );
        }
        const operationId = crypto.randomUUID();
        const submissionId = crypto.randomUUID();
        const submissionDocumentKey = `submission-${submissionId}-${crypto.randomUUID()}`;
        const stagedObjectKey = objectKey(
          "operations",
          operationId,
          "submission",
          crypto.randomUUID(),
          "docx"
        );
        const metadata: OperationMetadata = {
          action: "submit",
          data,
          finalObjectKey: objectKey(
            "submissions",
            submissionId,
            "filled",
            crypto.randomUUID(),
            "docx"
          ),
          formId: form.id,
          publicId: form.publicId,
          responseId: response.id,
          ...(form.fillMethod === FillMethod.onlyoffice
            ? callbackFieldMetadata(data, input.data, response.prefillSnapshot)
            : {}),
          stagedObjectKey,
          submissionDocumentKey,
          submissionId,
          ...(nativeDocumentBytes ? { nativeDocumentStaged: true } : {}),
        };
        let operation: Operation;
        try {
          operation = await prisma.$transaction(
            async (tx) => {
              const [lockedForm] = await tx.$queryRaw<
                { fillMethod: FillMethod }[]
              >(
                Prisma.sql`
                  SELECT "fill_method" AS "fillMethod"
                  FROM "forms"
                  WHERE "id" = ${form.id}::uuid
                  FOR UPDATE
                `
              );
              if (!lockedForm || lockedForm.fillMethod !== form.fillMethod) {
                fail(
                  409,
                  "fill_method_changed",
                  "The form Fill Method changed"
                );
              }
              await lockActiveEditorLease(tx, authorization, capabilityScope);
              const activeOperation = await tx.operation.findFirst({
                where: {
                  responseId: response.id,
                  status: {
                    in: [OperationStatus.pending, OperationStatus.processing],
                  },
                  targetId: response.id,
                  targetType: OperationTargetType.response,
                },
              });
              if (activeOperation) {
                fail(
                  409,
                  "operation_in_progress",
                  "Another response operation is already in progress"
                );
              }
              const claimed = await tx.response.updateMany({
                data: {
                  status: ResponseStatus.submitting,
                  updatedAt: new Date(),
                },
                where: { id: response.id, status: ResponseStatus.draft },
              });
              if (claimed.count !== 1) {
                fail(
                  409,
                  "operation_in_progress",
                  "Another response operation is already in progress"
                );
              }
              return tx.operation.create({
                data: {
                  actorId: identity.id,
                  documentKey,
                  errorCode: null,
                  formId: form.id,
                  metadata: jsonValue(metadata),
                  ownerUserId: identity.id,
                  responseId: response.id,
                  stagingObjectKey: stagedObjectKey,
                  status: OperationStatus.pending,
                  submissionId: null,
                  targetId: response.id,
                  targetType: OperationTargetType.response,
                  type: operationTypeForAction.submit,
                },
              });
            },
            { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
          );
        } catch (error) {
          if (databaseErrorCode(error) === "P2034") {
            fail(
              409,
              "operation_in_progress",
              "Another response operation is already in progress"
            );
          }
          throw error;
        }
        if (nativeDocumentBytes) {
          try {
            await storeObject(
              stagedObjectKey,
              nativeDocumentBytes,
              DOCX_CONTENT_TYPE
            );
          } catch (error) {
            await updateOperationFailed(operation.id, "document_save_failed");
            throw error;
          }
        }
        set.status = 202;
        if (form.fillMethod === FillMethod.native) {
          void processNativeResponseOperation(operation, storeObject);
        } else {
          launchForceSave(operation, onlyOffice, allowedCallbackOrigins);
        }
        return {
          operationCapability: operationEditorCapability(
            identity,
            capabilityScope,
            operation.id
          ),
          operationId: operation.id,
          responseId: response.id,
          status: operation.status,
          submissionId,
        };
      },
      { parse: "none" }
    );
}
