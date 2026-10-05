// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import {
  ACTIONS,
  API_ROUTES,
  MAX_OPERATION_POLLS,
  OPERATION_POLL_INTERVAL_MS,
} from "./constants.js";
import {
  requireOption,
  isRecord,
  firstString,
  actionLabel,
  wait,
  errorMessage,
} from "./values.js";

export function createOperations(
  options,
  office,
  prefill,
  bridge,
  api,
  callbacks
) {
  const { setStatus, isInitializationPending } = callbacks;
  let actionInFlight = false;

  function actionRequest(action, data, reason) {
    const documentKey = requireOption(options.documentKey, "documentKey");

    if (action === ACTIONS.SAVE_TEMPLATE || action === ACTIONS.PUBLISH) {
      const path = api.adminFormPath(
        action === ACTIONS.SAVE_TEMPLATE ? "save" : "publish"
      );

      return {
        body: {
          documentKey,
        },
        path,
      };
    }

    if (action === ACTIONS.SAVE_DRAFT || action === ACTIONS.SUBMIT) {
      const publicId = encodeURIComponent(
        requireOption(options.publicId, "publicId")
      );
      const body = {
        canonicalDateFields: data.canonicalDateFields,
        data: data.data,
        documentKey,
      };

      if (options.responseId) {
        body.responseId = options.responseId;
      }

      return {
        body,
        path: `${API_ROUTES.FORMS}/${publicId}/${
          action === ACTIONS.SAVE_DRAFT ? "draft" : "submit"
        }`,
      };
    }

    if (action === ACTIONS.SAVE_CORRECTION) {
      const responseId = encodeURIComponent(
        requireOption(options.responseId, "responseId")
      );
      return {
        body: {
          data: data.data,
          documentKey,
          reason: requireOption(reason, "correction reason"),
        },
        path: `${API_ROUTES.ADMIN_RESULTS}/${responseId}/correction`,
      };
    }

    throw new Error(`Unsupported form action: ${action || "none"}`);
  }

  async function postAction(action, data, reason, capability) {
    const request = actionRequest(action, data, reason);
    const result = await api.requestJson(
      request.path,
      {
        body: JSON.stringify(request.body),
        method: "POST",
      },
      capability
    );

    return result;
  }

  function operationIdFromResponse(payload) {
    if (!isRecord(payload)) {
      return;
    }

    const operation = isRecord(payload.operation) ? payload.operation : {};
    const data = isRecord(payload.data) ? payload.data : {};

    return firstString(
      payload.operationId,
      operation.operationId,
      operation.id,
      data.operationId
    );
  }

  function operationCapabilityFromResponse(payload) {
    if (!isRecord(payload)) {
      return;
    }

    return firstString(payload.operationCapability);
  }

  function operationStatus(payload) {
    const operation = isRecord(payload?.operation) ? payload.operation : {};
    const status = payload?.status ?? operation.status;

    if (typeof status !== "string") {
      return "";
    }

    return status
      .trim()
      .toLowerCase()
      .replaceAll(/[\s-]+/g, "_");
  }

  function operationFailureMessage(payload, operationId) {
    const operation = isRecord(payload?.operation) ? payload.operation : {};
    const detail = payload?.error ?? payload?.message ?? operation.error;

    return detail ? String(detail) : `การดำเนินการ ${operationId} ไม่สำเร็จ`;
  }

  async function pollOperation(operationId, label, operationCapability) {
    const operationLabel = label === undefined ? "operation" : label;
    const id = requireOption(operationId, "operationId");
    const capability = requireOption(
      operationCapability,
      "operationCapability"
    );
    let previousStatus = "";

    for (let attempt = 0; attempt < MAX_OPERATION_POLLS; attempt += 1) {
      const payload = await api.requestJson(
        `${API_ROUTES.OPERATIONS}/${encodeURIComponent(id)}`,
        { method: "GET" },
        capability
      );
      const status = operationStatus(payload);

      if (!status) {
        throw new Error(`การดำเนินการ ${id} ไม่มีสถานะ`);
      }

      if (
        status !== previousStatus &&
        (status === "pending" || status === "queued" || status === "processing")
      ) {
        previousStatus = status;
        setStatus(`${operationLabel} กำลังประมวลผล…`, "pending");
      }

      if (
        status === "completed" ||
        status === "complete" ||
        status === "succeeded" ||
        status === "success" ||
        status === "done"
      ) {
        return payload;
      }

      if (
        status === "failed" ||
        status === "failure" ||
        status === "error" ||
        status === "cancelled" ||
        status === "canceled"
      ) {
        throw new Error(operationFailureMessage(payload, id));
      }

      if (attempt < MAX_OPERATION_POLLS - 1) {
        await wait(OPERATION_POLL_INTERVAL_MS);
      }
    }

    throw new Error(`การดำเนินการ ${id} ใช้เวลานานเกินไป`);
  }

  async function runAction(action, reason = "") {
    if (actionInFlight || isInitializationPending()) {
      setStatus("มีการดำเนินการของฟอร์มกำลังทำงานอยู่", "pending");
      return { ignored: true, ok: false };
    }

    actionInFlight = true;
    let operationId;
    let completedPayload;
    let editorFrozen = false;
    const correctionReason = typeof reason === "string" ? reason.trim() : "";

    try {
      const needsData =
        action === ACTIONS.SAVE_DRAFT ||
        action === ACTIONS.SAVE_CORRECTION ||
        action === ACTIONS.SUBMIT;
      let data;

      if (needsData) {
        if (prefill.shouldApplyRuntimePrefill()) {
          await prefill.ensurePrefill();
        } else if (
          options.action === ACTIONS.DRAFT ||
          options.action === ACTIONS.CORRECTION
        ) {
          await office.restrictEditorToForms();
        }
        await office.freezeEditor();
        editorFrozen = true;
        setStatus(`กำลังอ่านข้อมูลของ ${actionLabel(action)}…`, "pending");
        data = await office.extractFormDataPromise();
      }
      if (!needsData) {
        await office.freezeEditor();
        editorFrozen = true;
      }
      await bridge.flushEditorChanges();

      setStatus(`${actionLabel(action)} กำลังรอดำเนินการ…`, "pending");
      const capability = await bridge.requestActionCapability(action);
      const response = await postAction(
        action,
        data,
        correctionReason,
        capability
      );
      operationId = operationIdFromResponse(response);
      bridge.notifyParent(action, "pending", operationId, response);

      if (operationId) {
        completedPayload = await pollOperation(
          operationId,
          actionLabel(action),
          operationCapabilityFromResponse(response)
        );
      }

      setStatus(`${actionLabel(action)} สำเร็จ`, "success");
      bridge.notifyParent(
        action,
        "completed",
        operationId,
        completedPayload || response
      );

      if (needsData) {
        bridge.setDirtyState(false);
      }

      return {
        ok: true,
        operationId: operationId || null,
        response,
      };
    } catch (error) {
      const message = errorMessage(error);
      setStatus(`${actionLabel(action)} ไม่สำเร็จ: ${message}`, "error");
      bridge.notifyParent(
        action,
        "failed",
        operationId,
        completedPayload,
        message
      );

      return {
        error: message,
        ok: false,
        operationId: operationId || null,
      };
    } finally {
      if (editorFrozen) {
        try {
          await office.executeMethodResult("SetEditingRestrictions", [
            options.action === ACTIONS.TEMPLATE_EDIT ? "none" : "forms",
          ]);
        } catch {
          setStatus("ไม่สามารถคืนค่าการแก้ไขเฉพาะช่องกรอกได้", "error");
        }
      }
      actionInFlight = false;
    }
  }

  function submitForm() {
    return runAction(ACTIONS.SUBMIT);
  }

  return { pollOperation, runAction, submitForm };
}
