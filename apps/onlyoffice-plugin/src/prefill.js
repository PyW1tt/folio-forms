// oxlint-disable func-style sort-keys no-implicit-globals no-unused-vars consistent-function-scoping complexity prefer-named-capture-group require-unicode-regexp avoid-new prefer-await-to-callbacks no-empty-function no-useless-return logical-assignment-operators no-useless-spread no-await-in-loop prefer-await-to-then prefer-dom-node-remove prefer-dom-node-append no-plusplus prefer-spread -- Plugin runs inside the constrained ONLYOFFICE host runtime.
import { ACTIONS, PREFILL_MAX_ATTEMPTS } from "./constants.js";
import { applyPrefillCommand } from "./document-commands.js";
import { isRecord, parseCommandResult, wait } from "./values.js";

function copyPolicies(target, source) {
  if (Array.isArray(source)) {
    for (const field of source) {
      if (typeof field === "string" && field) {
        target[field] = "locked";
      }
    }

    return;
  }

  if (!isRecord(source)) {
    return;
  }

  for (const [field, policy] of Object.entries(source)) {
    target[field] = policy;
  }
}
function policiesForEditableFields(source) {
  if (Array.isArray(source)) {
    return Object.fromEntries(
      source
        .filter((field) => typeof field === "string" && field)
        .map((field) => [field, "editable"])
    );
  }

  if (!isRecord(source)) {
    return;
  }

  return Object.fromEntries(
    Object.entries(source).map(([field, editable]) => [
      field,
      editable === false ? "locked" : "editable",
    ])
  );
}

function normalizePrefill(options) {
  const rawPrefill = isRecord(options.prefill) ? options.prefill : undefined;
  const fieldValues = {};
  const policies = {};
  let valuesCandidate =
    (isRecord(options.prefillValues) && options.prefillValues) ||
    (isRecord(options.prefillData) && options.prefillData) ||
    undefined;

  if (rawPrefill) {
    if (isRecord(rawPrefill.values)) {
      valuesCandidate = rawPrefill.values;
    } else if (isRecord(rawPrefill.data)) {
      valuesCandidate = rawPrefill.data;
    } else if (isRecord(rawPrefill.fields)) {
      valuesCandidate = rawPrefill.fields;
    } else if (!valuesCandidate) {
      valuesCandidate = rawPrefill;
    }

    copyPolicies(policies, rawPrefill.policies);
    copyPolicies(policies, rawPrefill.fieldPolicies);
    copyPolicies(policies, rawPrefill.lockedFields);
    copyPolicies(policies, rawPrefill.locked);
    copyPolicies(
      policies,
      policiesForEditableFields(rawPrefill.editableFields)
    );
    copyPolicies(policies, policiesForEditableFields(rawPrefill.editable));
  }

  if (isRecord(valuesCandidate)) {
    const hasFieldDescriptors =
      (valuesCandidate === rawPrefill ||
        rawPrefill?.fields === valuesCandidate) &&
      !rawPrefill?.values &&
      !rawPrefill?.data;

    if (hasFieldDescriptors) {
      for (const [field, entry] of Object.entries(valuesCandidate)) {
        if (isRecord(entry) && Object.hasOwn(entry, "value")) {
          fieldValues[field] = entry.value;

          if (entry.policy !== undefined) {
            policies[field] = entry.policy;
          } else if (
            entry.locked !== undefined ||
            entry.editable !== undefined
          ) {
            policies[field] = entry;
          }
        } else {
          fieldValues[field] = entry;
        }
      }
    } else {
      Object.assign(fieldValues, valuesCandidate);
    }
  }

  copyPolicies(policies, options.policies);
  copyPolicies(policies, options.fieldPolicies);
  copyPolicies(policies, options.prefillPolicies);
  copyPolicies(policies, options.lockedFields);
  copyPolicies(policies, options.locked);

  copyPolicies(policies, policiesForEditableFields(options.editableFields));

  let defaultPolicy =
    options.prefillPolicy ?? rawPrefill?.policy ?? rawPrefill?.defaultPolicy;

  if (options.prefillLocked === true) {
    defaultPolicy = "locked";
  }

  if (typeof defaultPolicy === "string" || typeof defaultPolicy === "boolean") {
    for (const field of Object.keys(fieldValues)) {
      if (!Object.hasOwn(policies, field)) {
        policies[field] = defaultPolicy;
      }
    }
  } else if (isRecord(defaultPolicy)) {
    copyPolicies(policies, defaultPolicy.fields);
  }

  return {
    policies,
    values: fieldValues,
  };
}

function hasPrefillValues(prefill) {
  return Object.keys(prefill.values).length > 0;
}

export function createPrefill(options, office) {
  const { callCommandResult, restrictEditorToForms } = office;
  let prefillApplied = false;
  let prefillPromise;

  function applyPrefill(prefill) {
    const normalizedPrefill = normalizePrefill({
      prefill: isRecord(prefill) ? prefill : {},
    });
    const scope = window.Asc.scope || (window.Asc.scope = {});
    scope.formBridgePrefill = normalizedPrefill;
    scope.formBridgeTagAliases = options.tagAliases;

    return callCommandResult(applyPrefillCommand)
      .then((result) => parseCommandResult(result))
      .finally(() => {
        if (scope.formBridgePrefill === normalizedPrefill) {
          delete scope.formBridgePrefill;
        }
      });
  }

  async function applyPrefillWhenReady(prefill) {
    for (let attempt = 0; attempt < PREFILL_MAX_ATTEMPTS; attempt += 1) {
      const result = await applyPrefill(prefill);
      if (
        result.applied?.length ||
        result.failed?.length ||
        result.skipped?.length
      ) {
        return result;
      }
      await wait(500);
    }
    return applyPrefill(prefill);
  }
  function ensurePrefill() {
    if (prefillApplied) {
      return Promise.resolve({ applied: [], failed: [], skipped: [] });
    }

    if (!prefillPromise) {
      prefillPromise = applyPrefillWhenReady(options.prefill).then(
        async (result) => {
          if (result.failed?.length) {
            throw new Error(
              `Could not prefill ${result.failed.length} field(s): ${result.failed.join(", ")}`
            );
          }
          prefillApplied = result.applied?.length > 0;
          await restrictEditorToForms();
          return result;
        }
      );
    }

    return prefillPromise;
  }

  function shouldApplyRuntimePrefill() {
    return (
      (options.action === ACTIONS.FILL ||
        options.action === ACTIONS.CORRECTION) &&
      hasPrefillValues(options.prefill)
    );
  }

  return { applyPrefill, ensurePrefill, shouldApplyRuntimePrefill };
}

export {
  normalizePrefill,
  hasPrefillValues,
  copyPolicies,
  policiesForEditableFields,
};
