// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.

export function databaseErrorCode(error: unknown): string | null {
  if (!error || typeof error !== "object" || !("code" in error)) {
    return null;
  }
  const code = error.code;
  return typeof code === "string" ? code : null;
}
export function isSerializationConflict(error: unknown): boolean {
  if (databaseErrorCode(error) === "P2034") {
    return true;
  }
  if (databaseErrorCode(error) !== "P2010") {
    return false;
  }
  if (
    !error ||
    typeof error !== "object" ||
    !("meta" in error) ||
    !error.meta ||
    typeof error.meta !== "object" ||
    !("code" in error.meta)
  ) {
    return false;
  }
  return error.meta.code === "40001";
}
