import { prisma } from "@onlyoffice/db";

// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { EditorCapabilityIdentity } from "../auth/identity";
import { requireIdentity } from "../auth/identity";
import { fail } from "../http/errors";
import type {
  EditorCapabilityClaims,
  EditorCapabilityAction,
  EditorCapabilityTarget,
} from "../onlyoffice";
import { verifyEditorCapability, createEditorCapability } from "../onlyoffice";
import type { EditorLeaseGrant } from "./leases";

interface EditorAuthorization {
  actor: EditorCapabilityIdentity;
  capability: EditorCapabilityClaims | null;
}
export type ActionEditorAuthorization = EditorAuthorization & {
  capability: EditorCapabilityClaims;
};
export interface EditorCapabilityScope {
  action: EditorCapabilityAction;
  documentKey: string;
  formId: string;
  operationId?: string;
  targetId: string;
  targetType: EditorCapabilityTarget;
}
export async function editorAuthorization(
  request: Request
): Promise<EditorAuthorization> {
  const token = request.headers.get("x-editor-capability")?.trim();
  if (!token) {
    return { actor: await requireIdentity(request), capability: null };
  }
  const capability = verifyEditorCapability(token);
  if (!capability) {
    fail(
      401,
      "invalid_editor_capability",
      "Editor capability is invalid or expired"
    );
  }
  const user = await prisma.user.findUnique({
    select: {
      email: true,
      enabled: true,
      id: true,
      mustChangePassword: true,
      name: true,
      role: true,
    },
    where: { id: capability.actorId },
  });
  const ssoUser = capability.isSso === true && user?.role === "user";
  if (
    !user?.enabled ||
    (capability.isSso === true && user.role !== "user") ||
    (user.mustChangePassword && !ssoUser) ||
    user.role !== capability.role
  ) {
    fail(
      401,
      "invalid_editor_capability",
      "Editor capability is invalid or expired"
    );
  }
  return {
    actor: {
      email: user.email,
      expiresAt: new Date(
        (capability.sessionExpiresAt ?? capability.expiresAt) * 1000
      ),
      id: user.id,
      isSso: capability.isSso === true,
      mustChangePassword: false,
      name: user.name,
      role: user.role,
    },
    capability,
  };
}
export function actionEditorCapability(
  identity: EditorCapabilityIdentity,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">,
  action: Exclude<EditorCapabilityAction, "poll-operation">,
  lease: EditorLeaseGrant
): string {
  return createEditorCapability({
    ...scope,
    action,
    actorId: identity.id,
    expiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
    isSso: identity.isSso === true,
    leaseId: lease.id,
    leaseProof: lease.proof,
    role: identity.role,
    sessionExpiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
  });
}

export function operationEditorCapability(
  identity: EditorCapabilityIdentity,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">,
  operationId: string
): string {
  return createEditorCapability({
    ...scope,
    action: "poll-operation",
    actorId: identity.id,
    expiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
    isSso: identity.isSso === true,
    operationId,
    role: identity.role,
    sessionExpiresAt: identity.isSso
      ? Math.floor(identity.expiresAt.getTime() / 1000)
      : undefined,
  });
}
export async function requireActionEditorAuthorization(
  request: Request
): Promise<ActionEditorAuthorization> {
  const authorization = await editorAuthorization(request);
  const { capability } = authorization;
  if (!capability) {
    fail(
      401,
      "editor_capability_required",
      "An editor capability is required for this action"
    );
  }
  return { actor: authorization.actor, capability };
}

export function requireEditorScope(
  authorization: EditorAuthorization,
  scope: EditorCapabilityScope
): void {
  const { capability } = authorization;
  if (!capability) {
    return;
  }
  if (
    capability.action !== scope.action ||
    capability.documentKey !== scope.documentKey ||
    capability.formId !== scope.formId ||
    capability.operationId !== scope.operationId ||
    capability.targetId !== scope.targetId ||
    capability.targetType !== scope.targetType
  ) {
    fail(
      403,
      "editor_capability_scope_mismatch",
      "Editor capability does not permit this action"
    );
  }
}
