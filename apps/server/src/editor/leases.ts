import { createHmac } from "node:crypto";

import {
  OperationTargetType,
  prisma,
  OperationStatus,
  Prisma,
} from "@onlyoffice/db";
import { env } from "@onlyoffice/env/server";

import type { Identity } from "../auth/identity";
import { tokenDigest } from "../digests";
import { fail } from "../http/errors";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { EditorCapabilityTarget } from "../onlyoffice";
import { expireOperationIfNeeded } from "../operations/lifecycle";
import { drainObjectCleanupIntents } from "../operations/object-cleanup";
import type {
  ActionEditorAuthorization,
  EditorCapabilityScope,
} from "./authorization";

const editorLeaseDurationMs = 90_000;
interface ClaimedEditorLease {
  expiresAt: Date;
  id: string;
  workspaceBaseDocumentKey?: string;
  workspaceBaseRevision?: number;
  workspaceDocumentKey?: string;
  workspaceObjectKey?: string;
}
export interface CorrectionWorkspaceInput {
  baseDocumentKey: string;
  baseRevision: number;
  documentKey: string;
  objectKey: string;
}
export interface EditorLeaseGrant extends ClaimedEditorLease {
  proof: string;
}
interface ActiveEditorLease {
  capabilityDigest: string;
  holderSessionId: string;
  holderUserId: string;
}

function editorLeaseTargetType(
  targetType: EditorCapabilityTarget
): OperationTargetType {
  if (targetType === "template-draft") {
    return OperationTargetType.template_draft;
  }
  return targetType === "correction"
    ? OperationTargetType.correction
    : OperationTargetType.response;
}

export function editorLeaseProof(
  holder: Pick<Identity, "id" | "sessionId">,
  targetType: EditorCapabilityTarget,
  targetId: string
): string {
  return createHmac("sha256", env.EDITOR_CAPABILITY_SECRET)
    .update(
      [
        "editor-lease-v1",
        targetType,
        targetId,
        holder.sessionId,
        holder.id,
      ].join("\0")
    )
    .digest("base64url");
}

function nextEditorLeaseExpiry(identity: Identity, now: Date): Date {
  const expiresAt = new Date(
    Math.min(
      now.getTime() + editorLeaseDurationMs,
      identity.expiresAt.getTime()
    )
  );
  if (expiresAt.getTime() <= now.getTime()) {
    fail(401, "unauthorized", "Authentication is required");
  }
  return expiresAt;
}

export function editorLeaseBridge(lease: ClaimedEditorLease) {
  return {
    expiresAt: lease.expiresAt.toISOString(),
    id: lease.id,
    releaseUrl: `/api/editor-leases/${lease.id}`,
    renewUrl: `/api/editor-leases/${lease.id}/renew`,
  };
}

export async function claimEditorLease(
  identity: Identity,
  targetType: EditorCapabilityTarget,
  targetId: string,
  formId: string,
  workspace?: CorrectionWorkspaceInput
): Promise<EditorLeaseGrant> {
  const now = new Date();
  const expiresAt = nextEditorLeaseExpiry(identity, now);
  const proof = editorLeaseProof(identity, targetType, targetId);
  const capabilityDigest = tokenDigest(proof);
  const databaseTargetType = editorLeaseTargetType(targetType);
  const activeOperation = await prisma.operation.findFirst({
    where: {
      status: { in: [OperationStatus.pending, OperationStatus.processing] },
      targetId,
      targetType: databaseTargetType,
    },
  });
  if (activeOperation) {
    await expireOperationIfNeeded(activeOperation);
  }
  const lease = await prisma.$transaction(async (tx) => {
    const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
      Prisma.sql`
        SELECT "id"
        FROM "forms"
        WHERE "id" = ${formId}::uuid
        FOR UPDATE
      `
    );
    if (!lockedForm) {
      fail(404, "not_found", "Form was not found");
    }
    const [current] = await tx.$queryRaw<
      (ActiveEditorLease & ClaimedEditorLease)[]
    >(
      Prisma.sql`
        SELECT
          "id",
          "capability_digest" AS "capabilityDigest",
          "expires_at" AS "expiresAt",
          "holder_session_id" AS "holderSessionId",
          "holder_user_id" AS "holderUserId",
          "workspace_base_document_key" AS "workspaceBaseDocumentKey",
          "workspace_base_revision" AS "workspaceBaseRevision",
          "workspace_document_key" AS "workspaceDocumentKey",
          "workspace_object_key" AS "workspaceObjectKey"
        FROM "editor_leases"
        WHERE
          "target_type" = CAST(${databaseTargetType} AS "OperationTargetType")
          AND "target_id" = ${targetId}::uuid
        FOR UPDATE
      `
    );
    const sameHolder =
      current?.holderSessionId === identity.sessionId &&
      current?.holderUserId === identity.id;
    const leaseId =
      sameHolder && current && current.expiresAt > now
        ? current.id
        : crypto.randomUUID();
    if (!sameHolder) {
      const activeOperation = await tx.operation.findFirst({
        select: { id: true },
        where: {
          status: {
            in: [OperationStatus.pending, OperationStatus.processing],
          },
          targetId,
          targetType: databaseTargetType,
        },
      });
      if (activeOperation) {
        return null;
      }
    }
    const [claimed] = await tx.$queryRaw<ClaimedEditorLease[]>(
      Prisma.sql`
        INSERT INTO "editor_leases" (
          "id",
          "target_type",
          "target_id",
          "holder_session_id",
          "holder_user_id",
          "capability_digest",
          "expires_at",
          "workspace_base_document_key",
          "workspace_base_revision",
          "workspace_document_key",
          "workspace_object_key",
          "renewed_at"
        )
        VALUES (
          ${leaseId}::uuid,
          CAST(${databaseTargetType} AS "OperationTargetType"),
          ${targetId}::uuid,
          ${identity.sessionId},
          ${identity.id},
          ${capabilityDigest},
          ${expiresAt},
          ${workspace?.baseDocumentKey ?? null},
          ${workspace?.baseRevision ?? null},
          ${workspace?.documentKey ?? null},
          ${workspace?.objectKey ?? null},
          ${now}
        )
        ON CONFLICT ("target_type", "target_id") DO UPDATE SET
          "id" = EXCLUDED."id",
          "holder_session_id" = EXCLUDED."holder_session_id",
          "holder_user_id" = EXCLUDED."holder_user_id",
          "capability_digest" = EXCLUDED."capability_digest",
          "expires_at" = EXCLUDED."expires_at",
          "workspace_base_document_key" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_base_document_key"
            ELSE "editor_leases"."workspace_base_document_key"
          END,
          "workspace_base_revision" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_base_revision"
            ELSE "editor_leases"."workspace_base_revision"
          END,
          "workspace_document_key" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_document_key"
            ELSE "editor_leases"."workspace_document_key"
          END,
          "workspace_object_key" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              OR "editor_leases"."holder_session_id" <> ${identity.sessionId}
              OR "editor_leases"."holder_user_id" <> ${identity.id}
              THEN EXCLUDED."workspace_object_key"
            ELSE "editor_leases"."workspace_object_key"
          END,
          "created_at" = CASE
            WHEN "editor_leases"."expires_at" <= ${now}
              THEN EXCLUDED."created_at"
            ELSE "editor_leases"."created_at"
          END,
          "renewed_at" = EXCLUDED."renewed_at"
        WHERE
          "editor_leases"."expires_at" <= ${now}
          OR (
            "editor_leases"."holder_session_id" = ${identity.sessionId}
            AND "editor_leases"."holder_user_id" = ${identity.id}
          )
        RETURNING
          "id",
          "expires_at" AS "expiresAt",
          "workspace_base_document_key" AS "workspaceBaseDocumentKey",
          "workspace_base_revision" AS "workspaceBaseRevision",
          "workspace_document_key" AS "workspaceDocumentKey",
          "workspace_object_key" AS "workspaceObjectKey"
      `
    );
    return claimed ?? null;
  });
  if (!lease) {
    fail(409, "editor_in_use", "This document is open in another session");
  }
  return { ...lease, proof };
}

export async function renewEditorLease(
  identity: Identity,
  leaseId: string
): Promise<ClaimedEditorLease> {
  const now = new Date();
  const expiresAt = nextEditorLeaseExpiry(identity, now);
  const renewed = await prisma.editorLease.updateMany({
    data: { expiresAt, renewedAt: now },
    where: {
      expiresAt: { gt: now },
      holderSessionId: identity.sessionId,
      holderUserId: identity.id,
      id: leaseId,
    },
  });
  if (renewed.count !== 1) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  return { expiresAt, id: leaseId };
}

export async function releaseEditorLease(
  identity: Identity,
  leaseId: string
): Promise<void> {
  const workspaceObjectKey = await prisma.$transaction(async (tx) => {
    const [lease] = await tx.$queryRaw<
      {
        targetId: string;
        targetType: OperationTargetType;
        workspaceObjectKey: string | null;
      }[]
    >(
      Prisma.sql`
        SELECT
          "target_id" AS "targetId",
          "target_type" AS "targetType",
          "workspace_object_key" AS "workspaceObjectKey"
        FROM "editor_leases"
        WHERE
          "id" = ${leaseId}::uuid
          AND "holder_session_id" = ${identity.sessionId}
          AND "holder_user_id" = ${identity.id}
        FOR UPDATE
      `
    );
    if (!lease) {
      fail(
        409,
        "editor_lease_inactive",
        "The editor lease is no longer active"
      );
    }
    const activeOperation = await tx.operation.findFirst({
      select: { id: true },
      where: {
        status: {
          in: [OperationStatus.pending, OperationStatus.processing],
        },
        targetId: lease.targetId,
        targetType: lease.targetType,
      },
    });
    if (activeOperation) {
      return null;
    }
    const released = await tx.editorLease.deleteMany({
      where: { id: leaseId },
    });
    if (released.count !== 1) {
      fail(
        409,
        "editor_lease_inactive",
        "The editor lease is no longer active"
      );
    }
    return lease.workspaceObjectKey;
  });
  if (workspaceObjectKey) {
    await drainObjectCleanupIntents([workspaceObjectKey]);
  }
}

function editorLeaseClaims(authorization: ActionEditorAuthorization): {
  id: string;
  proof: string;
} {
  const { leaseId, leaseProof } = authorization.capability;
  if (!leaseId || !leaseProof) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  return { id: leaseId, proof: leaseProof };
}

function requireEditorLeaseProof(
  scope: Omit<EditorCapabilityScope, "action" | "operationId">,
  lease: ActiveEditorLease,
  proof: string
): void {
  const expectedProof = editorLeaseProof(
    { id: lease.holderUserId, sessionId: lease.holderSessionId },
    scope.targetType,
    scope.targetId
  );
  if (
    proof !== expectedProof ||
    tokenDigest(proof) !== lease.capabilityDigest
  ) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
}

export async function requireActiveEditorLease(
  authorization: ActionEditorAuthorization,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">
): Promise<void> {
  const claim = editorLeaseClaims(authorization);
  const now = new Date();
  const lease = await prisma.editorLease.findFirst({
    select: {
      capabilityDigest: true,
      holderSessionId: true,
      holderUserId: true,
    },
    where: {
      expiresAt: { gt: now },
      holderSession: { expiresAt: { gt: now } },
      holderUserId: authorization.actor.id,
      id: claim.id,
      targetId: scope.targetId,
      targetType: editorLeaseTargetType(scope.targetType),
    },
  });
  if (!lease) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  requireEditorLeaseProof(scope, lease, claim.proof);
}

export async function lockActiveEditorLease(
  tx: Prisma.TransactionClient,
  authorization: ActionEditorAuthorization,
  scope: Omit<EditorCapabilityScope, "action" | "operationId">
): Promise<void> {
  const claim = editorLeaseClaims(authorization);
  const databaseTargetType = editorLeaseTargetType(scope.targetType);
  const [lease] = await tx.$queryRaw<ActiveEditorLease[]>(
    Prisma.sql`
      SELECT
        "editor_leases"."capability_digest" AS "capabilityDigest",
        "editor_leases"."holder_session_id" AS "holderSessionId",
        "editor_leases"."holder_user_id" AS "holderUserId"
      FROM "editor_leases"
      INNER JOIN "session"
        ON "session"."id" = "editor_leases"."holder_session_id"
      WHERE
        "editor_leases"."id" = ${claim.id}::uuid
        AND "editor_leases"."target_type" =
          CAST(${databaseTargetType} AS "OperationTargetType")
        AND "editor_leases"."target_id" = ${scope.targetId}::uuid
        AND "editor_leases"."holder_user_id" = ${authorization.actor.id}
        AND "editor_leases"."expires_at" > NOW()
        AND "session"."expires_at" > NOW()
      FOR UPDATE OF "editor_leases"
    `
  );
  if (!lease) {
    fail(409, "editor_lease_inactive", "The editor lease is no longer active");
  }
  requireEditorLeaseProof(scope, lease, claim.proof);
}
