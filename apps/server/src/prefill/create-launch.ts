import { randomBytes } from "node:crypto";

import { prisma, Prisma, AuditOutcome, HandoffStatus } from "@onlyoffice/db";

import { createFormAudit, createFormFailureAudit } from "../audit/events";
import { tokenDigest } from "../digests";
import { fail } from "../http/errors";
import { publicIdPattern } from "../http/input";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { JsonRecord } from "../model-types";
import {
  validatePrefillValuesAgainstManifest,
  jsonValue,
} from "../responses/data";
import {
  handoffUnavailable,
  handoffCodeLifetimeMs,
  pendingClaimLifetimeSeconds,
} from "./input";
import { publishedPrefillConfiguration, filteredPrefillValues } from "./values";

export interface PrefillHandoffCreateInput {
  email: string;
  externalReference: string;
  publicId: string;
  values: JsonRecord;
}

interface PrefillHandoffLaunch {
  claimToken: string;
  publicId: string;
}

export async function createPrefillHandoff(
  input: PrefillHandoffCreateInput
): Promise<{ code: string; launchPath: string }> {
  if (!publicIdPattern.test(input.publicId)) {
    fail(404, "not_found", "Form was not found");
  }
  const form = await prisma.form.findUnique({
    include: {
      publishedTemplate: {
        include: {
          manifest: { include: { fields: true } },
          prefillConfiguration: { include: { fields: true } },
        },
      },
    },
    where: { publicId: input.publicId },
  });
  if (!form) {
    fail(404, "not_found", "Form was not found");
  }
  const configuration = publishedPrefillConfiguration(form);
  if (!configuration) {
    fail(404, "not_found", "Form was not found");
  }
  const filteredValues = filteredPrefillValues(
    input.values,
    configuration.fields
  );
  const manifest = form.publishedTemplate?.manifest;
  if (
    !manifest ||
    manifest.configurationHash !== configuration.configurationHash
  ) {
    handoffUnavailable();
  }
  validatePrefillValuesAgainstManifest(filteredValues, manifest.fields);
  const code = randomBytes(32).toString("base64url");
  const codeDigest = tokenDigest(code);
  const externalReferenceDigest = tokenDigest(input.externalReference);
  const expiresAt = new Date(Date.now() + handoffCodeLifetimeMs);

  try {
    await prisma.$transaction(
      async (tx) => {
        const [lockedForm] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "forms"
            WHERE "id" = ${form.id}::uuid
            FOR UPDATE
          `
        );
        if (!lockedForm) {
          fail(404, "not_found", "Form was not found");
        }
        const current = await tx.form.findUnique({
          include: {
            publishedTemplate: {
              include: {
                prefillConfiguration: { include: { fields: true } },
              },
            },
          },
          where: { id: lockedForm.id },
        });
        const currentConfiguration = publishedPrefillConfiguration(current);
        if (!current || !currentConfiguration) {
          fail(404, "not_found", "Form was not found");
        }
        if (
          currentConfiguration.configurationHash !==
          configuration.configurationHash
        ) {
          fail(404, "not_found", "Form was not found");
        }
        const deletedReference = await tx.deletionTombstone.findUnique({
          select: { id: true },
          where: { externalReferenceDigest },
        });
        if (deletedReference) {
          handoffUnavailable();
        }
        const existingReference = await tx.handoff.findUnique({
          select: { id: true },
          where: { externalReferenceDigest },
        });
        if (existingReference) {
          handoffUnavailable();
        }
        await tx.handoff.create({
          data: {
            codeDigest,
            configurationHash: currentConfiguration.configurationHash,
            expiresAt,
            externalReferenceDigest,
            filteredValues: jsonValue(filteredValues),
            form: { connect: { id: current.id } },
            id: crypto.randomUUID(),
            normalizedEmail: input.email,
          },
        });
        await createFormAudit(tx, {
          action: "create_handoff",
          actorId: null,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: current.publicId,
        });
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
  } catch (error) {
    try {
      await createFormFailureAudit({
        action: "create_handoff",
        actorId: null,
        error,
        targetId: publicIdPattern.test(input.publicId) ? input.publicId : null,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw error;
  }
  return { code, launchPath: "/prefill/handoff" };
}
export async function launchPrefillHandoff(
  code: string,
  clock: () => Date = () => new Date()
): Promise<PrefillHandoffLaunch> {
  const codeDigest = tokenDigest(code);
  let auditTargetId: string | null = null;
  try {
    const launch = await prisma.$transaction(
      async (tx) => {
        const [lockedHandoff] = await tx.$queryRaw<{ id: string }[]>(
          Prisma.sql`
            SELECT "id"
            FROM "handoffs"
            WHERE "code_digest" = ${codeDigest}
            FOR UPDATE
          `
        );
        if (!lockedHandoff) {
          handoffUnavailable();
        }
        const handoff = await tx.handoff.findUnique({
          where: { id: lockedHandoff.id },
        });
        const form = handoff?.formId
          ? await tx.form.findUnique({
              include: {
                publishedTemplate: {
                  include: {
                    prefillConfiguration: { include: { fields: true } },
                  },
                },
              },
              where: { id: handoff.formId },
            })
          : null;
        auditTargetId =
          form && publicIdPattern.test(form.publicId) ? form.publicId : null;
        const configuration = publishedPrefillConfiguration(form);
        const now = clock();
        if (
          !handoff ||
          !form ||
          !configuration ||
          handoff.status !== HandoffStatus.pending ||
          handoff.expiresAt <= now
        ) {
          handoffUnavailable();
        }
        const claimToken = randomBytes(32).toString("base64url");
        const claimExpiresAt = new Date(
          now.getTime() + pendingClaimLifetimeSeconds * 1000
        );
        await tx.pendingClaim.create({
          data: {
            claimDigest: tokenDigest(claimToken),
            expiresAt: claimExpiresAt,
            handoff: { connect: { id: handoff.id } },
            id: crypto.randomUUID(),
          },
        });
        const reserved = await tx.handoff.updateMany({
          data: {
            expiresAt: claimExpiresAt,
            reservedAt: now,
            status: HandoffStatus.reserved,
          },
          where: {
            expiresAt: { gt: now },
            id: handoff.id,
            status: HandoffStatus.pending,
          },
        });
        if (reserved.count !== 1) {
          handoffUnavailable();
        }
        await createFormAudit(tx, {
          action: "launch_handoff",
          actorId: null,
          outcome: AuditOutcome.success,
          safeMetadata: {},
          targetId: form.publicId,
        });
        return { claimToken, publicId: form.publicId };
      },
      { isolationLevel: Prisma.TransactionIsolationLevel.Serializable }
    );
    return launch;
  } catch (error) {
    try {
      await createFormFailureAudit({
        action: "launch_handoff",
        actorId: null,
        error,
        targetId: auditTargetId,
      });
    } catch {
      // Preserve the route error if the failure audit cannot be persisted.
    }
    throw error;
  }
}
