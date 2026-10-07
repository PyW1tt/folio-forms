import type { Prisma } from "@onlyoffice/db";
import { prisma } from "@onlyoffice/db";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Elysia } from "elysia";

import { safeAuditMetadata } from "../audit/events";
import {
  auditEventCursor,
  auditFilterValue,
  auditOutcomeValue,
  auditEventPageSize,
  auditEventCursorValue,
} from "../audit/query";
import type { AuditEventCursor } from "../audit/query";
import { requireIdentity, requireAdmin } from "../auth/identity";
import { fail } from "../http/errors";
import { queryString, idPattern } from "../http/input";
import type { JsonRecord } from "../model-types";
import { adminResultDate } from "../responses/query";
import type { RouteDependencies } from "./dependencies";

export function registerAuditRoutes(
  app: Elysia,
  _dependencies: Pick<RouteDependencies, never>
): void {
  app.get("/api/admin/audit-events", async ({ request, query }) => {
    const identity = await requireIdentity(request);
    requireAdmin(identity);
    const queryRecord = query as unknown as JsonRecord;
    const cursor = auditEventCursor(queryString(queryRecord, "cursor"));
    const actorId = auditFilterValue(
      queryString(queryRecord, "actor"),
      "actor",
      64
    );
    const action = auditFilterValue(
      queryString(queryRecord, "action"),
      "action",
      80
    );
    const targetId = auditFilterValue(
      queryString(queryRecord, "target"),
      "target",
      200
    );
    const targetType = auditFilterValue(
      queryString(queryRecord, "targetType"),
      "targetType",
      80
    );
    const outcome = auditOutcomeValue(queryString(queryRecord, "outcome"));
    const from = adminResultDate(queryString(queryRecord, "from"), "from");
    const to = adminResultDate(queryString(queryRecord, "to"), "to");
    if (from && to && from > to) {
      fail(400, "invalid_request", "from must be before to");
    }
    if (actorId && !idPattern.test(actorId)) {
      fail(400, "invalid_request", "actor is invalid");
    }
    const and: Prisma.AuditEventWhereInput[] = [];
    if (cursor) {
      and.push({
        OR: [
          { createdAt: { lt: cursor.createdAt } },
          { createdAt: cursor.createdAt, id: { lt: cursor.id } },
        ],
      });
    }
    if (actorId) {
      and.push({ actorId });
    }
    if (action) {
      and.push({ action });
    }
    if (targetId) {
      and.push({ targetId });
    }
    if (targetType) {
      and.push({ targetType });
    }
    if (outcome) {
      and.push({ outcome });
    }
    if (from || to) {
      and.push({
        createdAt: {
          ...(from ? { gte: from } : {}),
          ...(to ? { lte: to } : {}),
        },
      });
    }
    const events = await prisma.auditEvent.findMany({
      orderBy: [{ createdAt: "desc" }, { id: "desc" }],
      select: {
        action: true,
        actorId: true,
        createdAt: true,
        id: true,
        outcome: true,
        safeMetadata: true,
        targetId: true,
        targetType: true,
      },
      take: auditEventPageSize + 1,
      where: and.length > 0 ? { AND: and } : {},
    });
    const page = events.slice(0, auditEventPageSize);
    return {
      events: page.map((event) => ({
        action: event.action,
        actorId: event.actorId,
        createdAt: event.createdAt,
        id: event.id,
        outcome: event.outcome,
        safeMetadata: safeAuditMetadata(event.safeMetadata),
        targetId: event.targetId,
        targetType: event.targetType,
      })),
      nextCursor:
        events.length > auditEventPageSize
          ? auditEventCursorValue(page.at(-1) as AuditEventCursor)
          : null,
    };
  });
}
