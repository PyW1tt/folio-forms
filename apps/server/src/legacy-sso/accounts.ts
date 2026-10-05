import { Prisma, LegacyAccountLinkStatus, prisma } from "@onlyoffice/db";

import {
  accountNameMaximumLength,
  accountMutationLockId,
} from "../accounts/mutations";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import type { Identity } from "../auth/identity";
import { normalizeEmail } from "../auth/identity";
import { databaseErrorCode } from "../db-errors";
import { accountEmailMaximumLength, accountEmailPattern } from "../http/input";
import { hasAsciiControlCharacters } from "./config";
import type { LegacySsoConfig } from "./config";

const legacySsoExchangeBodyMaximumBytes = 8 * 1024;
const legacyIdentityNameControlPattern = /\p{Cc}/u;
interface LegacyIdentity {
  email: string;
  name?: string;
  subject: string;
}
type LegacySsoAccountUser = Pick<Identity, "id" | "role"> & {
  enabled: boolean;
};
type LegacySsoAccountResolution =
  | { kind: "failed" }
  | { kind: "linked"; user: LegacySsoAccountUser }
  | { kind: "pending" };

async function readLegacyIdentityResponse(
  response: globalThis.Response
): Promise<LegacyIdentity | null> {
  if (
    !/^application\/json(?:\s*;|$)/iu.test(
      response.headers.get("content-type") ?? ""
    )
  ) {
    return null;
  }
  const contentLength = Number(response.headers.get("content-length") ?? "");
  if (
    Number.isFinite(contentLength) &&
    contentLength > legacySsoExchangeBodyMaximumBytes
  ) {
    return null;
  }
  const reader = response.body?.getReader();
  if (!reader) {
    return null;
  }
  const chunks: Uint8Array[] = [];
  let byteLength = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    byteLength += value.byteLength;
    if (byteLength > legacySsoExchangeBodyMaximumBytes) {
      await reader.cancel();
      return null;
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(byteLength);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  let payload: unknown;
  try {
    payload = JSON.parse(new TextDecoder().decode(bytes));
  } catch {
    return null;
  }
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    return null;
  }
  const claims = payload as Record<string, unknown>;
  if (
    typeof claims.sub !== "string" ||
    claims.sub.length === 0 ||
    claims.sub.length > 255 ||
    claims.sub.trim() !== claims.sub ||
    hasAsciiControlCharacters(claims.sub) ||
    typeof claims.email !== "string" ||
    claims.email.length > accountEmailMaximumLength ||
    !accountEmailPattern.test(claims.email) ||
    typeof claims.email_verified !== "boolean"
  ) {
    return null;
  }
  let name: string | undefined;
  if (
    typeof claims.name === "string" &&
    !legacyIdentityNameControlPattern.test(claims.name)
  ) {
    const trimmedName = claims.name.trim();
    if (
      trimmedName.length > 0 &&
      trimmedName.length <= accountNameMaximumLength
    ) {
      name = trimmedName;
    }
  }
  return { email: normalizeEmail(claims.email), name, subject: claims.sub };
}

export async function exchangeLegacyCode(
  config: LegacySsoConfig,
  code: string,
  codeVerifier: string
): Promise<LegacyIdentity | null> {
  try {
    const response = await fetch(config.exchangeUrl, {
      body: new URLSearchParams({
        client_id: config.clientId,
        code,
        code_verifier: codeVerifier,
        grant_type: "authorization_code",
        redirect_uri: config.callbackUrl,
      }),
      headers: {
        Accept: "application/json",
        Authorization: `Basic ${Buffer.from(
          `${config.clientId}:${config.clientSecret}`
        ).toString("base64")}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      method: "POST",
      redirect: "error",
      signal: AbortSignal.timeout(5000),
    });
    if (!response.ok) {
      return null;
    }
    return await readLegacyIdentityResponse(response);
  } catch {
    return null;
  }
}
export async function nextLegacySsoGeneration(
  tx: Prisma.TransactionClient,
  providerId: string
): Promise<bigint> {
  await tx.$executeRaw`
    SELECT pg_advisory_xact_lock(
      hashtext('legacy-sso-generation'),
      hashtext(${providerId})
    )
  `;
  const [row] = await tx.$queryRaw<{ generation: bigint }[]>(
    Prisma.sql`SELECT nextval('"legacy_sso_generation_seq"') AS generation`
  );
  if (!row) {
    throw new Error("Database did not return SSO generation");
  }
  return row.generation;
}

export async function invalidateLegacyAccountLinkReview(
  tx: Prisma.TransactionClient,
  requestId: string,
  providerId: string,
  email: string,
  userId: string
): Promise<void> {
  const reviewedGeneration = await nextLegacySsoGeneration(tx, providerId);
  await tx.legacyAccountLinkRequest.update({
    data: {
      email,
      reviewedAt: null,
      reviewedById: null,
      reviewedGeneration,
      status: LegacyAccountLinkStatus.pending,
      userId,
    },
    where: { id: requestId },
  });
}

async function linkApprovedLegacyAccount(
  tx: Prisma.TransactionClient,
  config: LegacySsoConfig,
  identity: LegacyIdentity,
  requestId: string,
  userId: string,
  transactionGeneration: bigint
): Promise<LegacySsoAccountUser | null> {
  const [lockedUser] = await tx.$queryRaw<{ id: string }[]>(
    Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${userId} FOR UPDATE`
  );
  if (!lockedUser) {
    return null;
  }
  const [lockedRequest] = await tx.$queryRaw<{ id: string }[]>(
    Prisma.sql`SELECT "id" FROM "legacy_account_link_requests" WHERE "id" = ${requestId}::uuid FOR UPDATE`
  );
  if (!lockedRequest) {
    return null;
  }
  const linkRequest = await tx.legacyAccountLinkRequest.findUnique({
    where: { id: requestId },
  });
  if (
    !linkRequest ||
    linkRequest.status !== LegacyAccountLinkStatus.approved ||
    linkRequest.reviewedGeneration === null ||
    transactionGeneration <= linkRequest.reviewedGeneration ||
    linkRequest.providerId !== config.providerId ||
    linkRequest.subject !== identity.subject ||
    linkRequest.email !== identity.email ||
    linkRequest.userId !== userId
  ) {
    return null;
  }
  const user = await tx.user.findUnique({
    select: { email: true, enabled: true, id: true, role: true },
    where: { id: userId },
  });
  if (
    !user ||
    !user.enabled ||
    user.role !== "user" ||
    user.email !== linkRequest.email
  ) {
    return null;
  }
  const account = await tx.account.findUnique({
    select: { userId: true },
    where: {
      providerId_accountId: {
        accountId: identity.subject,
        providerId: config.providerId,
      },
    },
  });
  if (account) {
    return null;
  }
  await tx.account.create({
    data: {
      accountId: identity.subject,
      id: crypto.randomUUID(),
      issuer: config.providerId,
      providerId: config.providerId,
      userId: user.id,
    },
  });
  await tx.legacyAccountLinkRequest.update({
    data: { status: LegacyAccountLinkStatus.linked },
    where: { id: requestId },
  });
  return user;
}

async function resolveLegacySsoLinkRequest(
  tx: Prisma.TransactionClient,
  config: LegacySsoConfig,
  identity: LegacyIdentity,
  transactionGeneration: bigint
): Promise<LegacySsoAccountResolution | null> {
  const account = await tx.account.findUnique({
    select: { userId: true },
    where: {
      providerId_accountId: {
        accountId: identity.subject,
        providerId: config.providerId,
      },
    },
  });
  const linkRequest = await tx.legacyAccountLinkRequest.findUnique({
    where: {
      providerId_subject: {
        providerId: config.providerId,
        subject: identity.subject,
      },
    },
  });
  if (
    linkRequest &&
    linkRequest.reviewedGeneration !== null &&
    transactionGeneration <= linkRequest.reviewedGeneration
  ) {
    return { kind: "failed" };
  }
  if (account) {
    await tx.$queryRaw(
      Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${account.userId} FOR UPDATE`
    );
    const user = await tx.user.findUnique({
      select: { enabled: true, id: true, role: true },
      where: { id: account.userId },
    });
    return user?.enabled && user.role === "user"
      ? { kind: "linked", user }
      : { kind: "failed" };
  }
  if (!linkRequest) {
    return null;
  }
  if (
    linkRequest.status !== LegacyAccountLinkStatus.pending &&
    linkRequest.status !== LegacyAccountLinkStatus.approved
  ) {
    return { kind: "failed" };
  }
  const matchingSnapshot = await tx.user.findUnique({
    select: { id: true },
    where: { email: identity.email },
  });
  const userIds = [linkRequest.userId];
  if (matchingSnapshot && matchingSnapshot.id !== linkRequest.userId) {
    userIds.push(matchingSnapshot.id);
  }
  userIds.sort();
  for (const userId of userIds) {
    await tx.$queryRaw(
      Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${userId} FOR UPDATE`
    );
  }
  const [lockedRequest] = await tx.$queryRaw<{ id: string }[]>(
    Prisma.sql`SELECT "id" FROM "legacy_account_link_requests" WHERE "id" = ${linkRequest.id}::uuid FOR UPDATE`
  );
  if (!lockedRequest) {
    return { kind: "failed" };
  }
  const currentRequest = await tx.legacyAccountLinkRequest.findUnique({
    where: { id: linkRequest.id },
  });
  if (
    !currentRequest ||
    currentRequest.userId !== linkRequest.userId ||
    (currentRequest.status !== LegacyAccountLinkStatus.pending &&
      currentRequest.status !== LegacyAccountLinkStatus.approved) ||
    (currentRequest.reviewedGeneration !== null &&
      transactionGeneration <= currentRequest.reviewedGeneration)
  ) {
    return { kind: "failed" };
  }
  const candidate = await tx.user.findUnique({
    select: { email: true, enabled: true, id: true, role: true },
    where: { id: currentRequest.userId },
  });
  const matchingCandidate = await tx.user.findUnique({
    select: { enabled: true, id: true, role: true },
    where: { email: identity.email },
  });
  const selectedUserId =
    matchingCandidate?.enabled && matchingCandidate.role === "user"
      ? matchingCandidate.id
      : currentRequest.userId;
  const candidateMatches =
    candidate?.enabled &&
    candidate.role === "user" &&
    candidate.email === identity.email;
  if (
    currentRequest.email !== identity.email ||
    selectedUserId !== currentRequest.userId ||
    (candidate?.email !== currentRequest.email &&
      currentRequest.reviewedGeneration === null) ||
    (currentRequest.status === LegacyAccountLinkStatus.approved &&
      !candidateMatches)
  ) {
    await invalidateLegacyAccountLinkReview(
      tx,
      currentRequest.id,
      config.providerId,
      identity.email,
      selectedUserId
    );
    return { kind: "pending" };
  }
  // Unchanged blocked pending evidence stays visible without advancing its fence.
  if (currentRequest.status === LegacyAccountLinkStatus.pending) {
    return { kind: "pending" };
  }
  const user = await linkApprovedLegacyAccount(
    tx,
    config,
    identity,
    currentRequest.id,
    currentRequest.userId,
    transactionGeneration
  );
  return user ? { kind: "linked", user } : { kind: "failed" };
}

export async function resolveLegacySsoAccount(
  config: LegacySsoConfig,
  identity: LegacyIdentity,
  transactionGeneration: bigint
): Promise<LegacySsoAccountResolution | null> {
  try {
    return await prisma.$transaction(async (tx) => {
      await tx.$executeRaw`SELECT pg_advisory_xact_lock(${accountMutationLockId})`;
      const existing = await resolveLegacySsoLinkRequest(
        tx,
        config,
        identity,
        transactionGeneration
      );
      if (existing) {
        return existing;
      }
      const user = await tx.user.create({
        data: {
          accounts: {
            create: {
              accountId: identity.subject,
              id: crypto.randomUUID(),
              issuer: config.providerId,
              providerId: config.providerId,
            },
          },
          email: identity.email,
          emailVerified: false,
          enabled: true,
          id: crypto.randomUUID(),
          mustChangePassword: false,
          name: identity.name ?? identity.email,
          role: "user",
        },
        select: { enabled: true, id: true, role: true },
      });
      return { kind: "linked" as const, user };
    });
  } catch (error) {
    if (databaseErrorCode(error) !== "P2002") {
      throw error;
    }
  }
  // Reconcile only after the optimistic transaction has rolled back.
  try {
    return await prisma.$transaction(async (tx) => {
      await tx.$executeRaw`SELECT pg_advisory_xact_lock(${accountMutationLockId})`;
      const existing = await resolveLegacySsoLinkRequest(
        tx,
        config,
        identity,
        transactionGeneration
      );
      if (existing) {
        return existing;
      }
      const candidate = await tx.user.findUnique({
        select: { id: true },
        where: { email: identity.email },
      });
      if (!candidate) {
        return { kind: "failed" as const };
      }
      const [lockedUser] = await tx.$queryRaw<{ id: string }[]>(
        Prisma.sql`SELECT "id" FROM "user" WHERE "id" = ${candidate.id} FOR UPDATE`
      );
      if (!lockedUser) {
        return { kind: "failed" as const };
      }
      await tx.legacyAccountLinkRequest.upsert({
        create: {
          email: identity.email,
          providerId: config.providerId,
          subject: identity.subject,
          userId: candidate.id,
        },
        update: {},
        where: {
          providerId_subject: {
            providerId: config.providerId,
            subject: identity.subject,
          },
        },
      });
      return await resolveLegacySsoLinkRequest(
        tx,
        config,
        identity,
        transactionGeneration
      );
    });
  } catch (error) {
    if (databaseErrorCode(error) !== "P2003") {
      throw error;
    }
  }
  return await prisma.$transaction(async (tx) => {
    await tx.$executeRaw`SELECT pg_advisory_xact_lock(${accountMutationLockId})`;
    const account = await tx.account.findUnique({
      include: { user: { select: { enabled: true, id: true, role: true } } },
      where: {
        providerId_accountId: {
          accountId: identity.subject,
          providerId: config.providerId,
        },
      },
    });
    const request = await tx.legacyAccountLinkRequest.findUnique({
      where: {
        providerId_subject: {
          providerId: config.providerId,
          subject: identity.subject,
        },
      },
    });
    if (
      request &&
      request.reviewedGeneration !== null &&
      transactionGeneration <= request.reviewedGeneration
    ) {
      return { kind: "failed" as const };
    }
    if (account) {
      return account.user.enabled && account.user.role === "user"
        ? { kind: "linked" as const, user: account.user }
        : { kind: "failed" as const };
    }
    return request?.status === LegacyAccountLinkStatus.pending &&
      request.email === identity.email &&
      (request.reviewedGeneration === null ||
        transactionGeneration > request.reviewedGeneration)
      ? { kind: "pending" as const }
      : { kind: "failed" as const };
  });
}
