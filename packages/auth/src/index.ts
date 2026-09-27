// oxlint-disable func-style -- Preserve the auth factory's exported function contract.
import { randomUUID } from "node:crypto";

import { prisma } from "@onlyoffice/db";
import { env } from "@onlyoffice/env/server";
import { betterAuth } from "better-auth";
import { prismaAdapter } from "better-auth/adapters/prisma";
import { bearer } from "better-auth/plugins";

const sessionDurationSeconds = 60 * 60;
const passwordMinimumLength = 12;
const passwordMaximumLength = 128;
const bootstrapLockId = 1_604_619_418;

const trustedOrigins = [env.CORS_ORIGIN, env.SERVER_ORIGIN];

export function createAuth() {
  return betterAuth({
    advanced: {
      defaultCookieAttributes: {
        httpOnly: true,
        sameSite: "none",
        secure: true,
      },
    },
    baseURL: env.BETTER_AUTH_URL,
    database: prismaAdapter(prisma, {
      provider: "postgresql",
      transaction: true,
    }),
    disabledPaths: ["/change-password", "/sign-up/email"],
    emailAndPassword: {
      disableSignUp: true,
      enabled: true,
      maxPasswordLength: passwordMaximumLength,
      minPasswordLength: passwordMinimumLength,
    },
    plugins: [bearer()],
    rateLimit: { enabled: false },
    session: {
      additionalFields: {
        isSso: {
          defaultValue: false,
          input: false,
          type: "boolean",
        },
      },
      disableSessionRefresh: true,
      expiresIn: sessionDurationSeconds,
    },
    trustedOrigins,
    user: {
      additionalFields: {
        enabled: {
          defaultValue: true,
          input: false,
          type: "boolean",
        },
        mustChangePassword: {
          defaultValue: false,
          input: false,
          type: "boolean",
        },
        role: {
          defaultValue: "user",
          input: false,
          type: "string",
        },
      },
    },
  });
}

export const auth = createAuth();

interface BootstrapAdmin {
  email: string;
  name: string;
  password: string;
}

function bootstrapAdminFromEnv(): BootstrapAdmin | null {
  const {
    BOOTSTRAP_ADMIN_EMAIL: email,
    BOOTSTRAP_ADMIN_NAME: name,
    BOOTSTRAP_ADMIN_PASSWORD: password,
  } = env;
  if (!email && !name && !password) {
    return null;
  }
  if (!email || !name || !password) {
    throw new Error(
      "BOOTSTRAP_ADMIN_NAME, BOOTSTRAP_ADMIN_EMAIL, and BOOTSTRAP_ADMIN_PASSWORD must be set together"
    );
  }
  return { email, name, password };
}

export async function ensureBootstrapAdmin(): Promise<boolean> {
  const bootstrapAdmin = bootstrapAdminFromEnv();
  const existingAdmin = await prisma.user.findFirst({
    select: { id: true },
    where: { role: "admin" },
  });
  if (existingAdmin) {
    return false;
  }

  if (!bootstrapAdmin) {
    throw new Error(
      "Bootstrap Admin credentials are required while no Admin exists"
    );
  }

  const context = await auth.$context;
  const passwordHash = await context.password.hash(bootstrapAdmin.password);
  return prisma.$transaction(async (tx) => {
    await tx.$executeRaw`SELECT pg_advisory_xact_lock(${bootstrapLockId})`;
    const adminCount = await tx.user.count({ where: { role: "admin" } });
    if (adminCount > 0) {
      return false;
    }

    const userId = randomUUID();
    await tx.user.create({
      data: {
        accounts: {
          create: {
            accountId: userId,
            id: randomUUID(),
            issuer: "local:credential",
            password: passwordHash,
            providerId: "credential",
          },
        },
        email: bootstrapAdmin.email,
        emailVerified: true,
        enabled: true,
        id: userId,
        mustChangePassword: true,
        name: bootstrapAdmin.name,
        role: "admin",
      },
    });
    return true;
  });
}
