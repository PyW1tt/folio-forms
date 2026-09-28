import "dotenv/config";
import { createEnv } from "@t3-oss/env-core";
import { z } from "zod";

const localDatabaseUrl =
  "postgresql://postgres:postgres@localhost:5432/onlyoffice";
const localServerOrigin = "http://localhost:3000";
const localOnlyOfficeUrl = "http://localhost:8080";

export const env = createEnv({
  emptyStringAsUndefined: true,
  runtimeEnv: process.env,
  server: {
    API_BASE: z.url().default(localServerOrigin),

    BETTER_AUTH_SECRET: z.string().min(32),

    BETTER_AUTH_URL: z.url().default(localServerOrigin),

    BOOTSTRAP_ADMIN_EMAIL: z
      .string()
      .trim()
      .toLowerCase()
      .pipe(z.email())
      .optional(),

    BOOTSTRAP_ADMIN_NAME: z.string().trim().min(1).optional(),

    BOOTSTRAP_ADMIN_PASSWORD: z.string().min(12).max(128).optional(),

    LEGACY_SSO_AUTHORIZE_URL: z.url().optional(),

    LEGACY_SSO_CALLBACK_URL: z.url().optional(),

    LEGACY_SSO_CLIENT_ID: z.string().trim().min(1).optional(),

    LEGACY_SSO_CLIENT_SECRET: z.string().min(32).optional(),

    LEGACY_SSO_EXCHANGE_URL: z.url().optional(),

    LEGACY_SSO_PROVIDER_ID: z.string().trim().min(1).max(128).optional(),

    CONVERTER_MAX_RETRIES: z.coerce.number().int().nonnegative().default(3),

    CONVERTER_POLL_INTERVAL_MS: z.coerce
      .number()
      .int()
      .positive()
      .default(1000),

    CONVERTER_TIMEOUT_MS: z.coerce.number().int().positive().default(120_000),

    CORS_ORIGIN: z.url().default("http://localhost:5173"),

    DATABASE_URL: z.url().default(localDatabaseUrl),

    EDITOR_CAPABILITY_SECRET: z.string().min(32),

    NODE_ENV: z
      .enum(["development", "production", "test"])
      .default("development"),

    OMNIROUTE_API_KEY: z.string().trim().min(1).optional(),

    OMNIROUTE_BASE_URL: z.url().optional(),

    OMNIROUTE_MODEL: z.string().trim().min(1).optional(),

    ONLYOFFICE_API_BASE: z.url().default(localOnlyOfficeUrl),

    ONLYOFFICE_DOCUMENT_BASE_URL: z
      .url()
      .default("http://host.docker.internal:3000"),

    ONLYOFFICE_INTERNAL_URL: z.url().default(localOnlyOfficeUrl),

    ONLYOFFICE_JWT_SECRET: z.string().min(32),

    ONLYOFFICE_URL: z.url().default(localOnlyOfficeUrl),

    PREFILL_HANDOFF_SECRET: z.string().min(32),

    PREFILL_RETURN_URL: z.url().default("http://localhost:3010/return"),

    RUSTFS_ACCESS_KEY_ID: z.string().min(1),

    RUSTFS_BUCKET: z.string().min(1),

    RUSTFS_ENDPOINT: z.url(),

    RUSTFS_REGION: z.string().min(1).default("us-east-1"),

    RUSTFS_SECRET_ACCESS_KEY: z.string().min(1),

    SERVER_ORIGIN: z.url().default("http://host.docker.internal:3000"),

    TEMPLATE_PATH: z
      .string()
      .default("../../onlyoffice-templates/template.docx"),
  },
  skipValidation: !!process.env.SKIP_ENV_VALIDATION,
});
