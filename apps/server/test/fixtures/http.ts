import { expect } from "bun:test";
import { createHash } from "node:crypto";
import { readdir } from "node:fs/promises";
import { tmpdir } from "node:os";

import { auth } from "@onlyoffice/auth";
import { prisma } from "@onlyoffice/db";

import type { createApp } from "../../src/app";
import type { EditorCapabilityAction } from "../../src/onlyoffice";
import { callbackClaim } from "../../src/onlyoffice";
import { DOCX_CONTENT_TYPE } from "../../src/storage";

export const jsonHeaders = { "Content-Type": "application/json" };

export const onlyOfficeSaveCapabilityFor = (config: {
  bridge?: { capabilities?: Record<"save-draft" | "submit", string> };
}): string => {
  const capability = config.bridge?.capabilities?.["save-draft"];
  if (!capability) {
    throw new Error("The ONLYOFFICE save capability is missing");
  }
  return capability;
};

export const temporaryDirectories = async () => {
  const entries = await readdir(tmpdir());
  return entries.filter((entry) => entry.startsWith("folio-authoring-"));
};

export const editorCapabilityHeaders = (capability: string) => ({
  "X-Editor-Capability": capability,
});

const trimTrailingSlashes = (value: string): string =>
  value.replace(/\/+$/u, "");

export const onlyOfficeBaseUrl = trimTrailingSlashes(
  process.env.ONLYOFFICE_URL ?? "http://localhost:8080"
);

export const apiBaseUrl = trimTrailingSlashes(
  process.env.API_BASE ?? "http://localhost:3000"
);

export const documentBaseUrl = trimTrailingSlashes(
  process.env.ONLYOFFICE_DOCUMENT_BASE_URL ?? "http://host.docker.internal:3000"
);

export const formCreationRequest = ({
  authorization,
  description = "",
  source,
  template,
  title,
}: {
  authorization?: string;
  description?: string;
  source?: "blank" | "upload";
  template?: { bytes: Uint8Array; name: string; type?: string };
  title?: string;
}): Request => {
  const body = new FormData();
  body.set("description", description);
  if (source) {
    body.set("source", source);
  }
  if (title !== undefined) {
    body.set("title", title);
  }
  if (template) {
    body.set(
      "template",
      new File([template.bytes], template.name, {
        type: template.type ?? DOCX_CONTENT_TYPE,
      })
    );
  }
  return new Request("http://test.local/api/admin/forms", {
    body,
    headers: authorization ? { Authorization: `Bearer ${authorization}` } : {},
    method: "POST",
  });
};

interface CredentialFixtureOptions {
  email: string;
  name: string;
  password: string;
  role?: "admin" | "user";
  mustChangePassword?: boolean;
}

export interface CallbackOperationFixture {
  documentKey: string;
  finalObjectKey: string;
  id: string;
  userdata: string;
}

export interface EditorConfigBody {
  apiUrl: string;
  bridge: {
    capabilities: Partial<Record<EditorCapabilityAction, string>>;
    id: string;
    lease: {
      expiresAt: string;
      id: string;
      releaseUrl: string;
      renewUrl: string;
    };
    pluginOrigin: string;
  };
  config: {
    document: { key: string; url: string };
    editorConfig: {
      callbackUrl: string;
      plugins: {
        options: Record<
          string,
          {
            authToken?: unknown;
            bridgeId: string;
            parentOrigin: string;
            publicId?: string;
            tagAliases?: Record<string, string>;
          }
        >;
        pluginsData: string[];
      };
    };
    token: string;
  };
}

export const createCredentialFixture = async ({
  email,
  mustChangePassword = false,
  name,
  password,
  role = "user",
}: CredentialFixtureOptions): Promise<{ email: string; id: string }> => {
  const id = crypto.randomUUID();
  const authContext = await auth.$context;
  const passwordHash = await authContext.password.hash(password);
  await prisma.user.create({
    data: {
      accounts: {
        create: {
          accountId: id,
          id: crypto.randomUUID(),
          issuer: "local:credential",
          password: passwordHash,
          providerId: "credential",
        },
      },
      email,
      emailVerified: true,
      enabled: true,
      id,
      mustChangePassword,
      name,
      role,
    },
  });
  return { email, id };
};

export const signIn = (
  app: ReturnType<typeof createApp>,
  email: string,
  password: string,
  sourceIp = `ticket-04-${crypto.randomUUID()}`
): Promise<Response> =>
  app.handle(
    new Request("http://test.local/api/auth/sign-in/email", {
      body: JSON.stringify({ email, password }),
      headers: { ...jsonHeaders, "x-test-ip": sourceIp },
      method: "POST",
    })
  );

export const bearerFor = async (
  app: ReturnType<typeof createApp>,
  email: string,
  password: string,
  sourceIp?: string
): Promise<string> => {
  const response = await signIn(app, email, password, sourceIp);
  expect(response.status).toBe(200);
  const body = (await response.json()) as {
    session?: { token?: string };
    token?: string;
  };
  const token =
    response.headers.get("set-auth-token")?.replace(/^Bearer\s+/iu, "") ??
    body.token ??
    body.session?.token;
  if (!token) {
    throw new Error(`Sign-in did not return a bearer token for ${email}`);
  }
  return token;
};

// oxlint-disable no-await-in-loop -- Preserve the original HTTP test's sequential operation polling.
export const waitForOperation = async (
  app: ReturnType<typeof createApp>,
  operationId: string,
  headers: Record<string, string>
): Promise<Record<string, unknown>> => {
  for (let attempt = 0; attempt < 100; attempt += 1) {
    const response = await app.handle(
      new Request(`http://test.local/api/operations/${operationId}`, {
        headers,
      })
    );
    expect(response.status).toBe(200);
    const body = (await response.json()) as {
      operation?: Record<string, unknown>;
    };
    const { operation } = body;
    if (operation?.status === "completed" || operation?.status === "failed") {
      return operation;
    }
    await Bun.sleep(5);
  }
  throw new Error(`Operation ${operationId} did not finish`);
};
// oxlint-enable no-await-in-loop

export const callbackPayload = (
  operation: CallbackOperationFixture,
  url: string
): Record<string, unknown> => ({
  key: operation.documentKey,
  status: 6,
  url,
  userdata: operation.userdata,
});

export const persistCallbackClaim = async ({
  expiresAt,
  operationId,
  userdata,
}: {
  expiresAt?: Date;
  operationId: string;
  userdata: string;
}): Promise<void> => {
  const claim = callbackClaim(userdata);
  const persistedExpiresAt =
    expiresAt ?? (claim ? new Date(claim.expiresAt * 1000) : undefined);
  if (!persistedExpiresAt) {
    throw new Error("The callback userdata did not contain a live claim");
  }
  await prisma.callbackClaim.create({
    data: {
      createdAt: new Date(
        Math.min(Date.now(), persistedExpiresAt.getTime() - 1)
      ),
      expiresAt: persistedExpiresAt,
      id: crypto.randomUUID(),
      operationId,
      tokenDigest: createHash("sha256").update(userdata).digest("hex"),
    },
  });
};

export const tamperAuthorization = (authorization: string): string => {
  const [scheme, token] = authorization.split(" ");
  const [header, payload, signature] = token?.split(".") ?? [];
  if (!scheme || !header || !payload || !signature) {
    throw new Error("The ONLYOFFICE authorization was malformed");
  }
  return `${scheme} ${header}.${payload}.${signature[0] === "a" ? "b" : "a"}${signature.slice(1)}`;
};
