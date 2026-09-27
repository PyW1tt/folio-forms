import { afterEach, expect, test } from "bun:test";
// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
import { createHash } from "node:crypto";

import { auth, ensureBootstrapAdmin } from "@onlyoffice/auth";
import { prisma } from "@onlyoffice/db";
import { strToU8, zipSync } from "fflate";
import {
  createInProcessFolioConnector,
  deterministicExternalRecord,
  startLegacySsoMock,
  startPrefillMock,
} from "prefill-mock/mock";
import type { LegacySsoMockServer, PrefillMockServer } from "prefill-mock/mock";

import {
  createApp,
  type LegacySsoConfig,
  reconcileRecoverableState,
  resolveCallbackDocumentUrl,
} from "../src/app";
import type { EditorCapabilityAction } from "../src/onlyoffice";
import {
  callbackClaim,
  createCallbackUserdata,
  createDocumentAccessToken,
  createEditorCapability,
  createOnlyOfficeBodyToken,
  createOnlyOfficeAuthorization,
  pluginGuid,
  verifyEditorCapability,
  verifyOnlyOfficeAuthorization,
} from "../src/onlyoffice";
import {
  DOCX_CONTENT_TYPE,
  deleteObject,
  objectExists,
  objectKey,
  putObject,
  readObject,
} from "../src/storage";

const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  throw new Error(
    "The HTTP application test requires DATABASE_URL for an isolated PostgreSQL database"
  );
}

const app = createApp({
  legacySso: null,
  onlyOffice: {
    convertDocxToPdf: () =>
      Promise.resolve(new TextEncoder().encode("%PDF-test")),
    forceSave: () => Promise.resolve(false),
  },
  prefillReturnUrl: "https://source.example.test/forms/return",
  requestIp: (request) => request.headers.get("x-test-ip"),
});

let externalMock: PrefillMockServer | undefined;
let legacySsoMock: LegacySsoMockServer | undefined;

afterEach(() => {
  externalMock?.close();
  externalMock = undefined;
  legacySsoMock?.close();
  legacySsoMock = undefined;
});
const jsonHeaders = { "Content-Type": "application/json" };
const trimTrailingSlashes = (value: string): string =>
  value.replace(/\/+$/u, "");
const onlyOfficeBaseUrl = trimTrailingSlashes(
  process.env.ONLYOFFICE_URL ?? "http://localhost:8080"
);
const apiBaseUrl = trimTrailingSlashes(
  process.env.API_BASE ?? "http://localhost:3000"
);
const documentBaseUrl = trimTrailingSlashes(
  process.env.ONLYOFFICE_DOCUMENT_BASE_URL ?? "http://host.docker.internal:3000"
);

const maxTemplateUploadBytes = 25 * 1024 * 1024;
const templateContentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const templateRelationshipsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const docxXmlFixture = ({
  additionalParts = {},
  contentTypes = templateContentTypesXml,
  document,
  paddingBytes = 0,
  relationships = templateRelationshipsXml,
}: {
  additionalParts?: Record<string, Uint8Array>;
  contentTypes?: string;
  document: string;
  paddingBytes?: number;
  relationships?: string;
}): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(relationships),
      "word/document.xml": strToU8(document),
      "word/media/padding.bin": new Uint8Array(paddingBytes),
      ...additionalParts,
    },
    { level: 0 }
  );

const docxFixture = (label: string, paddingBytes = 0): Uint8Array =>
  docxXmlFixture({
    document: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${label}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    paddingBytes,
  });
const contentControlDocument = (controls: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:word="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:body>${controls}<w:sectPr/></w:body></w:document>`;

const contentControl = ({
  alias,
  placeholderText,
  tag,
  type,
}: {
  alias?: string;
  placeholderText?: string;
  tag: string;
  type: string;
}): string => {
  const aliasProperty =
    alias === undefined ? "" : `<w:alias w:val="${alias}"/>`;
  const placeholderProperty =
    placeholderText === undefined ? "" : "<w:showingPlcHdr/>";
  const content = placeholderText ?? "fixture";
  return `<w:sdt><w:sdtPr>${aliasProperty}<w:tag w:val="${tag}"/>${type}${placeholderProperty}</w:sdtPr><w:sdtContent><w:r><w:t>${content}</w:t></w:r></w:sdtContent></w:sdt>`;
};
const pictureDrawing = (relationshipId: string): string =>
  `<w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blip r:embed="${relationshipId}"/></w:drawing>`;
const onePixelPng = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
    "base64"
  )
);
const pngFixture = (
  width = 1,
  height = 1,
  byteLength = onePixelPng.byteLength
): Uint8Array => {
  const bytes = new Uint8Array(byteLength);
  bytes.set(onePixelPng.subarray(0, Math.min(onePixelPng.length, byteLength)));
  const writeUint32 = (value: number, offset: number) => {
    bytes[offset] = Math.floor(value / 0x1_00_00_00) % 0x1_00;
    bytes[offset + 1] = Math.floor(value / 0x1_00_00) % 0x1_00;
    bytes[offset + 2] = Math.floor(value / 0x1_00) % 0x1_00;
    bytes[offset + 3] = value % 0x1_00;
  };
  writeUint32(width, 16);
  writeUint32(height, 20);
  return bytes;
};
const jpegFixture = (width = 1, height = 1): Uint8Array =>
  Uint8Array.from([
    0xff,
    0xd8,
    0xff,
    0xc0,
    0x00,
    0x11,
    0x08,
    Math.floor(height / 0x1_00),
    height % 0x1_00,
    Math.floor(width / 0x1_00),
    width % 0x1_00,
    0x03,
    0x01,
    0x11,
    0x00,
    0x02,
    0x11,
    0x00,
    0x03,
    0x11,
    0x00,
    0xff,
    0xd9,
  ]);
const gifFixture = Uint8Array.from(
  Buffer.from("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==", "base64")
);
const pictureDocumentFixture = ({
  images = [],
  includeStaticImage = false,
  showingPlaceholder = false,
}: {
  images?: { bytes: Uint8Array; extension: string }[];
  includeStaticImage?: boolean;
  showingPlaceholder?: boolean;
} = {}): Uint8Array => {
  const imageEntries = images.map((image, index) => ({
    bytes: image.bytes,
    extension: image.extension,
    relationshipId: `rIdPicture${index + 1}`,
  }));
  if (includeStaticImage) {
    imageEntries.push({
      bytes: pngFixture(),
      extension: "png",
      relationshipId: "rIdStatic",
    });
  }
  const relationships = imageEntries.map(
    ({ extension, relationshipId }, index) =>
      `<Relationship Id="${relationshipId}" Target="media/image${index + 1}.${extension}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>`
  );
  const staticDrawing = includeStaticImage ? pictureDrawing("rIdStatic") : "";
  const pictureDrawings = imageEntries
    .slice(0, images.length)
    .map(({ relationshipId }) => pictureDrawing(relationshipId))
    .join("");
  const additionalParts: Record<string, Uint8Array> = {
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`
    ),
  };
  for (const [index, image] of imageEntries.entries()) {
    additionalParts[`word/media/image${index + 1}.${image.extension}`] =
      image.bytes;
  }
  const extensions = new Set(imageEntries.map(({ extension }) => extension));
  const contentTypes = templateContentTypesXml.replace(
    "</Types>",
    `${[...extensions]
      .map(
        (extension) =>
          `<Default Extension="${extension}" ContentType="image/${extension === "jpg" ? "jpeg" : extension}"/>`
      )
      .join("")}</Types>`
  );
  const pictureProperties = showingPlaceholder ? "<w:showingPlcHdr/>" : "";
  return docxXmlFixture({
    additionalParts,
    contentTypes,
    document: contentControlDocument(
      `${staticDrawing}<w:sdt><w:sdtPr><w:tag w:val="photo"/><w:picture/>${pictureProperties}</w:sdtPr><w:sdtContent><w:r>${pictureDrawings || "<w:t>empty</w:t>"}</w:r></w:sdtContent></w:sdt>`
    ),
  });
};

const strictDocxFixture = (label: string): Uint8Array =>
  docxXmlFixture({
    contentTypes: `<?xml version="1.0"?><ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><ct:Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" PartName="/word/document.xml"/></ct:Types>`,
    document: `<?xml version="1.0"?><word:document xmlns:word="http://purl.oclc.org/ooxml/wordprocessingml/main"><word:body><word:p><word:r><word:t>${label}</word:t></word:r></word:p></word:body></word:document>`,
    relationships: `<?xml version="1.0"?><pkg:Relationships xmlns:pkg="http://schemas.openxmlformats.org/package/2006/relationships"><pkg:Relationship Id="rId1" Target="/word/document.xml" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"/></pkg:Relationships>`,
  });

const utf16Xml = (value: string, byteOrder: "be" | "le"): Uint8Array => {
  const littleEndian = Buffer.from(value, "utf16le");
  const bytes = new Uint8Array(littleEndian.byteLength + 2);
  bytes[0] = byteOrder === "le" ? 0xff : 0xfe;
  bytes[1] = byteOrder === "le" ? 0xfe : 0xff;
  for (let index = 0; index < littleEndian.byteLength; index += 2) {
    const target = index + 2;
    const firstByte = littleEndian[index] ?? 0;
    const secondByte = littleEndian[index + 1] ?? 0;
    bytes[target] = byteOrder === "le" ? firstByte : secondByte;
    bytes[target + 1] = byteOrder === "le" ? secondByte : firstByte;
  }
  return bytes;
};

const utf16DocxFixture = (): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": utf16Xml(
        templateContentTypesXml.replace(
          'encoding="UTF-8"',
          'encoding="UTF-16"'
        ),
        "le"
      ),
      "_rels/.rels": utf16Xml(
        templateRelationshipsXml.replace(
          'encoding="UTF-8"',
          'encoding="UTF-16"'
        ),
        "be"
      ),
      "word/document.xml": utf16Xml(
        '<?xml version="1.0" encoding="UTF-16"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
        "le"
      ),
    },
    { level: 0 }
  );

const sizedDocxFixture = (label: string, byteLength: number): Uint8Array => {
  const emptyPadding = docxFixture(label);
  const paddingLength = byteLength - emptyPadding.byteLength;
  if (paddingLength < 0) {
    throw new Error("The requested DOCX fixture size is too small");
  }
  const fixture = docxFixture(label, paddingLength);
  if (fixture.byteLength !== byteLength) {
    throw new Error("The DOCX fixture did not reach the requested size");
  }
  return fixture;
};

const formCreationRequest = ({
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
interface CallbackOperationFixture {
  documentKey: string;
  finalObjectKey: string;
  id: string;
  userdata: string;
}

interface EditorConfigBody {
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
          }
        >;
        pluginsData: string[];
      };
    };
    token: string;
  };
}

const createCredentialFixture = async ({
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

const signIn = (
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

const bearerFor = async (
  email: string,
  password: string,
  sourceIp?: string
): Promise<string> => {
  const response = await signIn(email, password, sourceIp);
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
type LegacySsoHttpHandler = (
  request: Request
) => Response | Promise<Response>;

interface LegacySsoBrowserTransaction {
  authorizationUrl: string;
  preLoginCookie: string;
  startResponse: Response;
}

const createLegacySsoTestApp = (
  config: LegacySsoConfig,
  clock: () => Date = () => new Date()
) =>
  createApp({
    clock,
    legacySso: config,
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.resolve(new TextEncoder().encode("%PDF-test")),
      forceSave: () => Promise.resolve(false),
    },
    prefillReturnUrl: "https://source.example.test/forms/return",
    requestIp: (request) => request.headers.get("x-test-ip"),
  });

const cookiePairFrom = (response: Response, name: string): string => {
  const setCookie = response.headers.get("set-cookie") ?? "";
  for (const cookie of setCookie.split(/,\s*(?=[^;,]+=)/u)) {
    const pair = cookie.split(";", 1)[0]?.trim();
    if (pair?.startsWith(`${name}=`)) {
      return pair;
    }
  }
  throw new Error(`The ${name} cookie was not set`);
};

const startLegacySsoBrowser = async (
  handle: LegacySsoHttpHandler,
  returnTo?: string
): Promise<LegacySsoBrowserTransaction> => {
  const startUrl = new URL(
    "/api/legacy-sso/start",
    "https://folio.example.test"
  );
  if (returnTo !== undefined) {
    startUrl.searchParams.set("returnTo", returnTo);
  }
  const startResponse = await handle(new Request(startUrl));
  const authorizationUrl = startResponse.headers.get("location");
  if (!authorizationUrl) {
    throw new Error("The legacy SSO start did not redirect to the old backend");
  }
  return {
    authorizationUrl,
    preLoginCookie: cookiePairFrom(startResponse, "__Host-folio-sso"),
    startResponse,
  };
};

const waitForOperation = async (
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
const callbackPayload = (
  operation: CallbackOperationFixture,
  url: string
): Record<string, unknown> => ({
  key: operation.documentKey,
  status: 6,
  url,
  userdata: operation.userdata,
});
const persistCallbackClaim = async ({
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

const tamperAuthorization = (authorization: string): string => {
  const [scheme, token] = authorization.split(" ");
  const [header, payload, signature] = token?.split(".") ?? [];
  if (!scheme || !header || !payload || !signature) {
    throw new Error("The ONLYOFFICE authorization was malformed");
  }
  return `${scheme} ${header}.${payload}.${signature[0] === "a" ? "b" : "a"}${signature.slice(1)}`;
};

test("rewrites public ONLYOFFICE callback paths to the internal base", () => {
  expect(
    resolveCallbackDocumentUrl(
      "http://localhost:8080/office/cache/files/data/example/output.docx?md5=x",
      new Set(["http://localhost:8080"]),
      "http://localhost:8080/office",
      "http://localhost:8081"
    )
  ).toBe("http://localhost:8081/cache/files/data/example/output.docx?md5=x");
});


const legacySsoCallbackUrl =
  "https://folio.example.test/api/legacy-sso/callback";
const legacySsoWebOrigin = new URL(
  process.env.CORS_ORIGIN ?? "http://localhost:5173"
).origin;
const legacySsoFailedLocation = new URL(
  "/login?legacySso=failed",
  legacySsoWebOrigin
).href;

test("linked role=user completes browser-bound SSO and receives a normal session", async () => {
  const email = `ticket-12-linked-${crypto.randomUUID()}@example.com`;
  const password = "Ticket12-local-password";
  const user = await createCredentialFixture({
    email,
    mustChangePassword: true,
    name: "Ticket 12 linked User",
    password,
  });
  const identity = {
    email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const authorizationCode = "ticket-12+/=opaque-code";
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    identity,
    authorizationCode,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const returnTo = `/forms/${"a".repeat(32)}/fill?responseId=${crypto.randomUUID()}`;
  const browserStart = await startLegacySsoBrowser(handle, returnTo);
  expect(browserStart.startResponse.status).toBe(303);
  const startCookieHeader =
    browserStart.startResponse.headers.get("set-cookie") ?? "";
  expect(startCookieHeader).toContain("HttpOnly");
  expect(startCookieHeader).toContain("Secure");
  expect(startCookieHeader).toContain("SameSite=Lax");
  expect(startCookieHeader).toContain("Path=/");

  const authorizeUrl = new URL(browserStart.authorizationUrl);
  expect(authorizeUrl.searchParams.get("client_id")).toBe(clientId);
  expect(authorizeUrl.searchParams.get("redirect_uri")).toBe(
    legacySsoCallbackUrl
  );
  expect(authorizeUrl.searchParams.get("code_challenge_method")).toBe("S256");
  expect(authorizeUrl.searchParams.get("state")).toMatch(/^[A-Za-z0-9_-]{43}$/u);
  expect(authorizeUrl.searchParams.get("code_challenge")).toMatch(
    /^[A-Za-z0-9_-]{43}$/u
  );
  expect(authorizeUrl.searchParams.has("returnTo")).toBe(false);
  expect(browserStart.authorizationUrl).not.toContain(clientSecret);

  const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(oldBackendResponse.status).toBe(303);
  const callbackLocation = oldBackendResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not return a callback");
  }
  const callbackUrl = new URL(callbackLocation);
  expect([...callbackUrl.searchParams.keys()].sort()).toEqual(["code", "state"]);
  expect(callbackUrl.searchParams.get("code")).toBe(authorizationCode);
  const callbackResponse = await handle(
    new Request(callbackUrl, {
      headers: {
        Cookie: browserStart.preLoginCookie,
        Host: "attacker.example.test",
        "X-Forwarded-Host": "attacker.example.test",
        "X-Forwarded-Proto": "https",
      },
    })
  );
  expect(callbackResponse.status).toBe(303);
  expect(callbackResponse.headers.get("location")).toBe(
    new URL(returnTo, legacySsoWebOrigin).href
  );
  expect(callbackResponse.headers.get("referrer-policy")).toBe("no-referrer");
  expect(callbackLocation).not.toContain(clientSecret);
  const callbackCookieHeader =
    callbackResponse.headers.get("set-cookie") ?? "";
  expect(callbackCookieHeader).toContain("HttpOnly");
  expect(callbackCookieHeader).toContain("Secure");
  expect(callbackCookieHeader).toContain("SameSite=Lax");

  const sessionCookie = cookiePairFrom(
    callbackResponse,
    "__Host-folio-sso-session"
  );
  const wrongOriginClaim = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: sessionCookie,
        Origin: "https://attacker.example.test",
      },
      method: "POST",
    })
  );
  expect(wrongOriginClaim.status).toBe(403);
  expect(wrongOriginClaim.headers.has("set-cookie")).toBe(false);
  const claimResponse = await handle(
    new Request(
      "https://folio.example.test/api/legacy-sso/session",
      {
        headers: {
          Cookie: sessionCookie,
          Origin: "https://folio.example.test",
        },
        method: "POST",
      }
    )
  );
  expect(claimResponse.status).toBe(200);
  const claimBody = await claimResponse.json();
  if (
    !claimBody ||
    typeof claimBody !== "object" ||
    Array.isArray(claimBody) ||
    !("token" in claimBody) ||
    typeof claimBody.token !== "string"
  ) {
    throw new Error("The SSO session handoff did not return a bearer");
  }
  const bearer = claimBody.token;
  expect(callbackResponse.headers.get("location")).not.toContain(bearer);
  const sessionResponse = await handle(
    new Request("https://folio.example.test/api/session", {
      headers: { Authorization: `Bearer ${bearer}` },
    })
  );
  expect(sessionResponse.status).toBe(200);
  expect(await sessionResponse.json()).toMatchObject({
    user: {
      email,
      id: user.id,
      mustChangePassword: false,
      name: "Ticket 12 linked User",
      role: "user",
    },
  });
  const ssoFormResponse = await handle(
    new Request(`https://folio.example.test/api/forms/${"b".repeat(32)}`, {
      headers: { Authorization: `Bearer ${bearer}` },
    })
  );
  expect(ssoFormResponse.status).toBe(404);
  const editorCapability = createEditorCapability({
    action: "save-draft",
    actorId: user.id,
    documentKey: "ticket-12-document",
    formId: crypto.randomUUID(),
    leaseId: crypto.randomUUID(),
    leaseProof: "ticket-12-lease-proof",
    role: "user",
    targetId: crypto.randomUUID(),
    targetType: "response",
  });
  const capabilityRequest = (token: string) =>
    handle(
      new Request(
        `https://folio.example.test/api/forms/${"b".repeat(32)}/draft`,
        {
          body: JSON.stringify({
            documentKey: "ticket-12-document",
            responseId: crypto.randomUUID(),
          }),
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${token}`,
            "X-Editor-Capability": editorCapability,
          },
          method: "POST",
        }
      )
    );
  expect((await capabilityRequest(bearer)).status).toBe(404);

  const localBearer = await bearerFor(
    email,
    password,
    `ticket-12-local-${crypto.randomUUID()}`
  );
  const localFormResponse = await handle(
    new Request(`https://folio.example.test/api/forms/${"b".repeat(32)}`, {
      headers: { Authorization: `Bearer ${localBearer}` },
    })
  );
  expect(localFormResponse.status).toBe(403);
  expect(await localFormResponse.json()).toMatchObject({
    error: "password_change_required",
  });
  expect((await capabilityRequest(localBearer)).status).toBe(401);

  const callbackReplay = await handle(
    new Request(callbackUrl, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(callbackReplay.status).toBe(303);
  expect(callbackReplay.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  const sessionReplay = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: sessionCookie,
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(sessionReplay.status).toBe(401);
});

test("rejects unsafe returns and mismatched callback state or cookie", async () => {
  const email = `ticket-12-binding-${crypto.randomUUID()}@example.com`;
  const password = "Ticket12-binding-test-password";
  const user = await createCredentialFixture({
    email,
    name: "Ticket 12 binding User",
    password,
  });
  const identity = {
    email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    identity,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  for (const unsafeReturn of [
    "//attacker.example/",
    "https://attacker.example/",
    "/admin",
    `/forms/${"a".repeat(32)}/fill?access_token=secret`,
  ]) {
    const unsafeStartUrl = new URL(
      "/api/legacy-sso/start",
      "https://folio.example.test"
    );
    unsafeStartUrl.searchParams.set("returnTo", unsafeReturn);
    const unsafeStart = await handle(new Request(unsafeStartUrl));
    expect(unsafeStart.status).toBe(400);
    expect(unsafeStart.headers.has("set-cookie")).toBe(false);
  }

  const returnTo = `/forms/${"c".repeat(32)}/fill`;
  const browserStart = await startLegacySsoBrowser(handle, returnTo);
  const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(oldBackendResponse.status).toBe(303);
  const callbackLocation = oldBackendResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not return a callback");
  }
  const callbackUrl = new URL(callbackLocation);
  const wrongStateUrl = new URL(callbackUrl);
  wrongStateUrl.searchParams.set("state", "x".repeat(43));
  const wrongStateResponse = await handle(
    new Request(wrongStateUrl, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(wrongStateResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  const wrongCookieResponse = await handle(
    new Request(callbackUrl, {
      headers: { Cookie: `__Host-folio-sso=${"y".repeat(43)}` },
    })
  );
  expect(wrongCookieResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  const wrongHostUrl = new URL(callbackUrl);
  wrongHostUrl.hostname = "attacker.example.test";
  const wrongHostResponse = await handle(
    new Request(wrongHostUrl, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(wrongHostResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );
  const wrongProtocolResponse = await handle(
    new Request(callbackUrl, {
      headers: {
        Cookie: browserStart.preLoginCookie,
        "X-Forwarded-Proto": "http",
      },
    })
  );
  expect(wrongProtocolResponse.headers.get("location")).toBe(
    legacySsoFailedLocation
  );

  const validCallback = await handle(
    new Request(callbackUrl, {
      headers: { Cookie: browserStart.preLoginCookie },
    })
  );
  expect(validCallback.status).toBe(303);
  expect(validCallback.headers.get("location")).toBe(
    new URL(returnTo, legacySsoWebOrigin).href
  );
  expect(
    await (
      await handle(
        new Request("https://folio.example.test/api/legacy-sso/status")
      )
    ).json()
  ).toEqual({ enabled: true });

  const separatedHandoff = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: prefillCookie,
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(separatedHandoff.status).toBe(401);
  expect(
    (await handle(new Request("https://folio.example.test/prefill/handoff")))
      .status
  ).toBe(405);
  expect(
    await (
      await app.handle(new Request("http://test.local/api/legacy-sso/status"))
    ).json()
  ).toEqual({ enabled: false });
});

test("rejects unlinked, unverified, disabled, email-mismatched, and Admin identities", async () => {
  const email = `ticket-12-eligibility-${crypto.randomUUID()}@example.com`;
  const password = "Ticket12-eligibility-test-password";
  const user = await createCredentialFixture({
    email,
    name: "Ticket 12 eligible User",
    password,
  });
  const identity = {
    email,
    email_verified: false,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const linkedSubject = identity.sub;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    identity,
  });
  const ssoApp = createLegacySsoTestApp({
    authorizeUrl: legacySsoMock.authorizeUrl,
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    exchangeUrl: legacySsoMock.exchangeUrl,
    providerId,
  });
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const rejectCurrentIdentity = async () => {
    const browserStart = await startLegacySsoBrowser(
      handle,
      `/forms/${"d".repeat(32)}/fill`
    );
    const oldBackendResponse = await fetch(browserStart.authorizationUrl, {
      redirect: "manual",
    });
    expect(oldBackendResponse.status).toBe(303);
    const location = oldBackendResponse.headers.get("location");
    if (!location) {
      throw new Error("The test old backend did not return a callback");
    }
    const callbackResponse = await handle(
      new Request(location, {
        headers: { Cookie: browserStart.preLoginCookie },
      })
    );
    expect(callbackResponse.status).toBe(303);
    expect(callbackResponse.headers.get("location")).toBe(
      legacySsoFailedLocation
    );
  };
  const userCountBefore = await prisma.user.count();
  await rejectCurrentIdentity();
  identity.email_verified = true;
  identity.email = `different-${crypto.randomUUID()}@example.com`;
  await rejectCurrentIdentity();
  identity.email = email;
  identity.sub = `unlinked-${crypto.randomUUID()}`;
  await rejectCurrentIdentity();
  identity.sub = linkedSubject;
  await prisma.user.update({
    data: { enabled: false },
    where: { id: user.id },
  });
  await rejectCurrentIdentity();
  await prisma.user.update({
    data: { enabled: true, role: "admin" },
    where: { id: user.id },
  });
  await rejectCurrentIdentity();
  expect(await prisma.user.count()).toBe(userCountBefore);
  expect(await prisma.session.count({ where: { userId: user.id } })).toBe(0);
});

test("expires browser-bound SSO transactions and session handoffs", async () => {
  let appNow = new Date();
  const backendNow = new Date();
  const email = `ticket-12-expiry-${crypto.randomUUID()}@example.com`;
  const user = await createCredentialFixture({
    email,
    name: "Ticket 12 expiry User",
    password: "Ticket12-expiry-test-password",
  });
  const identity = {
    email,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const providerId = `legacy-${crypto.randomUUID()}`;
  await prisma.account.create({
    data: {
      accountId: identity.sub,
      id: crypto.randomUUID(),
      providerId,
      userId: user.id,
    },
  });
  legacySsoMock = startLegacySsoMock({
    callbackUrl: legacySsoCallbackUrl,
    clientId,
    clientSecret,
    clock: () => new Date(backendNow),
    identity,
  });
  const ssoApp = createLegacySsoTestApp(
    {
      authorizeUrl: legacySsoMock.authorizeUrl,
      callbackUrl: legacySsoCallbackUrl,
      clientId,
      clientSecret,
      exchangeUrl: legacySsoMock.exchangeUrl,
      providerId,
    },
    () => new Date(appNow)
  );
  const handle: LegacySsoHttpHandler = (request) => ssoApp.handle(request);
  const returnTo = `/forms/${"e".repeat(32)}/fill`;
  const expiredStart = await startLegacySsoBrowser(handle, returnTo);
  const expiredAuthorization = await fetch(expiredStart.authorizationUrl, {
    redirect: "manual",
  });
  expect(expiredAuthorization.status).toBe(303);
  const expiredCallbackUrl = expiredAuthorization.headers.get("location");
  if (!expiredCallbackUrl) {
    throw new Error("The test old backend did not return an expiring code");
  }
  appNow = new Date(appNow.getTime() + 5 * 60_000 + 1);
  const expiredCallback = await handle(
    new Request(expiredCallbackUrl, {
      headers: { Cookie: expiredStart.preLoginCookie },
    })
  );
  expect(expiredCallback.headers.get("location")).toBe(
    legacySsoFailedLocation
  );

  const validStart = await startLegacySsoBrowser(handle, returnTo);
  const validAuthorization = await fetch(validStart.authorizationUrl, {
    redirect: "manual",
  });
  const validCallbackUrl = validAuthorization.headers.get("location");
  if (!validCallbackUrl) {
    throw new Error("The test old backend did not return a valid code");
  }
  const validCallback = await handle(
    new Request(validCallbackUrl, {
      headers: { Cookie: validStart.preLoginCookie },
    })
  );
  expect(validCallback.headers.get("location")).toBe(
    new URL(returnTo, legacySsoWebOrigin).href
  );
  const sessionCookie = cookiePairFrom(
    validCallback,
    "__Host-folio-sso-session"
  );
  appNow = new Date(appNow.getTime() + 60_001);
  const expiredSession = await handle(
    new Request("https://folio.example.test/api/legacy-sso/session", {
      headers: {
        Cookie: sessionCookie,
        Origin: "https://folio.example.test",
      },
      method: "POST",
    })
  );
  expect(expiredSession.status).toBe(401);
});

test("old-backend codes bind client, callback, verifier, expiry, and single use", async () => {
  let now = new Date();
  const callbackUrl = legacySsoCallbackUrl;
  const clientId = `folio-${crypto.randomUUID()}`;
  const clientSecret = `ticket-12-test-secret-${crypto.randomUUID()}`;
  const identity = {
    email: `ticket-12-mock-${crypto.randomUUID()}@example.com`,
    email_verified: true,
    sub: `legacy-${crypto.randomUUID()}`,
  };
  legacySsoMock = startLegacySsoMock({
    callbackUrl,
    clientId,
    clientSecret,
    codeLifetimeMs: 1000,
    clock: () => new Date(now),
    identity,
  });
  const verifier = "v".repeat(43);
  const challenge = createHash("sha256")
    .update(verifier)
    .digest("base64url");
  const state = "s".repeat(43);
  const authorize = new URL(legacySsoMock.authorizeUrl);
  authorize.searchParams.set("client_id", clientId);
  authorize.searchParams.set("redirect_uri", callbackUrl);
  authorize.searchParams.set("state", state);
  authorize.searchParams.set("code_challenge", challenge);
  authorize.searchParams.set("code_challenge_method", "S256");

  const wrongClientAuthorize = new URL(authorize);
  wrongClientAuthorize.searchParams.set("client_id", "other-client");
  expect(
    (await fetch(wrongClientAuthorize, { redirect: "manual" })).status
  ).toBe(400);
  const wrongCallbackAuthorize = new URL(authorize);
  wrongCallbackAuthorize.searchParams.set(
    "redirect_uri",
    "https://attacker.example.test/callback"
  );
  expect(
    (await fetch(wrongCallbackAuthorize, { redirect: "manual" })).status
  ).toBe(400);
  const authorizeResponse = await fetch(authorize, { redirect: "manual" });
  expect(authorizeResponse.status).toBe(303);
  const callbackLocation = authorizeResponse.headers.get("location");
  if (!callbackLocation) {
    throw new Error("The test old backend did not issue an authorization code");
  }
  const code = new URL(callbackLocation).searchParams.get("code");
  if (!code) {
    throw new Error("The test old backend callback omitted its code");
  }

  const exchange = async ({
    clientId: exchangeClientId = clientId,
    codeVerifier = verifier,
    redirectUri = callbackUrl,
  }: {
    clientId?: string;
    codeVerifier?: string;
    redirectUri?: string;
  } = {}) =>
    fetch(legacySsoMock.exchangeUrl, {
      body: new URLSearchParams({
        client_id: exchangeClientId,
        code,
        code_verifier: codeVerifier,
        grant_type: "authorization_code",
        redirect_uri: redirectUri,
      }),
      headers: {
        Authorization: `Basic ${Buffer.from(
          `${clientId}:${clientSecret}`
        ).toString("base64")}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      method: "POST",
    });
  expect((await exchange({ clientId: "other-client" })).status).toBe(400);
  expect(
    (
      await exchange({
        redirectUri: "https://attacker.example.test/callback",
      })
    ).status
  ).toBe(400);
  expect((await exchange({ codeVerifier: "w".repeat(43) })).status).toBe(400);
  expect((await exchange()).status).toBe(200);
  expect((await exchange()).status).toBe(400);

  const expiredAuthorization = await fetch(authorize, { redirect: "manual" });
  const expiredLocation = expiredAuthorization.headers.get("location");
  if (!expiredLocation) {
    throw new Error("The test old backend did not issue an expiring code");
  }
  const expiredCode = new URL(expiredLocation).searchParams.get("code");
  if (!expiredCode) {
    throw new Error("The test old backend callback omitted its expiring code");
  }
  now = new Date(now.getTime() + 1001);
  const expiredResponse = await fetch(legacySsoMock.exchangeUrl, {
    body: new URLSearchParams({
      client_id: clientId,
      code: expiredCode,
      code_verifier: verifier,
      grant_type: "authorization_code",
      redirect_uri: callbackUrl,
    }),
    headers: {
      Authorization: `Basic ${Buffer.from(
        `${clientId}:${clientSecret}`
      ).toString("base64")}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    method: "POST",
  });
  expect(expiredResponse.status).toBe(400);
});
test("serves authenticated Admin and User workflows through HTTP", async () => {
  const adminEmail = `ticket-02-admin-${crypto.randomUUID()}@example.com`;
  const userEmail = `ticket-02-user-${crypto.randomUUID()}@example.com`;
  const password = "Ticket02-password-for-test";
  const otherUserEmail = `ticket-03-other-${crypto.randomUUID()}@example.com`;
  const bootstrapEmail =
    process.env.BOOTSTRAP_ADMIN_EMAIL?.trim().toLowerCase();
  const bootstrapName = process.env.BOOTSTRAP_ADMIN_NAME?.trim();
  if (
    !bootstrapEmail ||
    !bootstrapName ||
    !process.env.BOOTSTRAP_ADMIN_PASSWORD
  ) {
    throw new Error(
      "The HTTP application test requires all BOOTSTRAP_ADMIN_* variables"
    );
  }
  expect(
    resolveCallbackDocumentUrl(
      "https://docs.example//169.254.169.254/latest/meta-data?key=value",
      new Set(["https://docs.example"]),
      "https://docs.example",
      "http://onlyoffice"
    )
  ).toBe("http://onlyoffice//169.254.169.254/latest/meta-data?key=value");
  const callbackDocumentUrl = new URL(
    "/onlyoffice/document/template.docx",
    process.env.ONLYOFFICE_DOCUMENT_BASE_URL ??
      "http://host.docker.internal:3000"
  ).toString();
  expect(resolveCallbackDocumentUrl(callbackDocumentUrl)).toBe(
    callbackDocumentUrl
  );
  expect(
    resolveCallbackDocumentUrl(
      "https://user:password@docs.example/document.docx",
      new Set(["https://docs.example"]),
      "https://docs.example",
      "http://onlyoffice"
    )
  ).toBeNull();
  const bootstrapUserSelect = {
    createdAt: true,
    email: true,
    enabled: true,
    id: true,
    mustChangePassword: true,
    name: true,
    role: true,
    updatedAt: true,
  } as const;

  const existingBootstrap = await prisma.user.findUnique({
    select: bootstrapUserSelect,
    where: { email: bootstrapEmail },
  });
  const existingBootstrapAccount = existingBootstrap
    ? await prisma.account.findFirst({
        where: { providerId: "credential", userId: existingBootstrap.id },
      })
    : null;
  const firstBootstrapResult = await ensureBootstrapAdmin();
  expect(firstBootstrapResult).toBe(!existingBootstrap);
  const bootstrapBefore = await prisma.user.findUnique({
    select: bootstrapUserSelect,
    where: { email: bootstrapEmail },
  });
  if (!bootstrapBefore) {
    throw new Error("Bootstrap Admin was not created");
  }
  const bootstrapAccountBefore = await prisma.account.findFirst({
    where: { providerId: "credential", userId: bootstrapBefore.id },
  });
  if (existingBootstrap) {
    expect(bootstrapBefore).toEqual(existingBootstrap);
    expect(bootstrapAccountBefore).toEqual(existingBootstrapAccount);
  } else {
    expect(bootstrapBefore).toMatchObject({
      email: bootstrapEmail,
      enabled: true,
      mustChangePassword: true,
      name: bootstrapName,
      role: "admin",
    });
    expect(bootstrapAccountBefore).toMatchObject({
      accountId: bootstrapBefore.id,
      issuer: "local:credential",
      password: expect.any(String),
      providerId: "credential",
      userId: bootstrapBefore.id,
    });
  }
  const adminCountBeforeSecondEnsure = await prisma.user.count({
    where: { role: "admin" },
  });

  const secondBootstrapResult = await ensureBootstrapAdmin();
  expect(secondBootstrapResult).toBe(false);
  expect(
    await prisma.user.count({
      where: { role: "admin" },
    })
  ).toBe(adminCountBeforeSecondEnsure);
  const bootstrapAfter = await prisma.user.findUnique({
    select: bootstrapUserSelect,
    where: { id: bootstrapBefore.id },
  });
  const bootstrapAccountAfter = await prisma.account.findFirst({
    where: { providerId: "credential", userId: bootstrapBefore.id },
  });
  expect(bootstrapAfter).toEqual(bootstrapBefore);
  expect(bootstrapAccountAfter).toEqual(bootstrapAccountBefore);

  const signupEmail = `ticket-04-signup-${crypto.randomUUID()}@example.com`;
  const signupResponse = await app.handle(
    new Request("http://test.local/api/auth/sign-up/email", {
      body: JSON.stringify({
        email: signupEmail,
        name: "Ticket 04 Direct Signup",
        password,
      }),
      headers: jsonHeaders,
      method: "POST",
    })
  );
  expect(signupResponse.status).toBe(404);
  expect(
    await prisma.user.findUnique({
      where: { email: signupEmail },
    })
  ).toBeNull();

  const healthResponse = await app.handle(
    new Request("http://test.local/health")
  );
  expect(healthResponse.status).toBe(200);
  expect(await healthResponse.json()).toEqual({ ok: true });

  const unauthorizedResponse = await app.handle(
    formCreationRequest({ source: "blank", title: "Denied Form" })
  );
  expect(unauthorizedResponse.status).toBe(401);
  expect(await unauthorizedResponse.json()).toMatchObject({
    error: "unauthorized",
  });

  const admin = await createCredentialFixture({
    email: adminEmail,
    name: "Ticket 04 Workflow Admin",
    password,
    role: "admin",
  });
  const adminId = admin.id;
  const adminBearer = await bearerFor(adminEmail, password);
  const secretFormTitle = `Ticket 04 secret title ${crypto.randomUUID()}`;
  const secretFormDescription = `Ticket 04 secret description ${crypto.randomUUID()}`;

  const jsonCreateResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      body: JSON.stringify({ source: "blank", title: "JSON is not accepted" }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${adminBearer}`,
      },
      method: "POST",
    })
  );
  expect(jsonCreateResponse.status).toBe(415);
  expect(await jsonCreateResponse.json()).toMatchObject({
    error: "invalid_file_type",
  });
  for (const invalidRequest of [
    formCreationRequest({
      authorization: adminBearer,
      title: "Missing source",
    }),
    formCreationRequest({
      authorization: adminBearer,
      source: "blank",
    }),
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      title: "Missing uploaded template",
    }),
    formCreationRequest({
      authorization: adminBearer,
      source: "blank",
      template: {
        bytes: docxFixture("unexpected-blank-template"),
        name: "unexpected.docx",
      },
      title: "Unexpected blank template",
    }),
  ]) {
    const invalidResponse = await app.handle(invalidRequest);
    expect(invalidResponse.status).toBe(400);
    expect(await invalidResponse.json()).toMatchObject({
      error: "invalid_request",
    });
  }
  const createResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      description: secretFormDescription,
      source: "blank",
      title: secretFormTitle,
    })
  );
  expect(createResponse.status).toBe(200);
  const createdForm = (await createResponse.json()) as {
    form?: {
      activeDraftCount?: number;
      hasTemplateDraft?: boolean;
      publicId?: string;
      status?: string;
      submissionCount?: number;
      title?: string;
    };
  };
  const formPublicId = createdForm.form?.publicId;
  expect(createdForm.form).toMatchObject({
    activeDraftCount: 0,
    hasTemplateDraft: true,
    status: "draft",
    submissionCount: 0,
    title: secretFormTitle,
  });
  if (!formPublicId) {
    throw new Error("The test form did not receive a public identifier");
  }
  expect(formPublicId).toMatch(/^[0-9a-f]{32}$/u);
  const createdFormRecord = await prisma.form.findUnique({
    include: { templateDraft: true },
    where: { publicId: formPublicId },
  });
  if (!createdFormRecord?.templateDraft) {
    throw new Error("The test form did not receive a template draft");
  }
  const formId = createdFormRecord.id;
  const templateDocumentKey = createdFormRecord.templateDraft.documentKey;
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: createdFormRecord.templateDraft.objectKey },
    })
  ).toBeNull();
  const serializedCreate = JSON.stringify(createdForm);
  expect(serializedCreate).not.toContain(formId);
  expect(serializedCreate).not.toContain(adminId);
  expect(serializedCreate).not.toContain(
    createdFormRecord.templateDraft.objectKey
  );
  expect(serializedCreate).not.toContain(templateDocumentKey);

  const listResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(listResponse.status).toBe(200);
  const listBody = (await listResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      status: string;
      submissionCount: number;
      title: string;
    }[];
  };
  const listedForm = listBody.forms?.find(
    (form) => form.publicId === formPublicId
  );
  expect(listedForm).toMatchObject({
    activeDraftCount: 0,
    publicId: formPublicId,
    status: "draft",
    submissionCount: 0,
    title: secretFormTitle,
  });
  for (const privateField of [
    "createdBy",
    "id",
    "objectKey",
    "publishedDocumentKey",
    "templateDocumentKey",
  ]) {
    expect(listedForm).not.toHaveProperty(privateField);
  }
  const serializedList = JSON.stringify(listBody);
  expect(serializedList).not.toContain(formId);
  expect(serializedList).not.toContain(adminId);
  expect(serializedList).not.toContain(
    createdFormRecord.templateDraft.objectKey
  );
  expect(serializedList).not.toContain(templateDocumentKey);
  const detailResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(detailResponse.status).toBe(200);
  const detailBody = (await detailResponse.json()) as {
    editorConfigUrl?: string;
    form?: { publicId?: string };
  };
  expect(detailBody).toMatchObject({
    editorConfigUrl: `/api/admin/forms/${formPublicId}/editor-config`,
    form: { publicId: formPublicId },
  });
  for (const privateField of [
    "createdBy",
    "id",
    "objectKey",
    "publishedDocumentKey",
    "templateDocumentKey",
  ]) {
    expect(detailBody.form).not.toHaveProperty(privateField);
  }
  const serializedDetail = JSON.stringify(detailBody);
  expect(serializedDetail).not.toContain(formId);
  expect(serializedDetail).not.toContain(adminId);
  expect(serializedDetail).not.toContain(
    createdFormRecord.templateDraft.objectKey
  );
  expect(serializedDetail).not.toContain(templateDocumentKey);
  const competingAdminEmail = `ticket-08-competing-${crypto.randomUUID()}@example.com`;
  const competingAdmin = await createCredentialFixture({
    email: competingAdminEmail,
    name: "Ticket 08 Competing Admin",
    password,
    role: "admin",
  });
  const competingAdminBearer = await bearerFor(
    competingAdminEmail,
    password,
    `ticket-08-competing-admin-${crypto.randomUUID()}`
  );
  const formCountBeforeUploadChecks = await prisma.form.count();
  const uploadBytes = docxFixture(`ticket-08-upload-${crypto.randomUUID()}`);
  const uploadTitle = `Ticket 08 upload ${crypto.randomUUID()}`;
  const uploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: { bytes: uploadBytes, name: "template.docx" },
      title: uploadTitle,
    })
  );
  expect(uploadResponse.status).toBe(200);
  const uploadBody = (await uploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const uploadPublicId = uploadBody.form?.publicId;
  const uploadRecord = uploadPublicId
    ? await prisma.form.findUnique({
        include: { templateDraft: true },
        where: { publicId: uploadPublicId },
      })
    : null;
  if (!uploadPublicId || !uploadRecord?.templateDraft) {
    throw new Error("The uploaded Template Draft was not created");
  }
  const uploadObjectKey = uploadRecord.templateDraft.objectKey;
  const uploadEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${uploadPublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(uploadEditorResponse.status).toBe(200);
  const uploadEditor = (await uploadEditorResponse.json()) as EditorConfigBody;
  const uploadLeaseId = uploadEditor.bridge.lease.id;
  expect(await readObject(uploadObjectKey)).toEqual(uploadBytes);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: uploadObjectKey },
    })
  ).toBeNull();
  const sameAdminOtherBearer = await bearerFor(
    adminEmail,
    password,
    `ticket-08-same-admin-other-session-${crypto.randomUUID()}`
  );
  const sameAdminOtherSessionDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${uploadPublicId}`, {
      headers: { Authorization: `Bearer ${sameAdminOtherBearer}` },
      method: "DELETE",
    })
  );
  expect(sameAdminOtherSessionDeleteResponse.status).toBe(409);
  expect(await sameAdminOtherSessionDeleteResponse.json()).toMatchObject({
    error: "editor_in_use",
  });
  const competingDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${uploadPublicId}`, {
      headers: { Authorization: `Bearer ${competingAdminBearer}` },
      method: "DELETE",
    })
  );
  expect(competingDeleteResponse.status).toBe(409);
  expect(await competingDeleteResponse.json()).toMatchObject({
    error: "editor_in_use",
  });
  expect(
    await prisma.form.findUnique({ where: { publicId: uploadPublicId } })
  ).not.toBeNull();
  expect(await objectExists(uploadObjectKey)).toBe(true);

  const maximumUploadBytes = sizedDocxFixture(
    `ticket-08-limit-${crypto.randomUUID()}`,
    maxTemplateUploadBytes
  );
  const maximumUploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: maximumUploadBytes,
        name: "maximum-size.docx",
        type: "application/octet-stream",
      },
      title: `Ticket 08 maximum upload ${crypto.randomUUID()}`,
    })
  );
  expect(maximumUploadResponse.status).toBe(200);
  const maximumUploadBody = (await maximumUploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const maximumUploadPublicId = maximumUploadBody.form?.publicId;
  expect(maximumUploadPublicId).toBeTruthy();

  const oversizedUploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: new Uint8Array(maxTemplateUploadBytes + 1),
        name: "too-large.docx",
      },
      title: "Ticket 08 oversized upload",
    })
  );
  expect(oversizedUploadResponse.status).toBe(413);
  expect(await oversizedUploadResponse.json()).toMatchObject({
    error: "payload_too_large",
  });
  const nonDocxResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxFixture("wrong-extension"),
        name: "template.pdf",
      },
      title: "Ticket 08 wrong upload type",
    })
  );
  expect(nonDocxResponse.status).toBe(415);
  expect(await nonDocxResponse.json()).toMatchObject({
    error: "invalid_file_type",
  });
  const malformedDocxResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: new TextEncoder().encode("not a ZIP package"),
        name: "malformed.docx",
      },
      title: "Ticket 08 malformed upload",
    })
  );
  expect(malformedDocxResponse.status).toBe(422);
  expect(await malformedDocxResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const incompleteDocxResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: zipSync({
          "word/document.xml": strToU8(
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
          ),
        }),
        name: "incomplete.docx",
      },
      title: "Ticket 08 incomplete upload",
    })
  );
  expect(incompleteDocxResponse.status).toBe(422);
  expect(await incompleteDocxResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const externalRelationshipResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          additionalParts: {
            "word/_rels/document.xml.rels": strToU8(
              '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="external" Target="http://127.0.0.1:80/" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/></Relationships>'
            ),
          },
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
        }),
        name: "external-relationship.docx",
      },
      title: "Ticket 10 external relationship upload",
    })
  );
  expect(externalRelationshipResponse.status).toBe(422);
  expect(await externalRelationshipResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const malformedXmlResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document><',
        }),
        name: "malformed-xml.docx",
      },
      title: "Ticket 08 malformed XML upload",
    })
  );
  expect(malformedXmlResponse.status).toBe(422);
  expect(await malformedXmlResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const doctypeResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          document:
            '<?xml version="1.0"?><!DOCTYPE word:document [<!ENTITY injected "value">]><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body><word:p>&injected;</word:p></word:body></word:document>',
        }),
        name: "doctype.docx",
      },
      title: "Ticket 08 XML doctype upload",
    })
  );
  expect(doctypeResponse.status).toBe(422);
  expect(await doctypeResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const declaredSecondaryXmlResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          additionalParts: {
            "word/_rels/document.xml.rels": strToU8(
              '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Target="header.bin" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"/></Relationships>'
            ),
            "word/header.bin": strToU8(
              '<!DOCTYPE w:hdr [<!ENTITY injected "value">]><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">&injected;</w:hdr>'
            ),
          },
          contentTypes: templateContentTypesXml.replace(
            "</Types>",
            '<Override PartName="/word/header.bin" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
          ),
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
        }),
        name: "declared-secondary-xml.docx",
      },
      title: "Ticket 08 declared secondary XML upload",
    })
  );
  expect(declaredSecondaryXmlResponse.status).toBe(422);
  expect(await declaredSecondaryXmlResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  const strictUploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: strictDocxFixture(`ticket-08-strict-${crypto.randomUUID()}`),
        name: "strict.docx",
      },
      title: "Ticket 08 Strict OOXML upload",
    })
  );
  const vmlDoctypeResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: docxXmlFixture({
          additionalParts: {
            "word/_rels/document.xml.rels": strToU8(
              '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="vml" Target="drawings/vmlDrawing1.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/></Relationships>'
            ),
            "word/drawings/vmlDrawing1.vml": strToU8(
              '<!DOCTYPE xml [<!ENTITY injected "value">]><xml xmlns:v="urn:schemas-microsoft-com:vml">&injected;</xml>'
            ),
          },
          contentTypes: templateContentTypesXml.replace(
            "</Types>",
            '<Override PartName="/word/drawings/vmlDrawing1.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>'
          ),
          document:
            '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
        }),
        name: "vml-doctype.docx",
      },
      title: "Ticket 08 VML doctype upload",
    })
  );
  expect(vmlDoctypeResponse.status).toBe(422);
  expect(await vmlDoctypeResponse.json()).toMatchObject({
    error: "invalid_template",
  });
  expect(strictUploadResponse.status).toBe(200);
  const strictUploadBody = (await strictUploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const strictUploadPublicId = strictUploadBody.form?.publicId;
  if (!strictUploadPublicId) {
    throw new Error("The Strict OOXML Template Draft was not created");
  }
  const strictDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${strictUploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(strictDeleteResponse.status).toBe(200);
  const utf16UploadResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: utf16DocxFixture(),
        name: "utf16.docx",
      },
      title: "Ticket 08 UTF-16 OOXML upload",
    })
  );
  expect(utf16UploadResponse.status).toBe(200);
  const utf16UploadBody = (await utf16UploadResponse.json()) as {
    form?: { publicId?: string };
  };
  const utf16UploadPublicId = utf16UploadBody.form?.publicId;
  if (!utf16UploadPublicId) {
    throw new Error("The UTF-16 OOXML Template Draft was not created");
  }
  const utf16DeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${utf16UploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(utf16DeleteResponse.status).toBe(200);
  expect(await prisma.form.count()).toBe(formCountBeforeUploadChecks + 2);

  const deleteUploadResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${uploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(deleteUploadResponse.status).toBe(200);
  expect(await deleteUploadResponse.json()).toEqual({ deleted: true });
  expect(
    await prisma.form.findUnique({ where: { publicId: uploadPublicId } })
  ).toBeNull();
  expect(
    await prisma.editorLease.findUnique({ where: { id: uploadLeaseId } })
  ).toBeNull();
  expect(await objectExists(uploadObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: uploadObjectKey },
    })
  ).toBeNull();

  if (!maximumUploadPublicId) {
    throw new Error("The maximum-sized Template Draft was not created");
  }
  const maximumUploadRecord = await prisma.form.findUnique({
    include: { templateDraft: true },
    where: { publicId: maximumUploadPublicId },
  });
  if (!maximumUploadRecord?.templateDraft) {
    throw new Error("The maximum-sized Template Draft was not persisted");
  }
  const maximumUploadObjectKey = maximumUploadRecord.templateDraft.objectKey;
  const deleteMaximumUploadResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${maximumUploadPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(deleteMaximumUploadResponse.status).toBe(200);
  expect(await objectExists(maximumUploadObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: maximumUploadObjectKey },
    })
  ).toBeNull();

  const adminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(adminEditorResponse.status).toBe(200);
  const adminEditor = (await adminEditorResponse.json()) as EditorConfigBody;
  const adminPluginOptions =
    adminEditor.config.editorConfig.plugins.options[pluginGuid];
  const adminLease = adminEditor.bridge.lease;
  const publishCapability = adminEditor.bridge.capabilities.publish;
  const configureFieldsCapability =
    adminEditor.bridge.capabilities["configure-fields"];
  let saveTemplateCapability = adminEditor.bridge.capabilities["save-template"];
  if (
    !adminPluginOptions ||
    !publishCapability ||
    !configureFieldsCapability ||
    !saveTemplateCapability
  ) {
    throw new Error("The Admin editor capabilities were not returned");
  }
  const adminEditorSerialized = JSON.stringify(adminEditor);
  expect(adminEditorSerialized).not.toContain(adminBearer);
  expect(adminEditorSerialized).not.toContain('"authToken"');
  expect(JSON.stringify(adminEditor.config)).not.toContain(publishCapability);
  expect(adminPluginOptions.publicId).toBe(formPublicId);
  expect(adminPluginOptions).not.toHaveProperty("formId");
  expect(adminEditorSerialized).not.toContain(formId);
  expect(adminEditor.bridge.id).toBe(adminPluginOptions.bridgeId);
  expect(adminEditor.bridge.pluginOrigin).toBe(
    new URL(process.env.API_BASE ?? "http://localhost:3000").origin
  );
  expect(adminPluginOptions.parentOrigin).toBe(
    new URL(process.env.CORS_ORIGIN ?? "http://localhost:5173").origin
  );
  expect(adminEditor.apiUrl).toBe(onlyOfficeBaseUrl);
  expect(adminEditor.config.editorConfig.callbackUrl).toBe(
    `${documentBaseUrl}/onlyoffice/callback`
  );
  expect(adminEditor.config.editorConfig.plugins.pluginsData).toEqual([
    `${apiBaseUrl}/onlyoffice-plugin/config.json`,
  ]);
  const pluginRootResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/")
  );
  expect(pluginRootResponse.status).toBe(200);
  expect(pluginRootResponse.headers.get("content-type")).toBe(
    "text/html; charset=utf-8"
  );
  const pluginHtmlResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/index.html")
  );
  expect(pluginHtmlResponse.status).toBe(200);
  expect(pluginHtmlResponse.headers.get("content-type")).toBe(
    "text/html; charset=utf-8"
  );
  const pluginHtml = await pluginHtmlResponse.text();
  const expectedPluginSdkUrl = `${onlyOfficeBaseUrl}/sdkjs-plugins/v1/plugins.js`;
  const escapedPluginSdkUrl = expectedPluginSdkUrl
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
  expect(pluginHtml).toContain(
    `<script src="${escapedPluginSdkUrl}"></script>`
  );

  expect(adminEditor.config.token.length).toBeGreaterThan(20);
  expect(adminLease).toEqual({
    expiresAt: expect.any(String),
    id: expect.any(String),
    releaseUrl: expect.any(String),
    renewUrl: expect.any(String),
  });
  for (const leaseValue of Object.values(adminLease)) {
    expect(JSON.stringify(adminEditor.config)).not.toContain(leaseValue);
  }
  expect(adminPluginOptions).not.toHaveProperty("lease");
  const adminLeaseRow = await prisma.editorLease.findUnique({
    where: { id: adminLease.id },
  });
  if (!adminLeaseRow) {
    throw new Error("The Admin editor lease was not persisted");
  }
  const configureClaims = verifyEditorCapability(configureFieldsCapability);
  expect(configureClaims).toMatchObject({
    action: "configure-fields",
    actorId: adminId,
    documentKey: templateDocumentKey,
    formId,
    leaseId: adminLease.id,
    targetId: createdFormRecord.templateDraft.id,
    targetType: "template-draft",
  });
  const schemaWithoutCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/schema`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(schemaWithoutCapabilityResponse.status).toBe(401);
  expect(await schemaWithoutCapabilityResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const configureHeaders = {
    ...jsonHeaders,
    "X-Editor-Capability": configureFieldsCapability,
  };
  const schemaFirstResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/schema`, {
      headers: configureHeaders,
    })
  );
  expect(schemaFirstResponse.status).toBe(200);
  const schemaFirst = (await schemaFirstResponse.json()) as {
    items: { pointer: string; type: string }[];
    nextCursor: string | null;
  };
  expect(schemaFirst.items).toHaveLength(5);
  expect(schemaFirst.nextCursor).toEqual(expect.any(String));
  expect(
    schemaFirst.items.every(
      (item) =>
        item.type === "string" ||
        item.type === "number" ||
        item.type === "boolean" ||
        item.type === "null"
    )
  ).toBe(true);
  expect(
    schemaFirst.items.some((item) => item.pointer.includes("/contacts/"))
  ).toBe(false);
  const schemaSecondResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?cursor=${encodeURIComponent(schemaFirst.nextCursor ?? "")}`,
      { headers: configureHeaders }
    )
  );
  expect(schemaSecondResponse.status).toBe(200);
  const schemaSecond = (await schemaSecondResponse.json()) as {
    items: { pointer: string; type: string }[];
    nextCursor: string | null;
  };
  expect(schemaSecond.items.length).toBeGreaterThan(0);
  expect(new Set(schemaSecond.items.map((item) => item.pointer)).size).toBe(
    schemaSecond.items.length
  );
  const schemaFilteredResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?q=${encodeURIComponent("ADDRESS")}`,
      { headers: configureHeaders }
    )
  );
  expect(schemaFilteredResponse.status).toBe(200);
  const schemaFiltered = (await schemaFilteredResponse.json()) as {
    items: { pointer: string; type: string }[];
    nextCursor: string | null;
  };
  expect(schemaFiltered.items.length).toBeGreaterThan(0);
  expect(
    schemaFiltered.items.every((item) =>
      item.pointer.toLowerCase().includes("address")
    )
  ).toBe(true);
  const escapedSchemaResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?q=${encodeURIComponent("display")}`,
      { headers: configureHeaders }
    )
  );
  expect(escapedSchemaResponse.status).toBe(200);
  expect(await escapedSchemaResponse.json()).toMatchObject({
    items: [{ pointer: "/account/display~1name", type: "string" }],
  });
  const invalidSchemaCursorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/schema?cursor=invalid`,
      { headers: configureHeaders }
    )
  );
  expect(invalidSchemaCursorResponse.status).toBe(400);
  expect(await invalidSchemaCursorResponse.json()).toMatchObject({
    error: "invalid_schema_cursor",
  });
  const emptyRulesResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        headers: configureHeaders,
      }
    )
  );
  expect(emptyRulesResponse.status).toBe(200);
  expect(await emptyRulesResponse.json()).toEqual({ rules: [] });
  const selectedPointer = schemaFirst.items.find(
    (item) => item.type === "string"
  )?.pointer;
  if (!selectedPointer) {
    throw new Error("Schema did not return a selectable pointer");
  }
  const createRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "lock-when-available",
          previousTag: null,
          required: true,
          tag: "ticket-09-field",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(createRuleResponse.status).toBe(200);
  expect(await createRuleResponse.json()).toEqual({
    rule: {
      prefillPointer: selectedPointer,
      prefillPolicy: "lock-when-available",
      required: true,
      tag: "ticket-09-field",
    },
  });
  const conflictingPointerResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "editable",
          previousTag: null,
          required: false,
          tag: "ticket-09-conflict",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(conflictingPointerResponse.status).toBe(409);
  expect(await conflictingPointerResponse.json()).toMatchObject({
    error: "field_rule_conflict",
  });
  const renamedRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "editable",
          previousTag: "ticket-09-field",
          required: false,
          tag: "ticket-09-renamed",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(renamedRuleResponse.status).toBe(200);
  expect(await renamedRuleResponse.json()).toEqual({
    rule: {
      prefillPointer: selectedPointer,
      prefillPolicy: "editable",
      required: false,
      tag: "ticket-09-renamed",
    },
  });
  const reopenedRulesResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        headers: configureHeaders,
      }
    )
  );
  expect(await reopenedRulesResponse.json()).toEqual({
    rules: [
      {
        prefillPointer: selectedPointer,
        prefillPolicy: "editable",
        required: false,
        tag: "ticket-09-renamed",
      },
    ],
  });
  const invalidPolicyResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: null,
          prefillPolicy: "lock-when-available",
          previousTag: "ticket-09-renamed",
          required: false,
          tag: "ticket-09-renamed",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(invalidPolicyResponse.status).toBe(400);
  expect(await invalidPolicyResponse.json()).toMatchObject({
    error: "invalid_field_config",
  });
  const staleDocumentResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: "template-stale",
          prefillPointer: selectedPointer,
          prefillPolicy: "editable",
          previousTag: "ticket-09-renamed",
          required: false,
          tag: "ticket-09-renamed",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(staleDocumentResponse.status).toBe(409);
  expect(await staleDocumentResponse.json()).toMatchObject({
    error: "stale_document",
  });
  const restoredRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: templateDocumentKey,
          prefillPointer: selectedPointer,
          prefillPolicy: "lock-when-available",
          previousTag: "ticket-09-renamed",
          required: true,
          tag: "full_name",
        }),
        headers: configureHeaders,
        method: "PATCH",
      }
    )
  );
  expect(restoredRuleResponse.status).toBe(200);
  expect(await restoredRuleResponse.json()).toEqual({
    rule: {
      prefillPointer: selectedPointer,
      prefillPolicy: "lock-when-available",
      required: true,
      tag: "full_name",
    },
  });
  for (const requiredTag of ["accept_terms", "department", "start_date"]) {
    const requiredRuleResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/forms/${formPublicId}/field-rules`,
        {
          body: JSON.stringify({
            documentKey: templateDocumentKey,
            prefillPointer: null,
            prefillPolicy: "editable",
            previousTag: null,
            required: true,
            tag: requiredTag,
          }),
          headers: configureHeaders,
          method: "PATCH",
        }
      )
    );
    expect(requiredRuleResponse.status).toBe(200);
  }

  expect(
    Math.abs(
      adminLeaseRow.expiresAt.getTime() -
        adminLeaseRow.createdAt.getTime() -
        90 * 1000
    )
  ).toBeLessThanOrEqual(1000);

  const repeatedAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(repeatedAdminEditorResponse.status).toBe(200);
  const repeatedAdminEditor =
    (await repeatedAdminEditorResponse.json()) as EditorConfigBody;
  expect(repeatedAdminEditor.bridge.lease.id).toBe(adminLease.id);

  const competingAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${competingAdminBearer}` },
      }
    )
  );
  expect(competingAdminEditorResponse.status).toBe(409);
  expect(await competingAdminEditorResponse.json()).toMatchObject({
    error: "editor_in_use",
  });

  const renewResponse = await app.handle(
    new Request(new URL(adminLease.renewUrl, "http://test.local").toString(), {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(renewResponse.status).toBe(200);
  const renewedBody = (await renewResponse.json()) as {
    lease?: { expiresAt?: string; id?: string };
  };
  expect(renewedBody).toMatchObject({
    lease: {
      expiresAt: expect.any(String),
      id: adminLease.id,
    },
  });
  const renewedLeaseRow = await prisma.editorLease.findUnique({
    where: { id: adminLease.id },
  });
  if (!renewedLeaseRow || !renewedBody.lease?.expiresAt) {
    throw new Error("The Admin editor lease was not renewed");
  }
  expect(renewedLeaseRow.renewedAt.getTime()).toBeGreaterThanOrEqual(
    renewedLeaseRow.createdAt.getTime()
  );
  expect(
    Math.abs(
      renewedLeaseRow.expiresAt.getTime() -
        renewedLeaseRow.renewedAt.getTime() -
        90 * 1000
    )
  ).toBeLessThanOrEqual(1000);

  const publishClaims = verifyEditorCapability(publishCapability);
  expect(publishClaims).toMatchObject({
    action: "publish",
    actorId: adminId,
    documentKey: templateDocumentKey,
    formId,
    leaseId: adminLease.id,
    leaseProof: expect.any(String),
    role: "admin",
    targetType: "template-draft",
  });
  expect((publishClaims?.expiresAt ?? 0) - (publishClaims?.issuedAt ?? 0)).toBe(
    5 * 60
  );

  const sessionOnlyAdminSaveResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyAdminSaveResponse.status).toBe(401);
  expect(await sessionOnlyAdminSaveResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const sessionOnlyAdminPublishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyAdminPublishResponse.status).toBe(401);
  expect(await sessionOnlyAdminPublishResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });

  const capabilityHeaders = (capability: string): Record<string, string> => ({
    ...jsonHeaders,
    "X-Editor-Capability": capability,
  });
  const crossActionResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(publishCapability),
      method: "POST",
    })
  );
  expect(crossActionResponse.status).toBe(403);

  const [capabilityHeader, capabilityPayload, capabilitySignature] =
    publishCapability.split(".");
  if (!capabilityHeader || !capabilityPayload || !capabilitySignature) {
    throw new Error("The Admin editor capability was malformed");
  }
  const tamperedPublishCapability = [
    capabilityHeader,
    capabilityPayload,
    `${capabilitySignature[0] === "a" ? "b" : "a"}${capabilitySignature.slice(1)}`,
  ].join(".");
  const tamperedCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(tamperedPublishCapability),
      method: "POST",
    })
  );
  expect(tamperedCapabilityResponse.status).toBe(401);

  if (!publishClaims) {
    throw new Error("The Admin publish capability did not verify");
  }
  const expiredPublishCapability = createEditorCapability({
    action: publishClaims.action,
    actorId: publishClaims.actorId,
    documentKey: publishClaims.documentKey,
    expiresAt: Math.floor(Date.now() / 1000) - 1,
    formId: publishClaims.formId,
    leaseId: publishClaims.leaseId,
    leaseProof: publishClaims.leaseProof,
    role: publishClaims.role,
    targetId: publishClaims.targetId,
    targetType: publishClaims.targetType,
  });
  const expiredCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(expiredPublishCapability),
      method: "POST",
    })
  );
  expect(expiredCapabilityResponse.status).toBe(401);

  const secondCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "blank",
      title: "Capability scope target",
    })
  );
  expect(secondCreateResponse.status).toBe(200);
  const secondCreatedForm = (await secondCreateResponse.json()) as {
    form?: { publicId?: string };
  };
  const secondFormPublicId = secondCreatedForm.form?.publicId;
  const secondFormRecord = secondFormPublicId
    ? await prisma.form.findUnique({
        include: { templateDraft: true },
        where: { publicId: secondFormPublicId },
      })
    : null;
  if (!secondFormPublicId || !secondFormRecord?.templateDraft) {
    throw new Error("The cross-target Form was not created");
  }
  const crossTargetResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${secondFormPublicId}/publish`,
      {
        body: JSON.stringify({
          documentKey: secondFormRecord.templateDraft.documentKey,
        }),
        headers: capabilityHeaders(publishCapability),
        method: "POST",
      }
    )
  );
  expect(crossTargetResponse.status).toBe(403);

  const documentUrl = adminEditor.config.document.url;
  const parsedDocumentUrl = new URL(documentUrl);
  expect(parsedDocumentUrl.origin).toBe(new URL(documentBaseUrl).origin);
  expect(parsedDocumentUrl.pathname).toBe(
    `/onlyoffice/document/${encodeURIComponent(templateDocumentKey)}`
  );
  const unsignedDocumentResponse = await app.handle(new Request(documentUrl));
  expect(unsignedDocumentResponse.status).toBe(401);
  const alteredDocumentUrl = new URL(documentUrl);
  alteredDocumentUrl.searchParams.set(
    "token",
    `${alteredDocumentUrl.searchParams.get("token") ?? ""}x`
  );
  const alteredDocumentResponse = await app.handle(
    new Request(alteredDocumentUrl.toString(), {
      headers: {
        Authorization: createOnlyOfficeAuthorization({
          url: alteredDocumentUrl.toString(),
        }),
      },
    })
  );
  expect(alteredDocumentResponse.status).toBe(401);
  const expiredDocumentUrl = new URL(documentUrl);
  expiredDocumentUrl.searchParams.set(
    "token",
    createDocumentAccessToken(
      templateDocumentKey,
      Math.floor(Date.now() / 1000) - 1
    )
  );
  const expiredDocumentResponse = await app.handle(
    new Request(expiredDocumentUrl.toString(), {
      headers: {
        Authorization: createOnlyOfficeAuthorization({
          url: expiredDocumentUrl.toString(),
        }),
      },
    })
  );
  expect(expiredDocumentResponse.status).toBe(401);
  const signedDocumentResponse = await app.handle(
    new Request(documentUrl, {
      headers: {
        Authorization: createOnlyOfficeAuthorization({ url: documentUrl }),
      },
    })
  );
  expect(signedDocumentResponse.status).toBe(200);
  expect(signedDocumentResponse.headers.get("content-type")).toBe(
    DOCX_CONTENT_TYPE
  );
  const initialTemplateBytes = new Uint8Array(
    await signedDocumentResponse.arrayBuffer()
  );
  expect(initialTemplateBytes.byteLength).toBeGreaterThan(0);

  const forbiddenPluginOriginResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/config.json", {
      headers: { Origin: "https://attacker.example" },
    })
  );
  expect(forbiddenPluginOriginResponse.status).toBe(403);
  const sameOriginPluginConfigResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/config.json")
  );
  expect(sameOriginPluginConfigResponse.status).toBe(200);
  expect(
    sameOriginPluginConfigResponse.headers.get("access-control-allow-origin")
  ).toBeNull();
  const allowedPluginOrigin = new URL(
    process.env.ONLYOFFICE_URL ?? "http://localhost:8080"
  ).origin;
  const pluginConfigResponse = await app.handle(
    new Request("http://test.local/onlyoffice-plugin/config.json", {
      headers: { Origin: allowedPluginOrigin },
    })
  );
  expect(pluginConfigResponse.status).toBe(200);
  expect(pluginConfigResponse.headers.get("access-control-allow-origin")).toBe(
    allowedPluginOrigin
  );
  const releaseResponse = await app.handle(
    new Request(
      new URL(adminLease.releaseUrl, "http://test.local").toString(),
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(releaseResponse.status).toBe(200);
  expect(await releaseResponse.json()).toEqual({ ok: true });
  expect(
    await prisma.editorLease.findUnique({ where: { id: adminLease.id } })
  ).toBeNull();

  const competingAdminEditorAfterReleaseResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${competingAdminBearer}` },
      }
    )
  );
  expect(competingAdminEditorAfterReleaseResponse.status).toBe(200);
  const competingAdminEditor =
    (await competingAdminEditorAfterReleaseResponse.json()) as EditorConfigBody;
  const competingLease = competingAdminEditor.bridge.lease;
  const competingSaveTemplateCapability =
    competingAdminEditor.bridge.capabilities["save-template"];
  if (!competingSaveTemplateCapability) {
    throw new Error("The competing Admin editor capability was not returned");
  }
  expect(competingLease.id).toBeTruthy();
  const releasedCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(releasedCapabilityResponse.status).toBe(409);
  expect(await releasedCapabilityResponse.json()).toMatchObject({
    error: "editor_lease_inactive",
  });

  await prisma.editorLease.update({
    data: { createdAt: new Date(0), expiresAt: new Date(1) },
    where: { id: competingLease.id },
  });
  const reclaimedAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(reclaimedAdminEditorResponse.status).toBe(200);
  const reclaimedAdminEditor =
    (await reclaimedAdminEditorResponse.json()) as EditorConfigBody;
  expect(reclaimedAdminEditor.bridge.lease.id).not.toBe(competingLease.id);
  const reclaimedSaveTemplateCapability =
    reclaimedAdminEditor.bridge.capabilities["save-template"];
  if (!reclaimedSaveTemplateCapability) {
    throw new Error("The reclaimed Admin editor capability was not returned");
  }
  const expiredCompetingCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(competingSaveTemplateCapability),
      method: "POST",
    })
  );
  expect(expiredCompetingCapabilityResponse.status).toBe(409);
  expect(await expiredCompetingCapabilityResponse.json()).toMatchObject({
    error: "editor_lease_inactive",
  });
  saveTemplateCapability = reclaimedSaveTemplateCapability;
  const operationCountBeforeInvalidSaveBodies = await prisma.operation.count({
    where: { formId },
  });
  const extraSaveFieldResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({
        documentKey: templateDocumentKey,
        unexpected: true,
      }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(extraSaveFieldResponse.status).toBe(400);
  expect(await extraSaveFieldResponse.json()).toMatchObject({
    error: "invalid_request",
  });
  const oversizedSaveBodyResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({
        documentKey: templateDocumentKey,
        padding: "x".repeat(8192),
      }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(oversizedSaveBodyResponse.status).toBe(413);
  expect(await oversizedSaveBodyResponse.json()).toMatchObject({
    error: "payload_too_large",
  });
  expect(await prisma.operation.count({ where: { formId } })).toBe(
    operationCountBeforeInvalidSaveBodies
  );
  expect(await readObject(createdFormRecord.templateDraft.objectKey)).toEqual(
    initialTemplateBytes
  );

  const saveTemplateResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: templateDocumentKey }),
      headers: capabilityHeaders(saveTemplateCapability),
      method: "POST",
    })
  );
  expect(saveTemplateResponse.status).toBe(202);
  const saveTemplateBody = (await saveTemplateResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!saveTemplateBody.operationCapability || !saveTemplateBody.operationId) {
    throw new Error("The template save operation was not created");
  }
  const saveTemplatePollClaims = verifyEditorCapability(
    saveTemplateBody.operationCapability
  );
  expect(saveTemplatePollClaims).toMatchObject({
    action: "poll-operation",
    operationId: saveTemplateBody.operationId,
  });
  expect(
    (saveTemplatePollClaims?.expiresAt ?? 0) -
      (saveTemplatePollClaims?.issuedAt ?? 0)
  ).toBe(6 * 60);
  const saveTemplateOperation = await waitForOperation(
    saveTemplateBody.operationId,
    {
      "X-Editor-Capability": saveTemplateBody.operationCapability,
    }
  );
  expect(saveTemplateOperation.status).toBe("completed");
  expect(saveTemplateOperation).toMatchObject({
    result: { publicId: formPublicId },
  });
  expect(JSON.stringify(saveTemplateOperation)).not.toContain(formId);

  const refreshedAdminEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(refreshedAdminEditorResponse.status).toBe(200);
  const refreshedAdminEditor =
    (await refreshedAdminEditorResponse.json()) as EditorConfigBody;
  const activeTemplateDocumentKey = refreshedAdminEditor.config.document.key;
  const activePublishCapability =
    refreshedAdminEditor.bridge.capabilities.publish;
  if (!activeTemplateDocumentKey || !activePublishCapability) {
    throw new Error("The refreshed Admin editor capability was not returned");
  }
  expect(activeTemplateDocumentKey).not.toBe(templateDocumentKey);
  const reopenedDocumentUrl = refreshedAdminEditor.config.document.url;
  const reopenedDocumentResponse = await app.handle(
    new Request(reopenedDocumentUrl, {
      headers: {
        Authorization: createOnlyOfficeAuthorization({
          url: reopenedDocumentUrl,
        }),
      },
    })
  );
  expect(reopenedDocumentResponse.status).toBe(200);
  expect(new Uint8Array(await reopenedDocumentResponse.arrayBuffer())).toEqual(
    initialTemplateBytes
  );
  const publishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: activeTemplateDocumentKey }),
      headers: capabilityHeaders(activePublishCapability),
      method: "POST",
    })
  );
  expect(publishResponse.status).toBe(202);
  const publishBody = (await publishResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!publishBody.operationCapability || !publishBody.operationId) {
    throw new Error("The publish operation was not created");
  }
  const publishPollClaims = verifyEditorCapability(
    publishBody.operationCapability
  );
  expect(publishPollClaims).toMatchObject({
    action: "poll-operation",
    operationId: publishBody.operationId,
  });
  expect(
    (publishPollClaims?.expiresAt ?? 0) - (publishPollClaims?.issuedAt ?? 0)
  ).toBe(6 * 60);
  const publishOperation = await waitForOperation(publishBody.operationId, {
    "X-Editor-Capability": publishBody.operationCapability,
  });
  expect(publishOperation.status).toBe("completed");
  expect(publishOperation).toMatchObject({
    result: { publicId: formPublicId, version: 1 },
  });
  expect(JSON.stringify(publishOperation)).not.toContain(formId);
  const publishedListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(publishedListResponse.status).toBe(200);
  const publishedList = (await publishedListResponse.json()) as {
    forms?: { publicId: string; status: string; version: number }[];
  };
  expect(
    publishedList.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({ status: "published", version: 1 });
  const publishedManifestRecord = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: { include: { fields: true } },
    },
    where: { formId },
  });
  if (!publishedManifestRecord?.manifest) {
    throw new Error("The published manifest was not persisted");
  }
  const publishedManifest = publishedManifestRecord.manifest;
  const publishedBytes = await readObject(publishedManifestRecord.objectKey);
  const expectedPublishedHash = createHash("sha256")
    .update(publishedBytes)
    .digest("hex");
  expect(publishedManifestRecord).toMatchObject({
    contentHash: expectedPublishedHash,
    version: 1,
  });
  expect(publishedManifestRecord.manifest).toMatchObject({
    configurationHash: expectedPublishedHash,
    displayMetadataVersion: 1,
  });
  expect(publishedManifestRecord.manifest.fields).toMatchObject([
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: true,
      tag: "accept_terms",
      type: "checkbox",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: [
        { displayText: "Choose an item", value: "" },
        { displayText: "Engineering", value: "engineering" },
        { displayText: "Human Resources", value: "hr" },
        { displayText: "Finance", value: "finance" },
      ],
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: true,
      tag: "department",
      type: "dropdown",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: false,
      tag: "description_1",
      type: "text",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: false,
      tag: "description_2",
      type: "text",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "lock_when_available",
      required: true,
      tag: "full_name",
      type: "text",
    },
    {
      id: expect.any(String),
      manifestId: expect.any(String),
      options: null,
      pictureMaxBytes: null,
      pictureMaxHeight: null,
      pictureMaxWidth: null,
      prefillPolicy: "editable",
      required: true,
      tag: "start_date",
      type: "date",
    },
  ]);
  expect(publishedManifestRecord.prefillConfiguration?.fields).toEqual([
    {
      configurationId: expect.any(String),
      id: expect.any(String),
      pointer: selectedPointer,
      policy: "lock_when_available",
      tag: "full_name",
    },
  ]);
  expect(publishedManifestRecord.prefillConfiguration).toMatchObject({
    configurationHash: expectedPublishedHash,
    publishedTemplateId: publishedManifestRecord.id,
  });
  const publishFixture = async (label: string, bytes: Uint8Array) => {
    const createFixtureResponse = await app.handle(
      formCreationRequest({
        authorization: adminBearer,
        source: "upload",
        template: { bytes, name: `${label}.docx` },
        title: `Ticket 10 ${label}`,
      })
    );
    expect(createFixtureResponse.status).toBe(200);
    const fixtureBody = (await createFixtureResponse.json()) as {
      form?: { publicId?: string };
    };
    const fixturePublicId = fixtureBody.form?.publicId;
    if (!fixturePublicId) {
      throw new Error(`The ${label} fixture did not receive a public ID`);
    }
    const fixtureEditorResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/forms/${fixturePublicId}/editor-config`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    expect(fixtureEditorResponse.status).toBe(200);
    const fixtureEditor =
      (await fixtureEditorResponse.json()) as EditorConfigBody;
    const fixturePublishCapability = fixtureEditor.bridge.capabilities.publish;
    if (!fixturePublishCapability) {
      throw new Error(
        `The ${label} fixture publish capability was not returned`
      );
    }
    const response = await app.handle(
      new Request(
        `http://test.local/api/admin/forms/${fixturePublicId}/publish`,
        {
          body: JSON.stringify({
            documentKey: fixtureEditor.config.document.key,
          }),
          headers: capabilityHeaders(fixturePublishCapability),
          method: "POST",
        }
      )
    );
    expect(response.status).toBe(202);
    const body = (await response.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
    if (!body.operationCapability || !body.operationId) {
      throw new Error(`The ${label} fixture operation was not created`);
    }
    const operation = await waitForOperation(body.operationId, {
      "X-Editor-Capability": body.operationCapability,
    });
    return { operation, publicId: fixturePublicId };
  };
  const validExtendedFixture = docxXmlFixture({
    document: contentControlDocument(
      contentControl({
        alias: "Applicant photo",
        tag: "photo",
        type: "<w:picture/>",
      }) +
        contentControl({
          alias: "Applicant name",
          placeholderText: "Enter applicant name",
          tag: "/person/name",
          type: "<w:text/>",
        }) +
        contentControl({
          alias: "   ",
          placeholderText: "Choose a department",
          tag: "department_choice",
          type: `<w:comboBox><w:listItem w:displayText="Engineering" w:value="engineering"/><w:listItem w:displayText="Finance" w:value="finance"/></w:comboBox>`,
        })
    ),
  });
  const validExtendedResult = await publishFixture(
    "combo-and-picture",
    validExtendedFixture
  );
  expect(validExtendedResult.operation.status).toBe("completed");
  const validExtendedForm = await prisma.form.findUniqueOrThrow({
    select: { id: true },
    where: { publicId: validExtendedResult.publicId },
  });
  const validExtendedPublished =
    await prisma.publishedTemplate.findUniqueOrThrow({
      include: {
        manifest: { include: { fields: { orderBy: { position: "asc" } } } },
      },
      where: { formId: validExtendedForm.id },
    });
  if (!validExtendedPublished.manifest) {
    throw new Error("The extended fixture manifest was not persisted");
  }
  const validExtendedFields = validExtendedPublished.manifest.fields;
  expect(validExtendedFields).toHaveLength(3);
  expect(
    validExtendedFields.map(({ label, placeholder, position, tag }) => ({
      label,
      placeholder,
      position,
      tag,
    }))
  ).toEqual([
    {
      label: "Applicant photo",
      placeholder: null,
      position: 0,
      tag: "photo",
    },
    {
      label: "Applicant name",
      placeholder: "Enter applicant name",
      position: 1,
      tag: "/person/name",
    },
    {
      label: "department_choice",
      placeholder: "Choose a department",
      position: 2,
      tag: "department_choice",
    },
  ]);
  expect(
    validExtendedFields.find((field) => field.tag === "/person/name")
  ).toMatchObject({
    options: null,
    pictureMaxBytes: null,
    pictureMaxHeight: null,
    pictureMaxWidth: null,
    tag: "/person/name",
    type: "text",
  });
  expect(
    validExtendedFields.find((field) => field.tag === "department_choice")
  ).toMatchObject({
    options: [
      { displayText: "Engineering", value: "engineering" },
      { displayText: "Finance", value: "finance" },
    ],
    type: "combo",
  });
  expect(
    validExtendedFields.find((field) => field.tag === "photo")
  ).toMatchObject({
    options: null,
    pictureMaxBytes: 10 * 1024 * 1024,
    pictureMaxHeight: 4096,
    pictureMaxWidth: 4096,
    type: "picture",
  });
  const orderedPlaceholderFixture = docxXmlFixture({
    additionalParts: {
      "word/_rels/document.xml.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdGlossary" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>`
      ),
      "word/glossary/document.xml": strToU8(
        `<?xml version="1.0"?><w:glossaryDocument xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:docParts><w:docPart><w:docPartPr><w:docPartName w:val="ReceiptPrompt"/></w:docPartPr><w:docPartBody><w:p><w:r><w:t>Configured glossary placeholder</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`
      ),
    },
    document: contentControlDocument(
      `<w:sdt><w:sdtPr><w:alias w:val="Parent label"/><w:tag w:val="parent"/><w:text/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:alias w:val="Child label"/><w:tag w:val="child"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Child answer</w:t></w:r></w:sdtContent></w:sdt></w:sdtContent></w:sdt>` +
        `<w:sdt><w:sdtPr><w:alias w:val="Prompt label"/><w:tag w:val="prompt"/><w:placeholder><w:docPart w:val="ReceiptPrompt"/></w:placeholder><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Entered answer</w:t></w:r></w:sdtContent></w:sdt>`
    ).replaceAll(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "http://purl.oclc.org/ooxml/wordprocessingml/main"
    ),
  });
  const orderedPlaceholderResult = await publishFixture(
    "nested-and-glossary-placeholder",
    orderedPlaceholderFixture
  );
  const orderedPlaceholderForm = await prisma.form.findUniqueOrThrow({
    select: { id: true },
    where: { publicId: orderedPlaceholderResult.publicId },
  });
  const orderedPlaceholderPublished =
    await prisma.publishedTemplate.findUniqueOrThrow({
      include: { manifest: { include: { fields: true } } },
      where: { formId: orderedPlaceholderForm.id },
    });
  expect(
    orderedPlaceholderPublished.manifest?.fields
      .toSorted((left, right) => left.position - right.position)
      .map(({ label, placeholder, position, tag }) => ({
        label,
        placeholder,
        position,
        tag,
      }))
  ).toEqual([
    { label: "Parent label", placeholder: null, position: 0, tag: "parent" },
    { label: "Child label", placeholder: null, position: 1, tag: "child" },
    {
      label: "Prompt label",
      placeholder: "Configured glossary placeholder",
      position: 2,
      tag: "prompt",
    },
  ]);
  const requiredPictureCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      source: "upload",
      template: {
        bytes: pictureDocumentFixture(),
        name: "required-picture.docx",
      },
      title: "Ticket 16 required picture",
    })
  );
  expect(requiredPictureCreateResponse.status).toBe(200);
  const requiredPictureCreateBody =
    (await requiredPictureCreateResponse.json()) as {
      form?: { publicId?: string };
    };
  const requiredPicturePublicId = requiredPictureCreateBody.form?.publicId;
  if (!requiredPicturePublicId) {
    throw new Error("The required picture form was not created");
  }
  const requiredPictureEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(requiredPictureEditorResponse.status).toBe(200);
  const requiredPictureEditor =
    (await requiredPictureEditorResponse.json()) as EditorConfigBody;
  const configurePictureCapability =
    requiredPictureEditor.bridge.capabilities["configure-fields"];
  if (!configurePictureCapability) {
    throw new Error(
      "The required picture configure capability was not returned"
    );
  }
  const requiredPictureRuleResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/field-rules`,
      {
        body: JSON.stringify({
          documentKey: requiredPictureEditor.config.document.key,
          prefillPointer: null,
          prefillPolicy: "editable",
          previousTag: null,
          required: true,
          tag: "photo",
        }),
        headers: capabilityHeaders(configurePictureCapability),
        method: "PATCH",
      }
    )
  );
  expect(requiredPictureRuleResponse.status).toBe(200);
  const requiredPicturePublishEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  const requiredPicturePublishEditor =
    (await requiredPicturePublishEditorResponse.json()) as EditorConfigBody;
  const requiredPicturePublishCapability =
    requiredPicturePublishEditor.bridge.capabilities.publish;
  if (!requiredPicturePublishCapability) {
    throw new Error("The required picture publish capability was not returned");
  }
  const requiredPicturePublishResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${requiredPicturePublicId}/publish`,
      {
        body: JSON.stringify({
          documentKey: requiredPicturePublishEditor.config.document.key,
        }),
        headers: capabilityHeaders(requiredPicturePublishCapability),
        method: "POST",
      }
    )
  );
  expect(requiredPicturePublishResponse.status).toBe(202);
  const requiredPicturePublishBody =
    (await requiredPicturePublishResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !requiredPicturePublishBody.operationCapability ||
    !requiredPicturePublishBody.operationId
  ) {
    throw new Error("The required picture publish operation was not created");
  }
  const requiredPicturePublishOperation = await waitForOperation(
    requiredPicturePublishBody.operationId,
    {
      "X-Editor-Capability": requiredPicturePublishBody.operationCapability,
    }
  );
  expect(requiredPicturePublishOperation.status).toBe("completed");
  const requiredPictureForm = await prisma.form.findUniqueOrThrow({
    select: { id: true },
    where: { publicId: requiredPicturePublicId },
  });
  const requiredPictureManifest =
    await prisma.publishedTemplate.findUniqueOrThrow({
      include: { manifest: { include: { fields: true } } },
      where: { formId: requiredPictureForm.id },
    });
  expect(requiredPictureManifest.manifest?.fields).toContainEqual(
    expect.objectContaining({
      pictureMaxBytes: 10 * 1024 * 1024,
      pictureMaxHeight: 4096,
      pictureMaxWidth: 4096,
      required: true,
      tag: "photo",
      type: "picture",
    })
  );
  let nextPictureDocument = pictureDocumentFixture({
    images: [{ bytes: pngFixture(), extension: "png" }],
    includeStaticImage: true,
  });
  const pictureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.resolve(new TextEncoder().encode("%PDF-picture")),
      forceSave: async (documentKey) => {
        const response = await prisma.response.findUnique({
          select: { draftObjectKey: true },
          where: { draftDocumentKey: documentKey },
        });
        if (!response?.draftObjectKey) {
          throw new Error("The picture response document was not found");
        }
        await putObject(
          response.draftObjectKey,
          nextPictureDocument,
          DOCX_CONTENT_TYPE
        );
        return false;
      },
    },
  });
  const pictureStartResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/forms/${requiredPicturePublicId}/start`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      }
    )
  );
  expect(pictureStartResponse.status).toBe(200);
  const pictureStartBody = (await pictureStartResponse.json()) as {
    response?: { id?: string };
  };
  const pictureResponseId = pictureStartBody.response?.id;
  if (!pictureResponseId) {
    throw new Error("The picture response was not started");
  }
  const pictureEditorResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/forms/${requiredPicturePublicId}/editor-config?responseId=${pictureResponseId}&action=draft`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(pictureEditorResponse.status).toBe(200);
  const pictureEditor =
    (await pictureEditorResponse.json()) as EditorConfigBody;
  let pictureDocumentKey = pictureEditor.config.document.key;
  let pictureSaveCapability: string =
    pictureEditor.bridge.capabilities["save-draft"] ?? "";
  let pictureSubmitCapability: string =
    pictureEditor.bridge.capabilities.submit ?? "";
  if (!pictureSaveCapability || !pictureSubmitCapability) {
    throw new Error("The picture response capabilities were not returned");
  }
  const refreshPictureEditor = async (): Promise<void> => {
    const refreshedResponse = await pictureApp.handle(
      new Request(
        `http://test.local/api/forms/${requiredPicturePublicId}/editor-config?responseId=${pictureResponseId}&action=draft`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    expect(refreshedResponse.status).toBe(200);
    const refreshedConfig =
      (await refreshedResponse.json()) as EditorConfigBody;
    const refreshedSaveCapability =
      refreshedConfig.bridge.capabilities["save-draft"];
    const refreshedSubmitCapability =
      refreshedConfig.bridge.capabilities.submit;
    if (!refreshedSaveCapability || !refreshedSubmitCapability) {
      throw new Error(
        "The refreshed picture response capabilities were not returned"
      );
    }
    pictureDocumentKey = refreshedConfig.config.document.key;
    pictureSaveCapability = refreshedSaveCapability;
    pictureSubmitCapability = refreshedSubmitCapability;
  };
  const pictureDraftRequest = (data: Record<string, unknown>) =>
    pictureApp.handle(
      new Request(
        `http://test.local/api/forms/${requiredPicturePublicId}/draft`,
        {
          body: JSON.stringify({
            data,
            documentKey: pictureDocumentKey,
            responseId: pictureResponseId,
          }),
          headers: capabilityHeaders(pictureSaveCapability),
          method: "POST",
        }
      )
    );
  const pictureSubmitRequest = (data: Record<string, unknown>) =>
    pictureApp.handle(
      new Request(
        `http://test.local/api/forms/${requiredPicturePublicId}/submit`,
        {
          body: JSON.stringify({
            data,
            documentKey: pictureDocumentKey,
            responseId: pictureResponseId,
          }),
          headers: capabilityHeaders(pictureSubmitCapability),
          method: "POST",
        }
      )
    );
  const savePicture = async (
    document: Uint8Array,
    data: Record<string, unknown>
  ) => {
    nextPictureDocument = document;
    const response = await pictureDraftRequest(data);
    expect(response.status).toBe(202);
    const body = (await response.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
    if (!body.operationCapability || !body.operationId) {
      throw new Error("The picture draft operation was not created");
    }
    const operation = await waitForOperation(body.operationId, {
      "X-Editor-Capability": body.operationCapability,
    });
    if (operation.status === "completed") {
      await refreshPictureEditor();
    }
    return operation;
  };
  const validPictureDocument = nextPictureDocument;
  const validPictureSave = await savePicture(validPictureDocument, {
    photo: "scalar data must be omitted",
  });
  expect(validPictureSave.status).toBe("completed");
  const savedPictureResponse = await prisma.response.findUniqueOrThrow({
    select: { draftData: true, draftObjectKey: true, status: true },
    where: { id: pictureResponseId },
  });
  expect(savedPictureResponse).toMatchObject({
    draftData: {},
    status: "draft",
  });
  if (!savedPictureResponse.draftObjectKey) {
    throw new Error("The saved picture draft object was not persisted");
  }
  expect(await readObject(savedPictureResponse.draftObjectKey)).toEqual(
    validPictureDocument
  );
  const resumedPictureStartResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/forms/${requiredPicturePublicId}/start`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      }
    )
  );
  expect(await resumedPictureStartResponse.json()).toMatchObject({
    response: { id: pictureResponseId, status: "draft" },
  });
  const pictureDocxExport = await pictureApp.handle(
    new Request(
      `http://test.local/api/responses/${pictureResponseId}/draft/docx`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(new Uint8Array(await pictureDocxExport.arrayBuffer())).toEqual(
    Uint8Array.from(validPictureDocument)
  );
  const picturePdfExport = await pictureApp.handle(
    new Request(
      `http://test.local/api/responses/${pictureResponseId}/draft/pdf`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(await picturePdfExport.text()).toBe("%PDF-picture");
  const missingPictureDraft = await savePicture(pictureDocumentFixture(), {});
  expect(missingPictureDraft.status).toBe("completed");
  expect(
    await prisma.submission.count({ where: { responseId: pictureResponseId } })
  ).toBe(0);
  const invalidPictureDocuments = [
    pictureDocumentFixture({
      images: [{ bytes: gifFixture, extension: "gif" }],
    }),
    pictureDocumentFixture({
      images: [
        { bytes: pngFixture(), extension: "png" },
        { bytes: jpegFixture(), extension: "jpg" },
      ],
    }),
    pictureDocumentFixture({
      images: [
        {
          bytes: pngFixture(1, 1, 10 * 1024 * 1024 + 1),
          extension: "png",
        },
      ],
    }),
    pictureDocumentFixture({
      images: [{ bytes: pngFixture(4097, 1), extension: "png" }],
    }),
  ];
  for (const invalidPictureDocument of invalidPictureDocuments) {
    const invalidPictureSave = await savePicture(invalidPictureDocument, {});
    expect(invalidPictureSave.status).toBe("failed");
    expect(
      await prisma.submission.count({
        where: { responseId: pictureResponseId },
      })
    ).toBe(0);
  }
  const placeholderPictureDocument = pictureDocumentFixture({
    images: [{ bytes: pngFixture(), extension: "png" }],
    showingPlaceholder: true,
  });
  const placeholderPictureDraft = await savePicture(
    placeholderPictureDocument,
    {}
  );
  expect(placeholderPictureDraft.status).toBe("completed");
  nextPictureDocument = placeholderPictureDocument;
  const placeholderPictureSubmitResponse = await pictureSubmitRequest({});
  expect(placeholderPictureSubmitResponse.status).toBe(202);
  const placeholderPictureSubmitBody =
    (await placeholderPictureSubmitResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !placeholderPictureSubmitBody.operationCapability ||
    !placeholderPictureSubmitBody.operationId
  ) {
    throw new Error("The placeholder picture submit operation was not created");
  }
  const placeholderPictureSubmitOperation = await waitForOperation(
    placeholderPictureSubmitBody.operationId,
    { "X-Editor-Capability": placeholderPictureSubmitBody.operationCapability }
  );
  expect(placeholderPictureSubmitOperation).toMatchObject({
    error: "invalid_template",
    status: "failed",
  });
  expect(
    await prisma.submission.count({ where: { responseId: pictureResponseId } })
  ).toBe(0);
  nextPictureDocument = pictureDocumentFixture();
  const missingPictureSubmitResponse = await pictureSubmitRequest({});
  expect(missingPictureSubmitResponse.status).toBe(202);
  const missingPictureSubmitBody =
    (await missingPictureSubmitResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !missingPictureSubmitBody.operationCapability ||
    !missingPictureSubmitBody.operationId
  ) {
    throw new Error("The missing picture submit operation was not created");
  }
  const missingPictureSubmitOperation = await waitForOperation(
    missingPictureSubmitBody.operationId,
    { "X-Editor-Capability": missingPictureSubmitBody.operationCapability }
  );
  expect(missingPictureSubmitOperation).toMatchObject({
    error: "invalid_template",
    status: "failed",
  });
  expect(
    await prisma.response.findUnique({
      select: { status: true },
      where: { id: pictureResponseId },
    })
  ).toMatchObject({ status: "draft" });
  expect(
    await prisma.submission.count({ where: { responseId: pictureResponseId } })
  ).toBe(0);
  nextPictureDocument = pictureDocumentFixture({
    images: [{ bytes: jpegFixture(), extension: "jpg" }],
  });
  const pictureSubmitResponse = await pictureSubmitRequest({
    photo: "scalar data must be omitted",
  });
  expect(pictureSubmitResponse.status).toBe(202);
  const pictureSubmitBody = (await pictureSubmitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !pictureSubmitBody.operationCapability ||
    !pictureSubmitBody.operationId
  ) {
    throw new Error("The picture submit operation was not created");
  }
  const pictureSubmitOperation = await waitForOperation(
    pictureSubmitBody.operationId,
    { "X-Editor-Capability": pictureSubmitBody.operationCapability }
  );
  expect(pictureSubmitOperation.status).toBe("completed");
  const pictureSubmission = await prisma.submission.findUniqueOrThrow({
    select: { data: true, objectKey: true, responseId: true },
    where: { responseId: pictureResponseId },
  });
  expect(pictureSubmission.data).toEqual({});
  expect(await readObject(pictureSubmission.objectKey)).toEqual(
    nextPictureDocument
  );
  const pictureCorrectionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${pictureResponseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(pictureCorrectionEditorResponse.status).toBe(200);
  const pictureCorrectionEditor =
    (await pictureCorrectionEditorResponse.json()) as EditorConfigBody;
  const pictureCorrectionCapability =
    pictureCorrectionEditor.bridge.capabilities["save-correction"];
  if (!pictureCorrectionCapability) {
    throw new Error("The picture correction capability was not returned");
  }
  const pictureCorrectionResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${pictureResponseId}/correction`,
      {
        body: JSON.stringify({
          data: {},
          documentKey: pictureCorrectionEditor.config.document.key,
          reason: "ตรวจสอบรูปภาพ",
        }),
        headers: capabilityHeaders(pictureCorrectionCapability),
        method: "POST",
      }
    )
  );
  expect(pictureCorrectionResponse.status).toBe(202);
  const pictureCorrectionBody = (await pictureCorrectionResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !pictureCorrectionBody.operationCapability ||
    !pictureCorrectionBody.operationId
  ) {
    throw new Error("The picture correction operation was not created");
  }
  const pictureCorrectionOperation = await waitForOperation(
    pictureCorrectionBody.operationId,
    {
      "X-Editor-Capability": pictureCorrectionBody.operationCapability,
    }
  );
  expect(pictureCorrectionOperation.status).toBe("completed");
  const pictureCorrection = await prisma.correction.findFirstOrThrow({
    select: { data: true, objectKey: true, revision: true },
    where: { responseId: pictureResponseId },
  });
  expect(pictureCorrection).toMatchObject({ data: {}, revision: 1 });
  expect(await readObject(pictureCorrection.objectKey)).toEqual(
    nextPictureDocument
  );
  await putObject(
    pictureCorrection.objectKey,
    pictureDocumentFixture(),
    DOCX_CONTENT_TYPE
  );
  const pictureRevisionsResponse = await pictureApp.handle(
    new Request(
      `http://test.local/api/responses/${pictureResponseId}/corrections`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(pictureRevisionsResponse.status).toBe(200);
  expect(await pictureRevisionsResponse.json()).toMatchObject({
    revisions: [
      { pictures: { photo: true }, revision: 0 },
      { pictures: { photo: false }, revision: 1 },
    ],
  });
  const invalidFixtureCases = [
    {
      bytes: docxFixture("no-controls"),
      label: "no-controls",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "duplicate", type: "<w:text/>" }) +
            contentControl({ tag: "duplicate", type: "<w:text/>" })
        ),
      }),
      label: "duplicate-tags",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({
            tag: "malformed-options",
            type: `<w:comboBox><w:listItem w:displayText="Missing value"/></w:comboBox>`,
          })
        ),
      }),
      label: "malformed-options",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "", type: "<w:text/>" })
        ),
      }),
      label: "blank-tag",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({
            tag: "duplicate-options",
            type: `<w:comboBox><w:listItem w:displayText="One" w:value="same"/><w:listItem w:displayText="Two" w:value="same"/></w:comboBox>`,
          })
        ),
      }),
      label: "duplicate-options",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "unsupported", type: "<w:group/>" })
        ),
      }),
      label: "unsupported-group",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag w:val="nested-marker"/><w:placeholder><w:text/></w:placeholder></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "nested-marker",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag w:val="nested-option"/><w:comboBox/><w:placeholder><w:listItem w:displayText="Wrong parent" w:value="wrong"/></w:placeholder></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "nested-option",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag w:val="nested-option-child"/><w:comboBox><w:listItem w:displayText="One" w:value="one"><w:bogus/></w:listItem></w:comboBox></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "nested-option-child",
    },
    {
      bytes: docxXmlFixture({
        additionalParts: {
          "word/header1.xml": strToU8(
            `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${contentControl({ tag: "orphan-header", type: "<w:text/>" })}</w:hdr>`
          ),
        },
        document:
          '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
      }),
      label: "orphan-header-control",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "unknown", type: "<w:unknown/>" })
        ),
      }),
      label: "unknown-control",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          contentControl({ tag: "wrong-namespace", type: "<w14:picture/>" })
        ),
      }),
      label: "wrong-namespace",
    },
    {
      bytes: docxXmlFixture({
        document: contentControlDocument(
          `<w:sdt><w:sdtPr><w:tag word:val="forged"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>fixture</w:t></w:r></w:sdtContent></w:sdt>`
        ),
      }),
      label: "wrong-attribute-namespace",
    },
  ];
  for (const fixture of invalidFixtureCases) {
    const result = await publishFixture(fixture.label, fixture.bytes);
    expect(result.operation.status).toBe("failed");
    const fixtureForm = await prisma.form.findUniqueOrThrow({
      select: { id: true },
      where: { publicId: result.publicId },
    });
    expect(
      await prisma.publishedTemplate.findUnique({
        where: { formId: fixtureForm.id },
      })
    ).toBeNull();
  }
  const publishedDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(publishedDeleteResponse.status).toBe(409);
  expect(await publishedDeleteResponse.json()).toMatchObject({
    error: "form_not_draft",
  });
  expect(
    await prisma.form.findUnique({ where: { publicId: formPublicId } })
  ).not.toBeNull();

  const immutableSaveCapability =
    refreshedAdminEditor.bridge.capabilities["save-template"];
  if (!immutableSaveCapability) {
    throw new Error("The immutable save capability was not returned");
  }
  const immutableSaveResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/save`, {
      body: JSON.stringify({ documentKey: activeTemplateDocumentKey }),
      headers: capabilityHeaders(immutableSaveCapability),
      method: "POST",
    })
  );
  expect(immutableSaveResponse.status).toBe(409);
  expect(await immutableSaveResponse.json()).toMatchObject({
    error: "published_immutable",
  });
  const immutablePublishResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/publish`, {
      body: JSON.stringify({ documentKey: activeTemplateDocumentKey }),
      headers: capabilityHeaders(activePublishCapability),
      method: "POST",
    })
  );
  expect(immutablePublishResponse.status).toBe(409);
  expect(await immutablePublishResponse.json()).toMatchObject({
    error: "published_immutable",
  });

  const user = await createCredentialFixture({
    email: userEmail,
    name: "Ticket 04 Workflow User",
    password,
  });
  const userId = user.id;
  const userBearer = await bearerFor(userEmail, password);

  const formRecord = await prisma.form.findUnique({
    select: { publicId: true },
    where: { id: formId },
  });
  if (!formRecord) {
    throw new Error("The published test form was not found");
  }
  const prefillHandoffSecret = process.env.PREFILL_HANDOFF_SECRET ?? "";
  const mock = startPrefillMock({
    connector: createInProcessFolioConnector(
      (request) => app.handle(request),
      prefillHandoffSecret
    ),
    folioOrigin: "http://test.local",
  });
  externalMock = mock;
  const schemaResponse = await fetch(`${mock.url}/schema?q=person`);
  expect(schemaResponse.status).toBe(200);
  const schemaBody = (await schemaResponse.json()) as {
    items?: { pointer: string; type: string }[];
  };
  expect(schemaBody.items).toContainEqual({
    pointer: "/person/name",
    type: "string",
  });
  const handoffExternalReference = `ticket-17-reference-${crypto.randomUUID()}`;
  const handoffCandidateValues = {
    ...deterministicExternalRecord,
    account: { ...deterministicExternalRecord.account, active: true },
    person: {
      ...deterministicExternalRecord.person,
      name: "Ticket 17 Prefill",
    },
    [selectedPointer]: selectedPointer.endsWith("/active")
      ? true
      : "Ticket 17 Prefill",
  };
  const ordinaryStartBeforeHandoff = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(ordinaryStartBeforeHandoff.status).toBe(409);
  expect(await ordinaryStartBeforeHandoff.json()).toMatchObject({
    error: "prefill_required",
  });
  const invalidHandoffSecret = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: userEmail,
        externalReference: handoffExternalReference,
        publicId: formRecord.publicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": "wrong-secret",
      },
      method: "POST",
    })
  );
  expect(invalidHandoffSecret.status).toBe(404);
  expect(await invalidHandoffSecret.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const handoffCreateResponse = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: ` ${userEmail.toUpperCase()} `,
      externalReference: handoffExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(handoffCreateResponse.status).toBe(200);
  const handoffCreateBody = (await handoffCreateResponse.json()) as {
    code?: string;
    launchPath?: string;
  };
  if (!handoffCreateBody.code) {
    throw new Error("The external mock did not return a handoff code");
  }
  expect(handoffCreateBody.code).toMatch(/^[A-Za-z0-9_-]{43}$/u);
  expect(handoffCreateBody.launchPath).toBe("/prefill/handoff");
  const duplicateHandoffCreateResponse = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: handoffExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(duplicateHandoffCreateResponse.status).toBe(409);
  expect(await duplicateHandoffCreateResponse.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const handoffRecord = await prisma.handoff.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(handoffExternalReference)
        .digest("hex"),
    },
  });
  expect(handoffRecord.codeDigest).toBe(
    createHash("sha256").update(handoffCreateBody.code).digest("hex")
  );
  expect(handoffRecord.codeDigest).not.toBe(handoffCreateBody.code);
  expect(handoffRecord.codeDigest).toHaveLength(64);
  expect(handoffRecord.expiresAt.getTime()).toBeGreaterThan(Date.now());
  expect(handoffRecord.expiresAt.getTime()).toBeLessThanOrEqual(
    Date.now() + 120_000 + 1000
  );
  expect(handoffRecord.filteredValues).toEqual({
    full_name: selectedPointer.endsWith("/active") ? true : "Ticket 17 Prefill",
  });
  const pendingStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: handoffExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(pendingStatusResponse.status).toBe(200);
  const pendingStatusBody = (await pendingStatusResponse.json()) as Record<
    string,
    unknown
  >;
  expect(Object.keys(pendingStatusBody).toSorted()).toEqual([
    "consumedAt",
    "createdAt",
    "expiresAt",
    "latestCorrectionNumber",
    "reservedAt",
    "status",
    "submittedAt",
    "updatedAt",
  ]);
  expect(pendingStatusBody).toMatchObject({
    latestCorrectionNumber: null,
    reservedAt: null,
    status: "pending",
    submittedAt: null,
  });
  expect(JSON.stringify(pendingStatusBody)).not.toContain(userEmail);
  expect(JSON.stringify(pendingStatusBody)).not.toContain("Ticket 17 Prefill");
  const unknownStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({
      externalReference: `ticket-19-unknown-${crypto.randomUUID()}`,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(unknownStatusResponse.status).toBe(404);
  expect(await unknownStatusResponse.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const invalidStatusSecretResponse = await app.handle(
    new Request("http://test.local/api/integrations/prefill/status", {
      body: JSON.stringify({ externalReference: handoffExternalReference }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": "wrong-secret",
      },
      method: "POST",
    })
  );
  expect(invalidStatusSecretResponse.status).toBe(404);
  expect(await invalidStatusSecretResponse.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const handoffReturnOverride = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: userEmail,
        externalReference: `ticket-19-return-override-${crypto.randomUUID()}`,
        publicId: formRecord.publicId,
        returnUrl: "https://attacker.example.test",
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(handoffReturnOverride.status).toBe(400);
  expect(await handoffReturnOverride.json()).toMatchObject({
    error: "invalid_request",
  });
  const getLaunchResponse = await app.handle(
    new Request(
      `http://test.local${handoffCreateBody.launchPath}?code=${handoffCreateBody.code}`
    )
  );
  expect(getLaunchResponse.status).toBe(405);
  const launchPageResponse = await fetch(`${mock.url}/launch`, {
    body: new URLSearchParams({ code: handoffCreateBody.code }),
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    method: "POST",
  });
  expect(launchPageResponse.status).toBe(200);
  const launchPage = await launchPageResponse.text();
  expect(launchPage).toContain('action="http://test.local/prefill/handoff"');
  expect(launchPage).not.toContain("?code=");
  const launchCode = /name="code" value="(?<code>[^"]+)"/u.exec(launchPage)
    ?.groups?.code;
  expect(launchCode).toBe(handoffCreateBody.code);
  if (!launchCode) {
    throw new Error("The external mock did not render a handoff code");
  }
  const launchResponse = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: launchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(launchResponse.status).toBe(303);
  expect(launchResponse.headers.get("location")).toBe(
    `/forms/${formRecord.publicId}/fill`
  );
  const pendingCookie = launchResponse.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!pendingCookie) {
    throw new Error("The handoff launch did not set a pending claim cookie");
  }
  const launchCookieHeader = launchResponse.headers.get("set-cookie") ?? "";
  expect(launchCookieHeader).toContain("__Host-folio-pending-claim=");
  expect(launchCookieHeader).toContain("Path=/");
  expect(launchCookieHeader).toContain("Max-Age=600");
  expect(launchCookieHeader).toContain("HttpOnly");
  expect(launchCookieHeader).toContain("Secure");
  expect(launchCookieHeader).toContain("SameSite=Lax");
  const redeemedStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: pendingCookie,
      },
      method: "POST",
    })
  );
  expect(redeemedStartResponse.status).toBe(200);
  const redeemedStartBody = (await redeemedStartResponse.json()) as {
    editorConfigUrl?: string;
    prefill?: {
      data?: Record<string, unknown>;
      editableFields?: Record<string, unknown>;
    };
    response?: { id?: string };
  };
  expect(redeemedStartBody.editorConfigUrl).toContain("action=fill");
  expect(redeemedStartBody.prefill?.data).toEqual({
    full_name: selectedPointer.endsWith("/active") ? true : "Ticket 17 Prefill",
  });
  expect(redeemedStartBody.prefill?.editableFields).toEqual({
    full_name: false,
  });
  const handoffResponseId = redeemedStartBody.response?.id;
  if (!handoffResponseId) {
    throw new Error("The handoff did not create a Response");
  }
  const redeemedSnapshot = await prisma.prefillSnapshot.findUniqueOrThrow({
    where: { responseId: handoffResponseId },
  });
  expect(redeemedSnapshot.values).toEqual({
    full_name: selectedPointer.endsWith("/active") ? true : "Ticket 17 Prefill",
  });
  expect(redeemedSnapshot.lockedFields).toEqual({ full_name: true });
  expect(
    await prisma.response.findUniqueOrThrow({
      select: { externalReferenceDigest: true },
      where: { id: handoffResponseId },
    })
  ).toEqual({
    externalReferenceDigest: createHash("sha256")
      .update(handoffExternalReference)
      .digest("hex"),
  });
  const replayedHandoffStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: pendingCookie,
      },
      method: "POST",
    })
  );
  expect(replayedHandoffStart.status).toBe(409);
  expect(await replayedHandoffStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const redeemedHandoffRecord = await prisma.handoff.findUniqueOrThrow({
    where: { id: handoffRecord.id },
  });
  expect(redeemedHandoffRecord.status).toBe("consumed");
  expect(redeemedHandoffRecord.responseId).toBe(handoffResponseId);
  const draftStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: handoffExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(draftStatusResponse.status).toBe(200);
  expect(await draftStatusResponse.json()).toMatchObject({
    latestCorrectionNumber: null,
    status: "draft",
    submittedAt: null,
  });
  const reentryExternalReference = `ticket-18-reentry-${crypto.randomUUID()}`;
  const reentryCreateResponse = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: reentryExternalReference,
      publicId: formRecord.publicId,
      values: {
        ...handoffCandidateValues,
        [selectedPointer]: "A newer external value",
      },
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(reentryCreateResponse.status).toBe(200);
  const reentryCode = ((await reentryCreateResponse.json()) as { code: string })
    .code;
  const reentryPage = await fetch(`${mock.url}/launch`, {
    body: new URLSearchParams({ code: reentryCode }),
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    method: "POST",
  });
  const reentryPageHtml = await reentryPage.text();
  const reentryLaunchCode = /name="code" value="(?<code>[^"]+)"/u.exec(
    reentryPageHtml
  )?.groups?.code;
  if (!reentryLaunchCode) {
    throw new Error("The reentry mock did not render a launch form");
  }
  const reentryLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: reentryLaunchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const reentryCookie = reentryLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!reentryCookie) {
    throw new Error("The reentry handoff did not set a pending claim cookie");
  }
  const reentryHandoff = await prisma.handoff.findFirstOrThrow({
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(reentryExternalReference)
        .digest("hex"),
    },
  });
  const reentryReservedAt = reentryHandoff.reservedAt;
  if (!reentryReservedAt) {
    throw new Error("The reentry handoff was not reserved");
  }
  const lateApp = createApp({
    clock: () => new Date(reentryReservedAt.getTime() + 121_000),
    prefillHandoffSecret,
  });
  const stableReentryStart = await lateApp.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: reentryCookie,
      },
      method: "POST",
    })
  );
  expect(stableReentryStart.status).toBe(200);
  expect(await stableReentryStart.json()).toMatchObject({
    prefill: {
      data: {
        full_name: selectedPointer.endsWith("/active")
          ? true
          : "Ticket 17 Prefill",
      },
      editableFields: { full_name: false },
    },
    response: { id: handoffResponseId },
  });
  const stableReentrySnapshot = await prisma.prefillSnapshot.findUniqueOrThrow({
    where: { responseId: handoffResponseId },
  });
  expect(stableReentrySnapshot).toMatchObject({
    lockedFields: { full_name: true },
    values: {
      full_name: selectedPointer.endsWith("/active")
        ? true
        : "Ticket 17 Prefill",
    },
  });
  expect(
    await prisma.response.findUniqueOrThrow({
      select: { externalReferenceDigest: true },
      where: { id: handoffResponseId },
    })
  ).toEqual({
    externalReferenceDigest: createHash("sha256")
      .update(handoffExternalReference)
      .digest("hex"),
  });
  const missingValueEmail = `ticket-18-missing-${crypto.randomUUID()}@example.com`;
  const missingValueUser = await createCredentialFixture({
    email: missingValueEmail,
    name: "Ticket 18 Missing Value User",
    password,
  });
  const missingValueBearer = await bearerFor(missingValueEmail, password);
  const missingValueReference = `ticket-18-missing-${crypto.randomUUID()}`;
  const missingValueCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: missingValueEmail,
      externalReference: missingValueReference,
      publicId: formRecord.publicId,
      values: { account: { active: true } },
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(missingValueCreate.status).toBe(200);
  const missingValueCode = (
    (await missingValueCreate.json()) as { code: string }
  ).code;
  const missingValueLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: missingValueCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const missingValueCookie = missingValueLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!missingValueCookie) {
    throw new Error("The missing-value handoff did not set a cookie");
  }
  const missingValueStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${missingValueBearer}`,
        Cookie: missingValueCookie,
      },
      method: "POST",
    })
  );
  expect(missingValueStart.status).toBe(200);
  expect(await missingValueStart.json()).toMatchObject({
    prefill: { data: {}, editableFields: {} },
  });
  const missingValueResponse = await prisma.response.findUniqueOrThrow({
    include: { prefillSnapshot: true },
    where: {
      formId_userId: { formId, userId: missingValueUser.id },
    },
  });
  expect(missingValueResponse.prefillSnapshot).toMatchObject({
    lockedFields: {},
    values: {},
  });
  const discardedMissingValue = await app.handle(
    new Request(`http://test.local/api/responses/${missingValueResponse.id}`, {
      headers: { Authorization: `Bearer ${missingValueBearer}` },
      method: "DELETE",
    })
  );
  expect(discardedMissingValue.status).toBe(200);
  const deletedStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: missingValueReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(deletedStatusResponse.status).toBe(200);
  expect(await deletedStatusResponse.json()).toMatchObject({
    status: "deleted",
    submittedAt: null,
  });
  const malformedValueCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: ` ${missingValueEmail.toUpperCase()} `,
      externalReference: `ticket-18-malformed-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: {
        account: {
          address: { city: 42 },
        },
      },
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(malformedValueCreate.status).toBe(409);
  expect(await malformedValueCreate.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const largePrefillEmail = `ticket-18-large-${crypto.randomUUID()}@example.com`;
  const handoffCountBeforeLargeValue = await prisma.handoff.count();
  const responseCountBeforeLargeValue = await prisma.response.count({
    where: { formId },
  });
  const largePrefillCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: largePrefillEmail,
      externalReference: `ticket-18-large-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: {
        ...handoffCandidateValues,
        [selectedPointer]: "x".repeat(10_001),
      },
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(largePrefillCreate.status).toBe(409);
  expect(await largePrefillCreate.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  expect(await prisma.handoff.count()).toBe(handoffCountBeforeLargeValue);
  expect(
    await prisma.response.count({
      where: { formId },
    })
  ).toBe(responseCountBeforeLargeValue);
  const mandatoryEmail = `ticket-17-password-${crypto.randomUUID()}@example.com`;
  const mandatoryUser = await createCredentialFixture({
    email: mandatoryEmail,
    mustChangePassword: true,
    name: "Ticket 17 Password User",
    password,
  });
  const mandatoryHandoffCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: mandatoryEmail,
      externalReference: `ticket-17-password-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const mandatoryHandoffCode = (
    (await mandatoryHandoffCreate.json()) as { code: string }
  ).code;
  const mandatoryHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: mandatoryHandoffCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(mandatoryHandoffLaunch.status).toBe(303);
  const mandatoryCookie = mandatoryHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!mandatoryCookie) {
    throw new Error("The mandatory-password handoff did not set a cookie");
  }
  const mandatoryBearer = await bearerFor(mandatoryEmail, password);
  const passwordChangeResponse = await app.handle(
    new Request("http://test.local/api/account/password", {
      body: JSON.stringify({
        currentPassword: password,
        newPassword: "Ticket17-new-password",
      }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${mandatoryBearer}`,
        Cookie: mandatoryCookie,
      },
      method: "POST",
    })
  );
  expect(passwordChangeResponse.status).toBe(200);
  const freshMandatoryBearer = await bearerFor(
    mandatoryEmail,
    "Ticket17-new-password"
  );
  const passwordContinuation = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${freshMandatoryBearer}`,
        Cookie: mandatoryCookie,
      },
      method: "POST",
    })
  );
  expect(passwordContinuation.status).toBe(200);
  const passwordContinuationBody = await passwordContinuation.json();
  expect(passwordContinuationBody).toMatchObject({
    response: { id: expect.any(String) },
  });
  const passwordResponseRecord = await prisma.response.findFirstOrThrow({
    select: { id: true, userId: true },
    where: { formId, userId: mandatoryUser.id },
  });
  expect(passwordResponseRecord.userId).toBe(mandatoryUser.id);
  const emailMismatchCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: `ticket-17-email-mismatch-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const emailMismatchCode = (
    (await emailMismatchCreate.json()) as {
      code: string;
    }
  ).code;
  const emailMismatchLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: emailMismatchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const emailMismatchCookie = emailMismatchLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!emailMismatchCookie) {
    throw new Error("The email-mismatch handoff did not set a cookie");
  }
  const emailMismatchStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${missingValueBearer}`,
        Cookie: emailMismatchCookie,
      },
      method: "POST",
    })
  );
  expect(emailMismatchStart.status).toBe(409);
  expect(await emailMismatchStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });

  const formMismatchCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: `ticket-17-form-mismatch-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const formMismatchCode = (
    (await formMismatchCreate.json()) as {
      code: string;
    }
  ).code;
  const formMismatchLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: formMismatchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const formMismatchCookie = formMismatchLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!formMismatchCookie) {
    throw new Error("The Form-mismatch handoff did not set a cookie");
  }
  const formMismatchStart = await app.handle(
    new Request(`http://test.local/api/forms/${secondFormPublicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: formMismatchCookie,
      },
      method: "POST",
    })
  );
  expect(formMismatchStart.status).toBe(409);
  expect(await formMismatchStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });

  const configMismatchCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: mandatoryEmail,
      externalReference: `ticket-17-config-mismatch-${crypto.randomUUID()}`,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const configMismatchCode = (
    (await configMismatchCreate.json()) as {
      code: string;
    }
  ).code;
  const configMismatchLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: configMismatchCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const configMismatchCookie = configMismatchLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!configMismatchCookie) {
    throw new Error("The config-mismatch handoff did not set a cookie");
  }
  const configMismatchHandoff = await prisma.handoff.findFirstOrThrow({
    where: {
      codeDigest: createHash("sha256").update(configMismatchCode).digest("hex"),
    },
  });
  await prisma.handoff.update({
    data: { configurationHash: "0".repeat(64) },
    where: { id: configMismatchHandoff.id },
  });
  const configMismatchStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${freshMandatoryBearer}`,
        Cookie: configMismatchCookie,
      },
      method: "POST",
    })
  );
  expect(configMismatchStart.status).toBe(409);
  expect(await configMismatchStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  const discardedMandatory = await app.handle(
    new Request(
      `http://test.local/api/responses/${passwordResponseRecord.id}`,
      {
        headers: { Authorization: `Bearer ${freshMandatoryBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(discardedMandatory.status).toBe(200);

  const expiredExternalReference = `ticket-17-expired-${crypto.randomUUID()}`;
  const expiredCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: expiredExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  const expiredCode = ((await expiredCreate.json()) as { code: string }).code;
  const expiredRecord = await prisma.handoff.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      codeDigest: createHash("sha256").update(expiredCode).digest("hex"),
    },
  });
  await prisma.handoff.update({
    data: {
      expiresAt: new Date(expiredRecord.createdAt.getTime() + 1),
    },
    where: { id: expiredRecord.id },
  });
  const expiredApp = createApp({
    clock: () => new Date(expiredRecord.createdAt.getTime() + 2),
    prefillHandoffSecret,
  });
  const expiredLaunch = await expiredApp.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: expiredCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(expiredLaunch.status).toBe(303);
  expect(expiredLaunch.headers.get("location")).toBe(
    "/handoff?error=handoff_unavailable"
  );
  await reconcileRecoverableState();
  const purgedExpiredRecord = await prisma.handoff.findUniqueOrThrow({
    select: {
      configurationHash: true,
      consumedAt: true,
      createdAt: true,
      expiresAt: true,
      externalReferenceDigest: true,
      filteredValues: true,
      formId: true,
      id: true,
      normalizedEmail: true,
      reservedAt: true,
      responseId: true,
      status: true,
      updatedAt: true,
    },
    where: { id: expiredRecord.id },
  });
  expect(purgedExpiredRecord).toEqual({
    configurationHash: null,
    consumedAt: null,
    createdAt: expiredRecord.createdAt,
    expiresAt: new Date(expiredRecord.createdAt.getTime() + 1),
    externalReferenceDigest: createHash("sha256")
      .update(expiredExternalReference)
      .digest("hex"),
    filteredValues: null,
    formId: null,
    id: expiredRecord.id,
    normalizedEmail: null,
    reservedAt: null,
    responseId: null,
    status: "expired",
    updatedAt: expect.any(Date),
  });
  expect(
    await prisma.pendingClaim.count({ where: { handoffId: expiredRecord.id } })
  ).toBe(0);
  const expiredStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: expiredExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(expiredStatusResponse.status).toBe(200);
  expect(await expiredStatusResponse.json()).toMatchObject({
    status: "expired",
    submittedAt: null,
  });
  const raceExternalReference = `ticket-19-expiry-race-${crypto.randomUUID()}`;
  const raceCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: raceExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(raceCreate.status).toBe(200);
  const raceCode = ((await raceCreate.json()) as { code: string }).code;
  const raceLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: raceCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  const raceCookie = raceLaunch.headers.get("set-cookie")?.split(";", 1)[0];
  if (!raceCookie) {
    throw new Error("The expiry-race handoff did not set a cookie");
  }
  const raceHandoff = await prisma.handoff.findFirstOrThrow({
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(raceExternalReference)
        .digest("hex"),
    },
  });
  await prisma.handoff.update({
    data: { expiresAt: new Date(Date.now() - 1) },
    where: { id: raceHandoff.id },
  });
  const [raceStatus, , raceRedeem] = await Promise.all([
    fetch(`${mock.url}/status`, {
      body: JSON.stringify({ externalReference: raceExternalReference }),
      headers: jsonHeaders,
      method: "POST",
    }),
    reconcileRecoverableState(),
    app.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
        headers: {
          Authorization: `Bearer ${userBearer}`,
          Cookie: raceCookie,
        },
        method: "POST",
      })
    ),
  ]);
  expect(raceStatus.status).toBe(200);
  expect(await raceStatus.json()).toMatchObject({ status: "expired" });
  expect(raceRedeem.status).toBe(409);
  const raceRedeemError = ((await raceRedeem.json()) as { error?: unknown })
    .error;
  if (typeof raceRedeemError !== "string") {
    throw new TypeError("The expiry-race redeem did not return an error code");
  }
  expect(["handoff_unavailable", "prefill_required"]).toContain(
    raceRedeemError
  );
  expect(
    await prisma.pendingClaim.count({ where: { handoffId: raceHandoff.id } })
  ).toBe(0);
  expect(
    await prisma.handoff.findUniqueOrThrow({ where: { id: raceHandoff.id } })
  ).toMatchObject({
    codeDigest: null,
    filteredValues: null,
    normalizedEmail: null,
    status: "expired",
  });
  const malformedLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: JSON.stringify({ code: handoffCreateBody.code }),
      headers: {
        "Content-Type": "application/json",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(malformedLaunch.status).toBe(303);
  expect(malformedLaunch.headers.get("location")).toBe(
    "/handoff?error=handoff_unavailable"
  );
  const nonNavigationLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: handoffCreateBody.code }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Mode": "cors",
      },
      method: "POST",
    })
  );
  expect(nonNavigationLaunch.status).toBe(303);
  expect(nonNavigationLaunch.headers.get("location")).toBe(
    "/handoff?error=handoff_unavailable"
  );
  const raceUserEmail = `ticket-17-race-${crypto.randomUUID()}@example.com`;
  await createCredentialFixture({
    email: raceUserEmail,
    name: "Ticket 17 Race User",
    password,
  });
  const raceUserBearer = await bearerFor(raceUserEmail, password);
  const raceHandoffReference = `ticket-17-race-${crypto.randomUUID()}`;
  const raceHandoffCreate = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: raceUserEmail,
        externalReference: raceHandoffReference,
        publicId: formRecord.publicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(raceHandoffCreate.status).toBe(200);
  const raceHandoffCode = ((await raceHandoffCreate.json()) as { code: string })
    .code;
  const raceHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: raceHandoffCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    })
  );
  expect(raceHandoffLaunch.status).toBe(303);
  const racePendingCookie = raceHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!racePendingCookie) {
    throw new Error("The race handoff did not set a pending claim cookie");
  }
  const raceStarts = await Promise.all(
    [0, 1].map(() =>
      app.handle(
        new Request(
          `http://test.local/api/forms/${formRecord.publicId}/start`,
          {
            headers: {
              Authorization: `Bearer ${raceUserBearer}`,
              Cookie: racePendingCookie,
            },
            method: "POST",
          }
        )
      )
    )
  );
  expect(raceStarts.map((response) => response.status).toSorted()).toEqual([
    200, 409,
  ]);
  const handoffRaceBodies = (await Promise.all(
    raceStarts.map((response) => response.json())
  )) as { response?: { id?: string } }[];
  const raceResponseId = handoffRaceBodies.find((body) => body.response?.id)
    ?.response?.id;
  if (!raceResponseId) {
    throw new Error("The handoff race did not create a Response");
  }
  const raceHandoffRecord = await prisma.handoff.findFirstOrThrow({
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(raceHandoffReference)
        .digest("hex"),
    },
  });
  expect(raceHandoffRecord.status).toBe("consumed");
  expect(raceHandoffRecord.responseId).toBe(raceResponseId);
  const raceDiscard = await app.handle(
    new Request(`http://test.local/api/responses/${raceResponseId}`, {
      headers: { Authorization: `Bearer ${raceUserBearer}` },
      method: "DELETE",
    })
  );
  expect(raceDiscard.status).toBe(200);
  const deletedRaceHandoff = await prisma.handoff.findUniqueOrThrow({
    where: { id: raceHandoffRecord.id },
  });
  expect(deletedRaceHandoff.status).toBe("deleted");
  const concurrentStarts = await Promise.all(
    [0, 1].map(() =>
      app.handle(
        new Request(
          `http://test.local/api/forms/${formRecord.publicId}/start`,
          {
            headers: { Authorization: `Bearer ${userBearer}` },
            method: "POST",
          }
        )
      )
    )
  );
  expect(concurrentStarts.every((response) => response.status === 200)).toBe(
    true
  );
  const concurrentStartBodies = (await Promise.all(
    concurrentStarts.map((response) => response.json())
  )) as { response?: { id?: string } }[];
  expect(concurrentStartBodies[0]).toMatchObject({
    response: { id: expect.any(String) },
  });
  expect(concurrentStartBodies[1]).toMatchObject({
    response: { id: concurrentStartBodies[0]?.response?.id },
  });
  expect(
    await prisma.response.count({ where: { formId, userId: user.id } })
  ).toBe(1);
  const existingResponseId = concurrentStartBodies[0]?.response?.id;
  if (!existingResponseId) {
    throw new Error("The concurrent response was not created");
  }
  const publishedContractBefore = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
    },
    where: { formId },
  });
  if (!publishedContractBefore?.manifest) {
    throw new Error("The Ticket 11 published contract was not found");
  }
  const unauthorizedMetadataUpdate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        description: "must not change",
        title: "must not change",
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "PATCH",
    })
  );
  expect(unauthorizedMetadataUpdate.status).toBe(403);
  const updatedTitle = `${secretFormTitle} metadata`;
  const updatedDescription = `${secretFormDescription} metadata`;
  const metadataUpdate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        description: updatedDescription,
        title: updatedTitle,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(metadataUpdate.status).toBe(200);
  expect(await metadataUpdate.json()).toMatchObject({
    form: {
      description: updatedDescription,
      publicId: formPublicId,
      status: "published",
      title: updatedTitle,
      version: 1,
    },
  });
  const publicMetadataAfterUpdate = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(await publicMetadataAfterUpdate.json()).toMatchObject({
    form: { description: updatedDescription, title: updatedTitle },
  });
  const publishedContractAfter = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
    },
    where: { formId },
  });
  if (!publishedContractAfter?.manifest) {
    throw new Error("The Ticket 11 published contract was removed");
  }
  expect(publishedContractAfter).toMatchObject({
    contentHash: publishedContractBefore.contentHash,
    documentKey: publishedContractBefore.documentKey,
    id: publishedContractBefore.id,
    objectKey: publishedContractBefore.objectKey,
    version: publishedContractBefore.version,
  });
  expect(publishedContractAfter.manifest).toMatchObject({
    configurationHash: publishedContractBefore.manifest.configurationHash,
    id: publishedContractBefore.manifest.id,
  });
  expect(publishedContractAfter.manifest.fields).toEqual(
    publishedContractBefore.manifest.fields
  );
  expect(publishedContractAfter.prefillConfiguration).toEqual(
    publishedContractBefore.prefillConfiguration
  );
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: { action: "update_form_metadata", targetId: formPublicId },
    })
  ).toMatchObject({ outcome: "success", targetId: formPublicId });
  const restoreMetadata = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        description: secretFormDescription,
        title: secretFormTitle,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(restoreMetadata.status).toBe(200);

  const unauthorizedDuplicate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/duplicate`, {
      body: "{}",
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(unauthorizedDuplicate.status).toBe(403);
  const duplicatePublishedResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}/duplicate`, {
      body: "{}",
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "POST",
    })
  );
  expect(duplicatePublishedResponse.status).toBe(200);
  const duplicatePublishedBody = (await duplicatePublishedResponse.json()) as {
    form?: {
      description?: string;
      publicId?: string;
      status?: string;
      title?: string;
      version?: number;
    };
  };
  const duplicatePublishedPublicId = duplicatePublishedBody.form?.publicId;
  if (!duplicatePublishedPublicId) {
    throw new Error("The published Form duplicate was not created");
  }
  expect(duplicatePublishedBody.form).toMatchObject({
    description: secretFormDescription,
    status: "draft",
    title: secretFormTitle,
    version: 0,
  });
  expect(duplicatePublishedPublicId).not.toBe(formPublicId);
  expect(duplicatePublishedPublicId).toMatch(/^[0-9a-f]{32}$/u);
  const duplicatePublished = await prisma.form.findUnique({
    include: {
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
      templateDraft: { include: { fieldRules: { orderBy: { tag: "asc" } } } },
    },
    where: { publicId: duplicatePublishedPublicId },
  });
  if (!duplicatePublished?.templateDraft) {
    throw new Error("The published Form duplicate has no Template Draft");
  }
  expect(
    Buffer.from(await readObject(duplicatePublished.templateDraft.objectKey))
  ).toEqual(Buffer.from(publishedBytes));
  const sourcePrefillByTag = new Map(
    (publishedContractBefore.prefillConfiguration?.fields ?? []).map(
      (field) => [field.tag, field]
    )
  );
  expect(
    duplicatePublished.templateDraft.fieldRules.map((field) => ({
      prefillPointer: field.prefillPointer,
      prefillPolicy: field.prefillPolicy,
      required: field.required,
      tag: field.tag,
    }))
  ).toEqual(
    publishedContractBefore.manifest.fields
      .map((field) => ({
        prefillPointer: sourcePrefillByTag.get(field.tag)?.pointer ?? null,
        prefillPolicy: field.prefillPolicy,
        required: field.required,
        tag: field.tag,
      }))
      .toSorted((left, right) => left.tag.localeCompare(right.tag))
  );
  expect(duplicatePublished.prefillConfiguration).toBeNull();
  const [
    duplicateResponseCount,
    duplicateSubmissionCount,
    duplicateOperationCount,
    duplicateLeaseCount,
    duplicateAuditEvents,
  ] = await Promise.all([
    prisma.response.count({ where: { formId: duplicatePublished.id } }),
    prisma.submission.count({ where: { formId: duplicatePublished.id } }),
    prisma.operation.count({ where: { formId: duplicatePublished.id } }),
    prisma.editorLease.count({
      where: {
        targetId: duplicatePublished.templateDraft.id,
        targetType: "template_draft",
      },
    }),
    prisma.auditEvent.findMany({
      orderBy: { createdAt: "asc" },
      select: { action: true, outcome: true },
      where: { targetId: duplicatePublishedPublicId },
    }),
  ]);
  expect(duplicateResponseCount).toBe(0);
  expect(duplicateSubmissionCount).toBe(0);
  expect(duplicateOperationCount).toBe(0);
  expect(duplicateLeaseCount).toBe(0);
  expect(duplicateAuditEvents).toEqual([
    { action: "duplicate_form", outcome: "success" },
  ]);
  const duplicateTitle = `${secretFormTitle} duplicate`;
  const duplicateMetadataUpdate = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        body: JSON.stringify({
          description: "Independent duplicate",
          title: duplicateTitle,
        }),
        headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
        method: "PATCH",
      }
    )
  );
  expect(duplicateMetadataUpdate.status).toBe(200);
  const sourceAfterDuplicate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(await sourceAfterDuplicate.json()).toMatchObject({
    form: { description: secretFormDescription, title: secretFormTitle },
  });
  const duplicatePublishedAfterUpdate = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(await duplicatePublishedAfterUpdate.json()).toMatchObject({
    form: {
      description: "Independent duplicate",
      publicId: duplicatePublishedPublicId,
      status: "draft",
      title: duplicateTitle,
    },
  });
  const duplicatePublishedObjectKey =
    duplicatePublished.templateDraft.objectKey;
  const deletePublishedDuplicate = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(deletePublishedDuplicate.status).toBe(200);
  expect(await objectExists(duplicatePublishedObjectKey)).toBe(false);
  expect(await objectExists(publishedContractBefore.objectKey)).toBe(true);
  const duplicateDeleteLookupResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${duplicatePublishedPublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(duplicateDeleteLookupResponse.status).toBe(404);

  const draftDuplicateSourceResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      description: "Ticket 11 draft source",
      source: "blank",
      title: "Ticket 11 draft source",
    })
  );
  expect(draftDuplicateSourceResponse.status).toBe(200);
  const draftDuplicateSourceBody =
    (await draftDuplicateSourceResponse.json()) as {
      form?: { publicId?: string };
    };
  const draftDuplicateSourcePublicId = draftDuplicateSourceBody.form?.publicId;
  if (!draftDuplicateSourcePublicId) {
    throw new Error("The draft duplicate source was not created");
  }
  const draftDuplicateSource = await prisma.form.findUniqueOrThrow({
    include: { templateDraft: true },
    where: { publicId: draftDuplicateSourcePublicId },
  });
  if (!draftDuplicateSource.templateDraft) {
    throw new Error("The draft duplicate source has no Template Draft");
  }
  const draftDuplicateResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${draftDuplicateSourcePublicId}/duplicate`,
      {
        body: "{}",
        headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
        method: "POST",
      }
    )
  );
  expect(draftDuplicateResponse.status).toBe(200);
  const draftDuplicateBody = (await draftDuplicateResponse.json()) as {
    form?: { publicId?: string; status?: string; title?: string };
  };
  const draftDuplicatePublicId = draftDuplicateBody.form?.publicId;
  if (!draftDuplicatePublicId) {
    throw new Error("The draft Form duplicate was not created");
  }
  expect(draftDuplicateBody.form).toMatchObject({
    status: "draft",
    title: "Ticket 11 draft source",
  });
  const draftDuplicate = await prisma.form.findUniqueOrThrow({
    include: { templateDraft: true },
    where: { publicId: draftDuplicatePublicId },
  });
  if (!draftDuplicate.templateDraft) {
    throw new Error("The draft Form duplicate has no Template Draft");
  }
  expect(
    Buffer.from(await readObject(draftDuplicate.templateDraft.objectKey))
  ).toEqual(
    Buffer.from(await readObject(draftDuplicateSource.templateDraft.objectKey))
  );
  const draftDuplicateObjectKey = draftDuplicate.templateDraft.objectKey;
  const draftDuplicateDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${draftDuplicatePublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(draftDuplicateDeleteResponse.status).toBe(200);
  expect(await objectExists(draftDuplicateObjectKey)).toBe(false);
  const draftDuplicateSourceDeleteResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${draftDuplicateSourcePublicId}`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
        method: "DELETE",
      }
    )
  );
  expect(draftDuplicateSourceDeleteResponse.status).toBe(200);
  const draftMetadataCreateResponse = await app.handle(
    formCreationRequest({
      authorization: adminBearer,
      description: "Ticket 10 draft metadata secret",
      source: "upload",
      template: {
        bytes: docxFixture("draft-metadata"),
        name: "draft-metadata.docx",
      },
      title: "Ticket 10 draft metadata secret",
    })
  );
  expect(draftMetadataCreateResponse.status).toBe(200);
  const draftMetadataCreateBody =
    (await draftMetadataCreateResponse.json()) as {
      form?: { publicId?: string };
    };
  const draftMetadataPublicId = draftMetadataCreateBody.form?.publicId;
  if (!draftMetadataPublicId) {
    throw new Error("The draft metadata fixture was not created");
  }
  const draftMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${draftMetadataPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftMetadataResponse.status).toBe(404);
  const draftMetadataBody = await draftMetadataResponse.json();
  expect(JSON.stringify(draftMetadataBody)).not.toContain(
    "Ticket 10 draft metadata secret"
  );
  const draftMetadataDeleteResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${draftMetadataPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(draftMetadataDeleteResponse.status).toBe(200);
  const unauthenticatedMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}`)
  );
  expect(unauthenticatedMetadataResponse.status).toBe(401);
  const unauthenticatedMetadataBody =
    (await unauthenticatedMetadataResponse.json()) as Record<string, unknown>;
  expect(JSON.stringify(unauthenticatedMetadataBody)).not.toContain(
    secretFormTitle
  );
  expect(JSON.stringify(unauthenticatedMetadataBody)).not.toContain(
    secretFormDescription
  );
  expect(JSON.stringify(unauthenticatedMetadataBody)).not.toContain(
    "published"
  );
  const authenticatedMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(authenticatedMetadataResponse.status).toBe(200);
  expect(await authenticatedMetadataResponse.json()).toMatchObject({
    form: {
      description: secretFormDescription,
      title: secretFormTitle,
    },
  });
  const archivedNoResponseEmail = `ticket-14-archived-new-${crypto.randomUUID()}@example.com`;
  await createCredentialFixture({
    email: archivedNoResponseEmail,
    name: "Ticket 14 Archived New User",
    password,
  });
  const archivedNoResponseBearer = await bearerFor(
    archivedNoResponseEmail,
    password
  );
  const mixedLifecycleUpdate = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({
        status: "archived",
        title: "must not mix lifecycle and metadata",
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(mixedLifecycleUpdate.status).toBe(400);
  expect(await mixedLifecycleUpdate.json()).toMatchObject({
    error: "invalid_request",
  });
  const unauthorizedArchive = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({ status: "archived" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "PATCH",
    })
  );
  expect(unauthorizedArchive.status).toBe(403);
  const [archiveResponse, archivedExistingStart] = await Promise.all([
    app.handle(
      new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
        body: JSON.stringify({ status: "archived" }),
        headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
        method: "PATCH",
      })
    ),
    app.handle(
      new Request(`http://test.local/api/forms/${formPublicId}/start`, {
        headers: { Authorization: `Bearer ${userBearer}` },
        method: "POST",
      })
    ),
  ]);
  expect(archiveResponse.status).toBe(200);
  expect(await archiveResponse.json()).toMatchObject({
    form: {
      publicId: formPublicId,
      status: "archived",
      version: 1,
    },
  });
  expect(archivedExistingStart.status).toBe(200);
  expect(await archivedExistingStart.json()).toMatchObject({
    response: { id: existingResponseId, status: "draft" },
  });
  const archivedNewMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
    })
  );
  expect(archivedNewMetadataResponse.status).toBe(404);
  expect(
    JSON.stringify(await archivedNewMetadataResponse.json())
  ).not.toContain(secretFormTitle);
  const archivedNewStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}/start`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "POST",
    })
  );
  expect(archivedNewStartResponse.status).toBe(409);
  expect(await archivedNewStartResponse.json()).toMatchObject({
    error: "form_unavailable",
  });
  const archivedExistingMetadataResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(archivedExistingMetadataResponse.status).toBe(200);
  expect(await archivedExistingMetadataResponse.json()).toMatchObject({
    form: {
      publicId: formPublicId,
      title: secretFormTitle,
    },
  });
  const archivedFormListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(archivedFormListResponse.status).toBe(200);
  const archivedFormListBody = (await archivedFormListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      status: string;
      submissionCount: number;
    }[];
  };
  expect(
    archivedFormListBody.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({
    activeDraftCount: 1,
    publicId: formPublicId,
    status: "archived",
    submissionCount: 0,
  });
  const archivedContract = await prisma.publishedTemplate.findUnique({
    include: {
      manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
      prefillConfiguration: {
        include: { fields: { orderBy: { tag: "asc" } } },
      },
    },
    where: { formId },
  });
  expect(archivedContract).toEqual(publishedContractBefore);
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "archive_form",
        outcome: "success",
        targetId: formPublicId,
      },
    })
  ).toMatchObject({ actorId: adminId, targetId: formPublicId });
  const userDeleteFormResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(userDeleteFormResponse.status).toBe(403);

  const startResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(startResponse.status).toBe(200);
  const startBody = (await startResponse.json()) as {
    response?: { id?: string };
  };
  const responseId = startBody.response?.id;
  if (!responseId) {
    throw new Error("The response was not started");
  }
  const draftCountListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(draftCountListResponse.status).toBe(200);
  const draftCountList = (await draftCountListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      submissionCount: number;
    }[];
  };
  expect(
    draftCountList.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({ activeDraftCount: 1, submissionCount: 0 });
  const formLifecycleBeforeResponseGuard = await prisma.form.findUnique({
    select: { status: true, version: true },
    where: { id: formId },
  });
  if (!formLifecycleBeforeResponseGuard) {
    throw new Error("The Form lifecycle guard baseline was not found");
  }
  await prisma.form.update({
    data: { status: "draft", version: 0 },
    where: { id: formId },
  });
  const responseProtectedDelete = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
      method: "DELETE",
    })
  );
  expect(responseProtectedDelete.status).toBe(409);
  expect(await responseProtectedDelete.json()).toMatchObject({
    error: "form_has_responses",
  });
  const responseProtectedTemplate = await prisma.templateDraft.findUnique({
    select: { objectKey: true },
    where: { formId },
  });
  expect(responseProtectedTemplate).not.toBeNull();
  if (responseProtectedTemplate) {
    expect(await objectExists(responseProtectedTemplate.objectKey)).toBe(true);
  }
  await prisma.form.update({
    data: formLifecycleBeforeResponseGuard,
    where: { id: formId },
  });

  const editorResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${formRecord.publicId}/editor-config?responseId=${responseId}&action=fill`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(editorResponse.status).toBe(200);
  const editorConfig = (await editorResponse.json()) as EditorConfigBody;
  let responseDocumentKey = editorConfig.config.document.key;
  const originalResponseDocumentKey = responseDocumentKey;
  const userPluginOptions =
    editorConfig.config.editorConfig.plugins.options[pluginGuid];
  let saveDraftCapability: string =
    editorConfig.bridge.capabilities["save-draft"] ?? "";
  let submitCapability: string = editorConfig.bridge.capabilities.submit ?? "";
  if (
    !userPluginOptions ||
    !responseDocumentKey ||
    !saveDraftCapability ||
    !submitCapability
  ) {
    throw new Error("The User editor capabilities were not returned");
  }
  let userLease = editorConfig.bridge.lease;
  expect(userLease).toEqual({
    expiresAt: expect.any(String),
    id: expect.any(String),
    releaseUrl: expect.any(String),
    renewUrl: expect.any(String),
  });
  for (const leaseValue of Object.values(userLease)) {
    expect(JSON.stringify(editorConfig.config)).not.toContain(leaseValue);
  }
  expect(userPluginOptions).not.toHaveProperty("lease");
  const originalSaveDraftCapability = saveDraftCapability;
  const refreshResponseEditor = async (): Promise<void> => {
    const refreshedResponse = await app.handle(
      new Request(
        `http://test.local/api/forms/${formRecord.publicId}/editor-config?responseId=${responseId}&action=draft`,
        { headers: { Authorization: `Bearer ${userBearer}` } }
      )
    );
    expect(refreshedResponse.status).toBe(200);
    const refreshedConfig =
      (await refreshedResponse.json()) as EditorConfigBody;
    const refreshedSaveDraftCapability =
      refreshedConfig.bridge.capabilities["save-draft"];
    const refreshedSubmitCapability =
      refreshedConfig.bridge.capabilities.submit;
    if (!refreshedSaveDraftCapability || !refreshedSubmitCapability) {
      throw new Error(
        "The refreshed User editor capabilities were not returned"
      );
    }
    responseDocumentKey = refreshedConfig.config.document.key;
    saveDraftCapability = refreshedSaveDraftCapability;
    submitCapability = refreshedSubmitCapability;
    userLease = refreshedConfig.bridge.lease;
  };
  await prisma.editorLease.update({
    data: { createdAt: new Date(0), expiresAt: new Date(1) },
    where: { id: userLease.id },
  });

  const previousUserLeaseId = userLease.id;
  await refreshResponseEditor();
  expect(userLease.id).not.toBe(previousUserLeaseId);
  const staleUserLeaseRelease = await app.handle(
    new Request(`http://test.local/api/editor-leases/${previousUserLeaseId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(staleUserLeaseRelease.status).toBe(409);
  expect(await staleUserLeaseRelease.json()).toMatchObject({
    error: "editor_lease_inactive",
  });
  const activeLeaseOperationId = crypto.randomUUID();
  await prisma.operation.create({
    data: {
      actorId: userId,
      documentKey: responseDocumentKey,
      formId,
      id: activeLeaseOperationId,
      metadata: {
        action: "save-draft",
        finalObjectKey: `operations/${activeLeaseOperationId}/final.docx`,
        formId,
        responseId,
        stagedObjectKey: `operations/${activeLeaseOperationId}/staged.docx`,
      },
      ownerUserId: userId,
      responseId,
      stagingObjectKey: `operations/${activeLeaseOperationId}/staged.docx`,
      status: "processing",
      targetId: responseId,
      targetType: "response",
      type: "save_draft",
    },
  });
  const releaseDuringOperationResponse = await app.handle(
    new Request(`http://test.local${userLease.releaseUrl}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(releaseDuringOperationResponse.status).toBe(200);
  expect(
    await prisma.editorLease.findUnique({ where: { id: userLease.id } })
  ).not.toBeNull();
  await prisma.operation.delete({ where: { id: activeLeaseOperationId } });
  const sessionOnlyUserDraftResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyUserDraftResponse.status).toBe(401);
  expect(await sessionOnlyUserDraftResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const sessionOnlyUserSubmitResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/submit`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(sessionOnlyUserSubmitResponse.status).toBe(401);
  expect(await sessionOnlyUserSubmitResponse.json()).toMatchObject({
    error: "editor_capability_required",
  });
  const userEditorSerialized = JSON.stringify(editorConfig);
  expect(userEditorSerialized).not.toContain(userBearer);
  expect(userEditorSerialized).not.toContain('"authToken"');
  expect(JSON.stringify(editorConfig.config)).not.toContain(
    saveDraftCapability
  );
  expect(editorConfig.bridge.id).toBe(userPluginOptions.bridgeId);
  const saveDraftClaims = verifyEditorCapability(saveDraftCapability);
  expect(saveDraftClaims).toMatchObject({
    action: "save-draft",
    actorId: userId,
    documentKey: responseDocumentKey,
    formId,
    leaseId: userLease.id,
    leaseProof: expect.any(String),
    role: "user",
    targetId: responseId,
    targetType: "response",
  });
  if (!saveDraftClaims) {
    throw new Error("The save-draft capability was invalid");
  }
  const submitClaims = verifyEditorCapability(submitCapability);
  expect(submitClaims).toMatchObject({
    action: "submit",
    actorId: userId,
    documentKey: responseDocumentKey,
    formId,
    leaseId: userLease.id,
    leaseProof: expect.any(String),
    role: "user",
    targetId: responseId,
    targetType: "response",
  });
  if (!submitClaims) {
    throw new Error("The submit capability was invalid");
  }
  await prisma.user.update({
    data: { enabled: false },
    where: { id: userId },
  });
  const disabledActorResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect(disabledActorResponse.status).toBe(401);
  await prisma.user.update({
    data: { enabled: true, role: "admin" },
    where: { id: userId },
  });
  const changedRoleResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect(changedRoleResponse.status).toBe(401);
  await prisma.user.update({
    data: { role: "user" },
    where: { id: userId },
  });
  const userCrossActionResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(submitCapability),
      method: "POST",
    })
  );
  expect(userCrossActionResponse.status).toBe(403);
  const draftRequest = (data: Record<string, unknown>) =>
    app.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
        body: JSON.stringify({
          data,
          documentKey: responseDocumentKey,
          responseId,
        }),
        headers: capabilityHeaders(saveDraftCapability),
        method: "POST",
      })
    );
  const invalidDraftCases = [
    { data: { unknown_tag: "value" }, status: 422 },
    { data: { accept_terms: "true" }, status: 422 },
    { data: { department: "not-an-option" }, status: 422 },
    { data: { start_date: "2026-02-30" }, status: 422 },
    { data: { description_1: "x".repeat(10_001) }, status: 422 },
    { data: { description_1: "🙂".repeat(100_000) }, status: 413 },
  ];
  for (const invalidDraftCase of invalidDraftCases) {
    const invalidDraftResponse = await draftRequest(invalidDraftCase.data);
    expect(invalidDraftResponse.status).toBe(invalidDraftCase.status);
    expect(await invalidDraftResponse.json()).toHaveProperty(
      "error",
      invalidDraftCase.status === 413
        ? "response_too_large"
        : "invalid_response_data"
    );
  }
  const trustedPrefillValue = selectedPointer.endsWith("/active")
    ? true
    : "Ticket 17 Prefill";
  const lockedSnapshotCases = [
    { error: "invalid_response_data", status: 422, value: "x".repeat(10_001) },
    { error: "response_too_large", status: 413, value: "🙂".repeat(100_000) },
    { error: "invalid_response_data", status: 422, value: true },
  ];
  for (const lockedSnapshotCase of lockedSnapshotCases) {
    await prisma.prefillSnapshot.update({
      data: { values: { full_name: lockedSnapshotCase.value } },
      where: { responseId },
    });
    const invalidTrustedValueResponse = await draftRequest({
      full_name: "client value",
    });
    expect(invalidTrustedValueResponse.status).toBe(lockedSnapshotCase.status);
    expect(await invalidTrustedValueResponse.json()).toMatchObject({
      error: lockedSnapshotCase.error,
    });
  }
  await prisma.prefillSnapshot.update({
    data: { values: { full_name: trustedPrefillValue } },
    where: { responseId },
  });
  const oversizedClientValueResponse = await draftRequest({
    full_name: "x".repeat(10_001),
  });
  expect(oversizedClientValueResponse.status).toBe(202);
  const oversizedClientValueBody =
    (await oversizedClientValueResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !oversizedClientValueBody.operationCapability ||
    !oversizedClientValueBody.operationId
  ) {
    throw new Error("The locked client value operation was not created");
  }
  const oversizedClientValueOperation = await waitForOperation(
    oversizedClientValueBody.operationId,
    { "X-Editor-Capability": oversizedClientValueBody.operationCapability }
  );
  expect(oversizedClientValueOperation.status).toBe("completed");
  await refreshResponseEditor();
  const normalizedClientTamperResponse =
    await prisma.response.findUniqueOrThrow({
      select: { draftData: true },
      where: { id: responseId },
    });
  expect(normalizedClientTamperResponse.draftData).toEqual({
    full_name: trustedPrefillValue,
  });
  const savedDraftData = {
    accept_terms: true,
    department: "engineering",
    description_1: "line one\nline two",
    description_2: "",
    full_name: selectedPointer.endsWith("/active") ? true : "Ticket 17 Prefill",
    start_date: "2026-09-15",
  };
  const submitRequest = (data: Record<string, unknown>, targetApp = app) =>
    targetApp.handle(
      new Request(`http://test.local/api/forms/${formRecord.publicId}/submit`, {
        body: JSON.stringify({
          data,
          documentKey: responseDocumentKey,
          responseId,
        }),
        headers: capabilityHeaders(submitCapability),
        method: "POST",
      })
    );
  const invalidSubmitCases = [
    { ...savedDraftData, accept_terms: false },
    { ...savedDraftData, department: "not-an-option" },
    { ...savedDraftData, start_date: "2026-02-30" },
  ];
  for (const invalidSubmitData of invalidSubmitCases) {
    const invalidSubmitResponse = await submitRequest(invalidSubmitData);
    expect(invalidSubmitResponse.status).toBe(422);
    expect(await invalidSubmitResponse.json()).toMatchObject({
      error: "invalid_response_data",
    });
  }

  const saveResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: savedDraftData,
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect(saveResponse.status).toBe(202);
  const saveBody = (await saveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!saveBody.operationCapability || !saveBody.operationId) {
    throw new Error("The draft operation was not created");
  }
  const saveOperation = await waitForOperation(saveBody.operationId, {
    "X-Editor-Capability": saveBody.operationCapability,
  });
  expect(saveOperation.status).toBe("completed");
  const saveOperationResult = saveOperation.result as {
    documentKey?: unknown;
    responseId?: unknown;
  };
  expect(saveOperationResult.documentKey).toEqual(expect.any(String));
  expect(saveOperationResult.responseId).toBe(responseId);
  const savedResponse = await prisma.response.findUnique({
    select: {
      draftData: true,
      draftDocumentKey: true,
      draftObjectKey: true,
      status: true,
      updatedAt: true,
    },
    where: { id: responseId },
  });
  const savedResponseDocumentKey = savedResponse?.draftDocumentKey;
  const savedResponseObjectKey = savedResponse?.draftObjectKey;
  expect(savedResponse).toMatchObject({
    draftData: savedDraftData,
    draftDocumentKey: expect.any(String),
    status: "draft",
    updatedAt: expect.any(Date),
  });
  if (!savedResponseDocumentKey || !savedResponseObjectKey) {
    throw new Error("The saved response document was not persisted");
  }
  expect(savedResponseDocumentKey).not.toBe(originalResponseDocumentKey);
  expect(saveOperationResult.documentKey).toBe(savedResponseDocumentKey);
  const staleDraftResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: savedDraftData,
        documentKey: originalResponseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(originalSaveDraftCapability),
      method: "POST",
    })
  );
  expect(staleDraftResponse.status).toBe(409);
  expect(await staleDraftResponse.json()).toMatchObject({
    error: "stale_response",
  });
  responseDocumentKey = savedResponseDocumentKey;
  const resumedStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(resumedStartResponse.status).toBe(200);
  expect(await resumedStartResponse.json()).toMatchObject({
    response: { id: responseId, status: "draft" },
  });
  const resumedEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${formRecord.publicId}/editor-config?responseId=${responseId}&action=draft`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(resumedEditorResponse.status).toBe(200);
  const resumedEditorConfig =
    (await resumedEditorResponse.json()) as EditorConfigBody;
  expect(resumedEditorConfig).toMatchObject({
    config: { document: { key: responseDocumentKey } },
  });
  const resumedSaveDraftCapability =
    resumedEditorConfig.bridge.capabilities["save-draft"];
  const resumedSubmitCapability =
    resumedEditorConfig.bridge.capabilities.submit;
  if (!resumedSaveDraftCapability || !resumedSubmitCapability) {
    throw new Error("The resumed User editor capabilities were not returned");
  }
  saveDraftCapability = resumedSaveDraftCapability;
  submitCapability = resumedSubmitCapability;
  userLease = resumedEditorConfig.bridge.lease;
  const responseList = await app.handle(
    new Request("http://test.local/api/responses/me", {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(await responseList.json()).toMatchObject({
    responses: [
      {
        formPublicId: formRecord.publicId,
        formTitle: secretFormTitle,
        id: responseId,
        status: "draft",
        updatedAt: expect.any(String),
      },
    ],
  });
  const paginationUsers = Array.from({ length: 26 }, (_, index) => ({
    email: `ticket20-pagination-${String(index).padStart(2, "0")}@example.com`,
    id: `ticket20-pagination-${String(index).padStart(2, "0")}`,
    name: `Ticket 20 pagination ${index}`,
  }));
  const [firstPaginationUser] = paginationUsers;
  if (!firstPaginationUser) {
    throw new Error("The pagination fixture user was not created");
  }
  await prisma.user.createMany({ data: paginationUsers });
  await prisma.response.createMany({
    data: paginationUsers.map((paginationUser, index) => ({
      createdAt: new Date(2020, 0, index + 1),
      draftDocumentKey: `responses/pagination-${index}/draft/document.docx`,
      draftObjectKey: `responses/pagination-${index}/draft/object.docx`,
      formId,
      publishedTemplateId: publishedManifestRecord.id,
      publishedVersion: publishedManifestRecord.version,
      updatedAt: new Date(2020, 0, index + 1),
      userId: paginationUser.id,
    })),
  });
  const paginatedResultsFirstResponse = await app.handle(
    new Request("http://test.local/api/admin/results", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(paginatedResultsFirstResponse.status).toBe(200);
  const paginatedResultsFirstBody =
    (await paginatedResultsFirstResponse.json()) as {
      nextCursor: string | null;
      results: { id: string }[];
    };
  expect(paginatedResultsFirstBody.results).toHaveLength(25);
  expect(paginatedResultsFirstBody.nextCursor).toEqual(expect.any(String));
  const paginatedResultsSecondResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?cursor=${encodeURIComponent(paginatedResultsFirstBody.nextCursor as string)}`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(paginatedResultsSecondResponse.status).toBe(200);
  const paginatedResultsSecondBody =
    (await paginatedResultsSecondResponse.json()) as {
      nextCursor: string | null;
      results: { id: string }[];
    };
  expect(paginatedResultsSecondBody.results.length).toBeGreaterThan(0);
  expect(
    new Set([
      ...paginatedResultsFirstBody.results.map((result) => result.id),
      ...paginatedResultsSecondBody.results.map((result) => result.id),
    ]).size
  ).toBe(
    paginatedResultsFirstBody.results.length +
      paginatedResultsSecondBody.results.length
  );
  const filteredPaginationResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&user=${encodeURIComponent(firstPaginationUser.email)}&state=draft&correction=0`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(filteredPaginationResponse.status).toBe(200);
  expect(await filteredPaginationResponse.json()).toMatchObject({
    nextCursor: null,
    results: [
      {
        formPublicId: formRecord.publicId,
        state: "draft",
        userEmail: firstPaginationUser.email,
      },
    ],
  });
  const excludedPaginationResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&user=${encodeURIComponent(firstPaginationUser.email)}&state=submitted`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(excludedPaginationResponse.status).toBe(200);
  expect(await excludedPaginationResponse.json()).toEqual({
    nextCursor: null,
    results: [],
  });
  const adminDraftResultsResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&user=${encodeURIComponent(userEmail)}&state=draft&correction=0&from=2020-01-01T00:00:00.000Z&to=2030-01-01T00:00:00.000Z`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminDraftResultsResponse.status).toBe(200);
  const adminDraftResultsBody = (await adminDraftResultsResponse.json()) as {
    nextCursor: string | null;
    results: Record<string, unknown>[];
  };
  await prisma.response.deleteMany({
    where: {
      userId: {
        in: paginationUsers.map((paginationUser) => paginationUser.id),
      },
    },
  });
  await prisma.user.deleteMany({
    where: {
      id: { in: paginationUsers.map((paginationUser) => paginationUser.id) },
    },
  });
  expect(adminDraftResultsBody).toMatchObject({
    nextCursor: null,
    results: [
      {
        formPublicId: formRecord.publicId,
        id: responseId,
        latestCorrectionNumber: null,
        state: "draft",
        userEmail,
      },
    ],
  });
  const adminDraftDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/results/${responseId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminDraftDetailResponse.status).toBe(200);
  const adminDraftDetailBody = (await adminDraftDetailResponse.json()) as {
    result: Record<string, unknown>;
  };
  expect(adminDraftDetailBody.result).toMatchObject({
    data: savedDraftData,
    document: { available: true, state: "draft" },
    id: responseId,
    state: "draft",
  });
  expect(JSON.stringify(adminDraftDetailBody)).not.toContain("draftObjectKey");
  const adminDraftExportResponse = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/json`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminDraftExportResponse.status).toBe(403);
  const userResultsResponse = await app.handle(
    new Request("http://test.local/api/admin/results", {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(userResultsResponse.status).toBe(403);
  const invalidAdminResultsCursor = await app.handle(
    new Request("http://test.local/api/admin/results?cursor=not-a-cursor", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(invalidAdminResultsCursor.status).toBe(400);
  const draftViewAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "view_response",
      actorId: adminId,
      targetId: responseId,
      targetType: "response",
    },
  });
  expect(draftViewAudit).toMatchObject({
    outcome: "success",
    safeMetadata: { state: "draft" },
  });
  expect(draftViewAudit.safeMetadata).toEqual({
    revision: 0,
    state: "draft",
  });
  const draftResponseBeforeExport = await prisma.response.findUnique({
    select: { draftData: true, draftObjectKey: true },
    where: { id: responseId },
  });
  if (!draftResponseBeforeExport?.draftObjectKey) {
    throw new Error("The Draft export fixture was not created");
  }
  const draftJsonExport = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/json`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftJsonExport.status).toBe(200);
  expect(draftJsonExport.headers.get("content-type")).toBe(
    "application/json; charset=utf-8"
  );
  expect(draftJsonExport.headers.get("content-disposition")).toBe(
    `attachment; filename="response-${responseId}.json"`
  );
  expect(JSON.parse(await draftJsonExport.text())).toEqual(savedDraftData);
  const draftDocxExport = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/docx`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftDocxExport.status).toBe(200);
  expect(draftDocxExport.headers.get("content-type")).toBe(DOCX_CONTENT_TYPE);
  expect(draftDocxExport.headers.get("content-disposition")).toBe(
    `attachment; filename="response-${responseId}.docx"`
  );
  expect(new Uint8Array(await draftDocxExport.arrayBuffer())).toEqual(
    Uint8Array.from(await readObject(draftResponseBeforeExport.draftObjectKey))
  );
  const draftPdfExport = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/draft/pdf`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(draftPdfExport.status).toBe(200);
  expect(draftPdfExport.headers.get("content-type")).toBe("application/pdf");
  expect(draftPdfExport.headers.get("content-disposition")).toBe(
    `attachment; filename="response-${responseId}.pdf"`
  );
  expect(await draftPdfExport.text()).toBe("%PDF-test");

  const stableBeforeFailure = await prisma.response.findUnique({
    select: { draftData: true, draftObjectKey: true },
    where: { id: responseId },
  });
  const failureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.reject(new Error("deterministic draft failure")),
      forceSave: () => Promise.reject(new Error("deterministic draft failure")),
    },
  });
  const failedSaveResponse = await failureApp.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: savedDraftData,
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect(failedSaveResponse.status).toBe(202);
  const failedSaveBody = (await failedSaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!failedSaveBody.operationCapability || !failedSaveBody.operationId) {
    throw new Error("The failed draft operation was not created");
  }
  const failedSaveOperation = await waitForOperation(
    failedSaveBody.operationId,
    { "X-Editor-Capability": failedSaveBody.operationCapability }
  );
  expect(failedSaveOperation).toMatchObject({
    error: "force_save_failed",
    status: "failed",
  });
  expect(
    await prisma.response.findUnique({
      select: { draftData: true, draftObjectKey: true },
      where: { id: responseId },
    })
  ).toEqual(stableBeforeFailure);
  const retryDocumentKeyBeforeSave = responseDocumentKey;
  const retrySaveResponse = await draftRequest(savedDraftData);
  expect(retrySaveResponse.status).toBe(202);
  const retrySaveBody = (await retrySaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!retrySaveBody.operationCapability || !retrySaveBody.operationId) {
    throw new Error("The retry draft operation was not created");
  }
  const retrySaveOperation = await waitForOperation(retrySaveBody.operationId, {
    "X-Editor-Capability": retrySaveBody.operationCapability,
  });
  expect(retrySaveOperation.status).toBe("completed");
  await refreshResponseEditor();
  expect(responseDocumentKey).not.toBe(retryDocumentKeyBeforeSave);
  const stableBeforeSubmitFailure = await prisma.response.findUnique({
    select: { draftData: true, draftObjectKey: true, status: true },
    where: { id: responseId },
  });
  const failedSubmitResponse = await submitRequest(savedDraftData, failureApp);
  expect(failedSubmitResponse.status).toBe(202);
  const failedSubmitBody = (await failedSubmitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (!failedSubmitBody.operationCapability || !failedSubmitBody.operationId) {
    throw new Error("The failed submit operation was not created");
  }
  const failedSubmitOperation = await waitForOperation(
    failedSubmitBody.operationId,
    { "X-Editor-Capability": failedSubmitBody.operationCapability }
  );
  expect(failedSubmitOperation).toMatchObject({
    error: "force_save_failed",
    status: "failed",
  });
  expect(
    await prisma.response.findUnique({
      select: { draftData: true, draftObjectKey: true, status: true },
      where: { id: responseId },
    })
  ).toEqual(stableBeforeSubmitFailure);
  expect(await prisma.submission.count({ where: { responseId } })).toBe(0);

  const concurrentSubmitResponses = await Promise.all([
    submitRequest(savedDraftData),
    submitRequest(savedDraftData),
  ]);
  expect(
    concurrentSubmitResponses.filter((response) => response.status === 202)
  ).toHaveLength(1);
  const rejectedSubmitResponses = concurrentSubmitResponses.filter(
    (response) => response.status === 409
  );
  expect(rejectedSubmitResponses).toHaveLength(1);
  expect(await rejectedSubmitResponses[0]?.json()).toMatchObject({
    error: "operation_in_progress",
  });
  const submitResponse = concurrentSubmitResponses.find(
    (response) => response.status === 202
  );
  if (!submitResponse) {
    throw new Error("The concurrent submit operation was not created");
  }
  const submitBody = (await submitResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
    submissionId?: string;
  };
  const completedSubmissionId = submitBody.submissionId;
  if (
    !submitBody.operationCapability ||
    !submitBody.operationId ||
    !completedSubmissionId
  ) {
    throw new Error("The submit operation was not created");
  }
  const submitOperation = await waitForOperation(submitBody.operationId, {
    "X-Editor-Capability": submitBody.operationCapability,
  });
  expect(submitOperation.status).toBe("completed");
  const repeatStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "POST",
    })
  );
  expect(repeatStartResponse.status).toBe(200);
  expect(await repeatStartResponse.json()).toMatchObject({
    receiptUrl: `/receipt/${completedSubmissionId}`,
    response: {
      id: responseId,
      status: "submitted",
      submissionId: completedSubmissionId,
    },
    submissionId: completedSubmissionId,
  });
  expect(
    await prisma.response.count({
      where: { formId, userId },
    })
  ).toBe(1);
  const postSubmitUnarchive = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({ status: "published" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(postSubmitUnarchive.status).toBe(200);
  const postSubmitUnarchiveBody = (await postSubmitUnarchive.json()) as {
    form?: { publicId?: string; status?: string };
  };
  expect(postSubmitUnarchiveBody.form).toMatchObject({
    publicId: formPublicId,
    status: "published",
  });
  const postSubmitExternalReference = `ticket-17-post-submit-${crypto.randomUUID()}`;
  const postSubmitHandoffCreate = await fetch(`${mock.url}/handoffs`, {
    body: JSON.stringify({
      email: userEmail,
      externalReference: postSubmitExternalReference,
      publicId: formRecord.publicId,
      values: handoffCandidateValues,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(postSubmitHandoffCreate.status).toBe(200);
  const postSubmitCode = (
    (await postSubmitHandoffCreate.json()) as {
      code: string;
    }
  ).code;
  if (!postSubmitCode) {
    throw new Error("The post-submit handoff code was not created");
  }
  const postSubmitLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: postSubmitCode }),
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
      },
      method: "POST",
    })
  );
  expect(postSubmitLaunch.status).toBe(303);
  expect(postSubmitLaunch.headers.get("location")).toBe(
    `/forms/${formRecord.publicId}/fill`
  );
  const postSubmitCookie = postSubmitLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!postSubmitCookie) {
    throw new Error("The post-submit handoff did not set a cookie");
  }
  const postSubmitStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: postSubmitCookie,
      },
      method: "POST",
    })
  );
  expect(postSubmitStart.status).toBe(200);
  expect(await postSubmitStart.json()).toMatchObject({
    receiptUrl: `/receipt/${completedSubmissionId}`,
    response: { id: responseId, status: "submitted" },
    submissionId: completedSubmissionId,
  });
  expect(postSubmitStart.headers.get("set-cookie")).toContain("Max-Age=0");
  const postSubmitHandoffRecord = await prisma.handoff.findFirstOrThrow({
    where: {
      externalReferenceDigest: createHash("sha256")
        .update(postSubmitExternalReference)
        .digest("hex"),
    },
  });
  expect(postSubmitHandoffRecord).toMatchObject({
    responseId,
    status: "consumed",
  });
  expect(
    await prisma.response.count({
      where: { formId, userId },
    })
  ).toBe(1);
  const replayPostSubmitStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: {
        Authorization: `Bearer ${userBearer}`,
        Cookie: postSubmitCookie,
      },
      method: "POST",
    })
  );
  expect(replayPostSubmitStart.status).toBe(409);
  expect(await replayPostSubmitStart.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "redeem_handoff",
        outcome: "failure",
        targetId: formPublicId,
      },
    })
  ).toMatchObject({ targetId: formPublicId });
  const submittedStatusResponse = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(submittedStatusResponse.status).toBe(200);
  expect(await submittedStatusResponse.json()).toMatchObject({
    latestCorrectionNumber: null,
    status: "submitted",
    submittedAt: expect.any(String),
  });
  const immutableDraftResponse = await draftRequest(savedDraftData);
  expect(immutableDraftResponse.status).toBe(409);
  expect(await immutableDraftResponse.json()).toMatchObject({
    error: "stale_response",
  });
  const submittedCountListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(submittedCountListResponse.status).toBe(200);
  const submittedCountList = (await submittedCountListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      submissionCount: number;
    }[];
  };
  expect(
    submittedCountList.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({ activeDraftCount: 0, submissionCount: 1 });
  const ownerOperationVisibilityResponse = await app.handle(
    new Request(`http://test.local/api/operations/${submitBody.operationId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(ownerOperationVisibilityResponse.status).toBe(200);
  expect(await ownerOperationVisibilityResponse.json()).toMatchObject({
    operation: { id: submitBody.operationId, status: "completed" },
  });
  const adminOperationVisibilityResponse = await app.handle(
    new Request(`http://test.local/api/operations/${submitBody.operationId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminOperationVisibilityResponse.status).toBe(200);
  expect(await adminOperationVisibilityResponse.json()).toMatchObject({
    operation: { id: submitBody.operationId, status: "completed" },
  });
  const crossOperationPollResponse = await app.handle(
    new Request(`http://test.local/api/operations/${submitBody.operationId}`, {
      headers: {
        "X-Editor-Capability": saveBody.operationCapability,
      },
    })
  );
  expect(crossOperationPollResponse.status).toBe(403);
  const otherUser = await createCredentialFixture({
    email: otherUserEmail,
    name: "Ticket 04 Other User",
    password,
  });
  const otherUserBearer = await bearerFor(otherUserEmail, password);
  const forbiddenDocxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenDocxResponse.status).toBe(403);
  const forbiddenPdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenPdfResponse.status).toBe(403);
  const forbiddenDataResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenDataResponse.status).toBe(403);

  await prisma.$transaction(async (tx) => {
    await tx.$executeRaw`ALTER TABLE "field_manifests" DISABLE TRIGGER "field_manifests_immutable"`;
    await tx.$executeRaw`ALTER TABLE "manifest_fields" DISABLE TRIGGER "manifest_fields_immutable"`;
    try {
      await tx.$executeRaw`
        UPDATE "manifest_fields"
        SET "position" = "position" + 100000
        WHERE "manifest_id" = ${publishedManifest.id}::uuid
      `;
      await tx.$executeRaw`
        WITH ranked_fields AS (
          SELECT
            "id",
            (ROW_NUMBER() OVER (ORDER BY "tag") - 1)::integer AS "position"
          FROM "manifest_fields"
          WHERE "manifest_id" = ${publishedManifest.id}::uuid
        )
        UPDATE "manifest_fields" AS field
        SET
          "label" = field."tag",
          "placeholder" = NULL,
          "position" = ranked_fields."position"
        FROM ranked_fields
        WHERE field."id" = ranked_fields."id"
      `;
      await tx.fieldManifest.update({
        data: { displayMetadataVersion: 0 },
        where: { id: publishedManifest.id },
      });
    } finally {
      await tx.$executeRaw`ALTER TABLE "manifest_fields" ENABLE TRIGGER "manifest_fields_immutable"`;
      await tx.$executeRaw`ALTER TABLE "field_manifests" ENABLE TRIGGER "field_manifests_immutable"`;
    }
  });
  const dataResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(dataResponse.status).toBe(200);
  const dataBody = (await dataResponse.json()) as {
    data: Record<string, unknown>;
    fields: {
      label: string;
      options: { displayText: string; value: string }[];
      placeholder: string | null;
      position: number;
      tag: string;
      type: string;
    }[];
    returnUrl: string;
    submission: Record<string, unknown>;
  };
  expect(dataBody).toMatchObject({
    data: savedDraftData,
    returnUrl: "https://source.example.test/forms/return",
    submission: { id: completedSubmissionId, responseId },
  });
  expect(dataBody.submission).not.toHaveProperty("userEmail");
  expect(
    dataBody.fields.map(({ label, placeholder, position, tag, type }) => ({
      label,
      placeholder,
      position,
      tag,
      type,
    }))
  ).toEqual(
    [...publishedManifestRecord.manifest.fields]
      .toSorted((left, right) => left.position - right.position)
      .map(({ label, placeholder, position, tag, type }) => ({
        label,
        placeholder,
        position,
        tag,
        type,
      }))
  );
  expect(
    dataBody.fields.find((field) => field.tag === "department")?.options
  ).toEqual([
    { displayText: "Choose an item", value: "" },
    { displayText: "Engineering", value: "engineering" },
    { displayText: "Human Resources", value: "hr" },
    { displayText: "Finance", value: "finance" },
  ]);
  await prisma.$transaction(async (tx) => {
    await tx.$executeRaw`ALTER TABLE "field_manifests" DISABLE TRIGGER "field_manifests_immutable"`;
    await tx.$executeRaw`ALTER TABLE "manifest_fields" DISABLE TRIGGER "manifest_fields_immutable"`;
    try {
      await tx.$executeRaw`
        UPDATE "manifest_fields"
        SET "position" = "position" + 100000
        WHERE "manifest_id" = ${publishedManifest.id}::uuid
      `;
      for (const field of dataBody.fields) {
        await tx.$executeRaw`
          UPDATE "manifest_fields"
          SET
            "label" = ${field.label},
            "placeholder" = ${field.placeholder},
            "position" = ${field.position}
          WHERE "manifest_id" = ${publishedManifest.id}::uuid AND "tag" = ${field.tag}
        `;
      }
      await tx.fieldManifest.update({
        data: {
          displayMetadataVersion: publishedManifest.displayMetadataVersion,
        },
        where: { id: publishedManifest.id },
      });
    } finally {
      await tx.$executeRaw`ALTER TABLE "manifest_fields" ENABLE TRIGGER "manifest_fields_immutable"`;
      await tx.$executeRaw`ALTER TABLE "field_manifests" ENABLE TRIGGER "field_manifests_immutable"`;
    }
  });
  const ownerJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(ownerJsonResponse.status).toBe(200);
  expect(ownerJsonResponse.headers.get("content-type")).toBe(
    "application/json; charset=utf-8"
  );
  expect(ownerJsonResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.json"`
  );
  expect(JSON.parse(await ownerJsonResponse.text())).toEqual(savedDraftData);
  const adminJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminJsonResponse.status).toBe(200);
  expect(adminJsonResponse.headers.get("content-type")).toBe(
    "application/json; charset=utf-8"
  );
  expect(adminJsonResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.json"`
  );
  expect(JSON.parse(await adminJsonResponse.text())).toEqual(savedDraftData);
  const forbiddenJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json`,
      { headers: { Authorization: `Bearer ${otherUserBearer}` } }
    )
  );
  expect(forbiddenJsonResponse.status).toBe(403);

  const docxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(docxResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.docx"`
  );
  expect(docxResponse.status).toBe(200);
  expect(docxResponse.headers.get("content-type")).toBe(DOCX_CONTENT_TYPE);
  const submissionDocument = new Uint8Array(await docxResponse.arrayBuffer());
  const templateDraft = await prisma.templateDraft.findUnique({
    select: { documentKey: true, id: true, objectKey: true },
    where: { formId },
  });
  if (!templateDraft) {
    throw new Error("The test template draft was not found");
  }
  const sourceDocument = await readObject(templateDraft.objectKey);
  expect(submissionDocument).toEqual(Uint8Array.from(sourceDocument));
  const adminDocxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminDocxResponse.status).toBe(200);
  expect(adminDocxResponse.headers.get("content-type")).toBe(DOCX_CONTENT_TYPE);
  expect(adminDocxResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.docx"`
  );
  await adminDocxResponse.arrayBuffer();

  const pdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(pdfResponse.status).toBe(200);
  expect(pdfResponse.headers.get("content-type")).toBe("application/pdf");
  expect(pdfResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.pdf"`
  );
  expect(await pdfResponse.text()).toBe("%PDF-test");
  const adminPdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminPdfResponse.status).toBe(200);
  expect(adminPdfResponse.headers.get("content-type")).toBe("application/pdf");
  expect(adminPdfResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}.pdf"`
  );
  expect(await adminPdfResponse.text()).toBe("%PDF-test");
  const adminSubmittedResultsResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results?form=${formRecord.publicId}&state=submitted`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(adminSubmittedResultsResponse.status).toBe(200);
  expect(await adminSubmittedResultsResponse.json()).toMatchObject({
    results: [
      {
        formPublicId: formRecord.publicId,
        id: responseId,
        latestCorrectionNumber: null,
        state: "submitted",
        submissionId: completedSubmissionId,
      },
    ],
  });
  const adminSubmittedDetailResponse = await app.handle(
    new Request(`http://test.local/api/admin/results/${responseId}`, {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  expect(adminSubmittedDetailResponse.status).toBe(200);
  expect(await adminSubmittedDetailResponse.json()).toMatchObject({
    result: {
      data: savedDraftData,
      document: { available: true, state: "submission" },
      id: responseId,
      state: "submitted",
      submissionId: completedSubmissionId,
    },
  });
  const originalSubmissionBeforeCorrection =
    await prisma.submission.findUniqueOrThrow({
      select: {
        data: true,
        documentKey: true,
        objectKey: true,
      },
      where: { id: completedSubmissionId },
    });
  const correctionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(correctionEditorResponse.status).toBe(200);
  const correctionEditorConfig =
    (await correctionEditorResponse.json()) as EditorConfigBody;
  const correctionDocumentKey = correctionEditorConfig.config.document.key;
  const correctionCapability =
    correctionEditorConfig.bridge.capabilities["save-correction"];
  if (!correctionCapability || !correctionDocumentKey) {
    throw new Error("The correction editor capabilities were not returned");
  }
  expect(correctionDocumentKey).not.toBe(responseDocumentKey);
  expect(
    correctionEditorConfig.config.editorConfig.plugins.options[pluginGuid]
  ).toMatchObject({
    bridgeId: correctionEditorConfig.bridge.id,
    publicId: formRecord.publicId,
  });
  const correctionClaims = verifyEditorCapability(correctionCapability);
  expect(correctionClaims).toMatchObject({
    action: "save-correction",
    actorId: adminId,
    documentKey: correctionDocumentKey,
    formId,
    leaseId: correctionEditorConfig.bridge.lease.id,
    role: "admin",
    targetId: responseId,
    targetType: "correction",
  });
  const competingCorrectionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${competingAdminBearer}` } }
    )
  );
  expect(competingCorrectionEditorResponse.status).toBe(409);
  expect(await competingCorrectionEditorResponse.json()).toMatchObject({
    error: "editor_in_use",
  });
  const correctionData = {
    ...savedDraftData,
    description_1: "แก้ไขข้อมูลโดยผู้ดูแล",
  };
  const correctionSaveResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction`,
      {
        body: JSON.stringify({
          data: correctionData,
          documentKey: correctionDocumentKey,
          reason: "แก้ไขตามเอกสารต้นฉบับ",
        }),
        headers: capabilityHeaders(correctionCapability),
        method: "POST",
      }
    )
  );
  expect(correctionSaveResponse.status).toBe(202);
  const correctionSaveBody = (await correctionSaveResponse.json()) as {
    operationCapability?: string;
    operationId?: string;
  };
  if (
    !correctionSaveBody.operationCapability ||
    !correctionSaveBody.operationId
  ) {
    throw new Error("The correction operation was not created");
  }
  const correctionOperation = await waitForOperation(
    correctionSaveBody.operationId,
    { "X-Editor-Capability": correctionSaveBody.operationCapability }
  );
  const correctionOperationRecord = await prisma.operation.findUniqueOrThrow({
    select: { result: true },
    where: { id: correctionSaveBody.operationId },
  });
  const correctionResult = correctionOperationRecord.result as
    | { correctionId?: string; revision?: number }
    | undefined;
  expect(correctionOperation).toMatchObject({
    result: { correctionId: expect.any(String), revision: 1 },
    status: "completed",
  });
  if (!correctionResult?.correctionId) {
    throw new Error("The correction result was not returned");
  }
  expect(correctionResult.revision).toBe(1);
  const correctedSubmission = await prisma.submission.findUniqueOrThrow({
    select: { data: true, documentKey: true, objectKey: true },
    where: { id: completedSubmissionId },
  });
  expect(correctedSubmission).toEqual(originalSubmissionBeforeCorrection);
  const correction = await prisma.correction.findFirstOrThrow({
    where: { responseId, revision: 1 },
  });
  expect(String(correctionResult.correctionId)).toBe(correction.id);
  expect(correction).toMatchObject({
    actorId: adminId,
    data: correctionData,
    reason: "แก้ไขตามเอกสารต้นฉบับ",
    responseId,
    revision: 1,
    submissionId: completedSubmissionId,
  });
  expect(await objectExists(correction.objectKey)).toBe(true);
  const correctionReplayClaim = await prisma.callbackClaim.findUnique({
    where: { operationId: correctionSaveBody.operationId },
  });
  if (!correctionReplayClaim) {
    throw new Error("The correction callback claim was not persisted");
  }
  const correctionReplayUserdata = createCallbackUserdata({
    documentKey: correctionDocumentKey,
    expiresAt: Math.floor(correctionReplayClaim.expiresAt.getTime() / 1000),
    operationId: correctionSaveBody.operationId,
    operationType: "save_correction",
  });
  expect(
    createHash("sha256").update(correctionReplayUserdata).digest("hex")
  ).toBe(correctionReplayClaim.tokenDigest);
  const correctionReplayPayload = {
    key: correctionDocumentKey,
    status: 6,
    userdata: correctionReplayUserdata,
  };
  const correctionReplayResponse = await app.handle(
    new Request("http://test.local/onlyoffice/callback", {
      body: JSON.stringify({
        ...correctionReplayPayload,
        token: createOnlyOfficeBodyToken(correctionReplayPayload),
      }),
      headers: {
        Authorization: createOnlyOfficeAuthorization(correctionReplayPayload),
        ...jsonHeaders,
      },
      method: "POST",
    })
  );
  expect(correctionReplayResponse.status).toBe(200);
  expect(await correctionReplayResponse.json()).toEqual({ error: 0 });
  expect(await prisma.correction.count({ where: { responseId } })).toBe(1);
  const correctionHistoryResponse = await app.handle(
    new Request(`http://test.local/api/responses/${responseId}/corrections`, {
      headers: { Authorization: `Bearer ${userBearer}` },
    })
  );
  expect(correctionHistoryResponse.status).toBe(200);
  expect(await correctionHistoryResponse.json()).toMatchObject({
    latestRevision: 1,
    revisions: [
      {
        data: savedDraftData,
        reason: null,
        revision: 0,
      },
      {
        data: correctionData,
        reason: "แก้ไขตามเอกสารต้นฉบับ",
        revision: 1,
      },
    ],
  });
  const latestDataResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/data?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestDataResponse.status).toBe(200);
  const latestDataBody =
    (await latestDataResponse.json()) as typeof dataBody & {
      correction: { reason: string; revision: number } | null;
      revision: number;
    };
  expect(latestDataBody).toMatchObject({
    correction: { reason: "แก้ไขตามเอกสารต้นฉบับ", revision: 1 },
    data: correctionData,
    fields: dataBody.fields,
    revision: 1,
  });
  const latestJsonResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestJsonResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}-revision-1.json"`
  );
  expect(JSON.parse(await latestJsonResponse.text())).toEqual(correctionData);
  const originalJsonAfterCorrection = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/json?revision=original`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(JSON.parse(await originalJsonAfterCorrection.text())).toEqual(
    savedDraftData
  );
  const latestDocxResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/docx?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestDocxResponse.status).toBe(200);
  expect(latestDocxResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}-revision-1.docx"`
  );
  const latestPdfResponse = await app.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf?revision=latest`,
      { headers: { Authorization: `Bearer ${userBearer}` } }
    )
  );
  expect(latestPdfResponse.status).toBe(200);
  expect(latestPdfResponse.headers.get("content-disposition")).toBe(
    `attachment; filename="submission-${completedSubmissionId}-revision-1.pdf"`
  );
  expect(await latestPdfResponse.text()).toBe("%PDF-test");
  const correctedResultsResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}?revision=latest`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(correctedResultsResponse.status).toBe(200);
  expect(await correctedResultsResponse.json()).toMatchObject({
    result: {
      correction: { reason: "แก้ไขตามเอกสารต้นฉบับ", revision: 1 },
      data: correctionData,
      document: { available: true, state: "correction" },
      latestCorrectionNumber: 1,
      revision: 1,
    },
  });
  const correctionAudits = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: {
      actorId: adminId,
      targetId: { in: [responseId, correction.id, completedSubmissionId] },
    },
  });
  for (const audit of correctionAudits) {
    expect(JSON.stringify(audit.safeMetadata)).not.toContain(
      "แก้ไขตามเอกสารต้นฉบับ"
    );
    expect(JSON.stringify(audit.safeMetadata)).not.toContain(
      correction.objectKey
    );
  }
  const externalStatusAfterCorrection = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(await externalStatusAfterCorrection.json()).toMatchObject({
    latestCorrectionNumber: 1,
    status: "submitted",
  });
  const failedCorrectionEditorResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction/editor-config`,
      { headers: { Authorization: `Bearer ${adminBearer}` } }
    )
  );
  expect(failedCorrectionEditorResponse.status).toBe(200);
  const failedCorrectionEditorConfig =
    (await failedCorrectionEditorResponse.json()) as EditorConfigBody;
  const failedCorrectionCapability =
    failedCorrectionEditorConfig.bridge.capabilities["save-correction"];
  if (!failedCorrectionCapability) {
    throw new Error("The failed correction capability was not returned");
  }
  const correctionFailureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.resolve(new TextEncoder().encode("%PDF-test")),
      forceSave: () =>
        Promise.reject(
          new Error("deterministic correction force-save failure")
        ),
    },
  });
  const failedCorrectionSaveResponse = await correctionFailureApp.handle(
    new Request(
      `http://test.local/api/admin/results/${responseId}/correction`,
      {
        body: JSON.stringify({
          data: { description_1: "ควรไม่ถูกบันทึก" },
          documentKey: failedCorrectionEditorConfig.config.document.key,
          reason: "การแก้ไขที่ล้มเหลว",
        }),
        headers: capabilityHeaders(failedCorrectionCapability),
        method: "POST",
      }
    )
  );
  expect(failedCorrectionSaveResponse.status).toBe(202);
  const failedCorrectionSaveBody =
    (await failedCorrectionSaveResponse.json()) as {
      operationCapability?: string;
      operationId?: string;
    };
  if (
    !failedCorrectionSaveBody.operationCapability ||
    !failedCorrectionSaveBody.operationId
  ) {
    throw new Error("The failed correction operation was not created");
  }
  const failedCorrectionOperation = await waitForOperation(
    failedCorrectionSaveBody.operationId,
    { "X-Editor-Capability": failedCorrectionSaveBody.operationCapability }
  );
  expect(failedCorrectionOperation).toMatchObject({
    error: "force_save_failed",
    status: "failed",
  });
  expect(await prisma.correction.count({ where: { responseId } })).toBe(1);
  expect(
    await prisma.correction.findFirstOrThrow({
      orderBy: { revision: "desc" },
      where: { responseId },
    })
  ).toMatchObject({ data: correctionData, revision: 1 });
  const failedCorrectionReleaseResponse = await app.handle(
    new Request(
      `http://test.local/api/editor-leases/${failedCorrectionEditorConfig.bridge.lease.id}`,
      { headers: { Authorization: `Bearer ${adminBearer}` }, method: "DELETE" }
    )
  );
  expect(failedCorrectionReleaseResponse.status).toBe(200);
  const exportAudits = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: {
      action: "export_response",
      actorId: adminId,
      targetId: completedSubmissionId,
      targetType: "submission",
    },
  });
  expect(exportAudits.map((audit) => audit.safeMetadata)).toEqual([
    { format: "json", revision: 0, state: "submitted" },
    { format: "docx", revision: 0, state: "submitted" },
    { format: "pdf", revision: 0, state: "submitted" },
  ]);
  const stableSubmissionBeforeConversion = await prisma.submission.findUnique({
    select: {
      data: true,
      documentKey: true,
      objectKey: true,
    },
    where: { id: completedSubmissionId },
  });
  const conversionFailureApp = createApp({
    onlyOffice: {
      convertDocxToPdf: () =>
        Promise.reject(new Error("deterministic conversion failure")),
      forceSave: () => Promise.resolve(false),
    },
  });
  const conversionFailureResponse = await conversionFailureApp.handle(
    new Request(
      `http://test.local/api/submissions/${completedSubmissionId}/pdf`,
      {
        headers: { Authorization: `Bearer ${userBearer}` },
      }
    )
  );
  expect(conversionFailureResponse.status).toBe(500);
  const failedPdfAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "export_response",
      actorId: userId,
      outcome: "failure",
      targetId: completedSubmissionId,
      targetType: "submission",
    },
  });
  expect(failedPdfAudit.safeMetadata).toEqual({
    errorCode: "pdf_conversion_failed",
    format: "pdf",
    revision: 0,
    state: "submitted",
  });
  expect(
    await prisma.submission.findUnique({
      select: {
        data: true,
        documentKey: true,
        objectKey: true,
      },
      where: { id: completedSubmissionId },
    })
  ).toEqual(stableSubmissionBeforeConversion);
  if (!stableSubmissionBeforeConversion) {
    throw new Error("The stable submission was not found");
  }
  expect(await readObject(stableSubmissionBeforeConversion.objectKey)).toEqual(
    submissionDocument
  );
  const unarchiveResponse = await app.handle(
    new Request(`http://test.local/api/admin/forms/${formPublicId}`, {
      body: JSON.stringify({ status: "published" }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${adminBearer}` },
      method: "PATCH",
    })
  );
  expect(unarchiveResponse.status).toBe(200);
  expect(await unarchiveResponse.json()).toMatchObject({
    form: {
      publicId: formPublicId,
      status: "published",
      version: 1,
    },
  });
  const archivedHandoffReference = `ticket-17-archived-${crypto.randomUUID()}`;
  const archivedHandoffCreate = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: archivedNoResponseEmail,
        externalReference: archivedHandoffReference,
        publicId: formPublicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(archivedHandoffCreate.status).toBe(200);
  const archivedHandoffCode = (
    (await archivedHandoffCreate.json()) as { code: string }
  ).code;
  const archivedHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: archivedHandoffCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    })
  );
  expect(archivedHandoffLaunch.status).toBe(303);
  const archivedPendingCookie = archivedHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!archivedPendingCookie) {
    throw new Error("The archived handoff did not set a pending claim cookie");
  }
  const unarchivedStartResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}/start`, {
      headers: {
        Authorization: `Bearer ${archivedNoResponseBearer}`,
        Cookie: archivedPendingCookie,
      },
      method: "POST",
    })
  );
  expect(unarchivedStartResponse.status).toBe(200);
  const unarchivedStartBody = (await unarchivedStartResponse.json()) as {
    response?: { id?: string };
  };
  const discardedResponseId = unarchivedStartBody.response?.id;
  if (!discardedResponseId) {
    throw new Error("The discard fixture was not created");
  }
  const discardResponseBefore = await prisma.response.findUnique({
    select: { draftObjectKey: true },
    where: { id: discardedResponseId },
  });
  if (!discardResponseBefore?.draftObjectKey) {
    throw new Error("The discard Draft document was not created");
  }
  const discardEditorConfigResponse = await app.handle(
    new Request(
      `http://test.local/api/forms/${formPublicId}/editor-config?responseId=${discardedResponseId}&action=draft`,
      { headers: { Authorization: `Bearer ${archivedNoResponseBearer}` } }
    )
  );
  expect(discardEditorConfigResponse.status).toBe(200);
  const unauthorizedDiscard = await app.handle(
    new Request(`http://test.local/api/responses/${discardedResponseId}`, {
      headers: { Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(unauthorizedDiscard.status).toBe(403);
  const discardResponse = await app.handle(
    new Request(`http://test.local/api/responses/${discardedResponseId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "DELETE",
    })
  );
  expect(discardResponse.status).toBe(200);
  expect(await discardResponse.json()).toEqual({ discarded: true });
  expect(
    await prisma.response.findUnique({ where: { id: discardedResponseId } })
  ).toBeNull();
  expect(
    await prisma.editorLease.count({
      where: {
        targetId: discardedResponseId,
        targetType: "response",
      },
    })
  ).toBe(0);
  expect(
    await prisma.operation.count({ where: { responseId: discardedResponseId } })
  ).toBe(0);
  expect(await objectExists(discardResponseBefore.draftObjectKey)).toBe(false);
  const repeatedDiscardResponse = await app.handle(
    new Request(`http://test.local/api/responses/${discardedResponseId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "DELETE",
    })
  );
  expect(repeatedDiscardResponse.status).toBe(200);
  expect(await repeatedDiscardResponse.json()).toEqual({ discarded: true });
  const responsesAfterDiscard = await app.handle(
    new Request("http://test.local/api/responses/me", {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
    })
  );
  expect(await responsesAfterDiscard.json()).toEqual({ responses: [] });
  const restartHandoffReference = `ticket-17-restart-${crypto.randomUUID()}`;
  const restartHandoffCreate = await app.handle(
    new Request("http://test.local/api/integrations/prefill/handoffs", {
      body: JSON.stringify({
        email: archivedNoResponseEmail,
        externalReference: restartHandoffReference,
        publicId: formPublicId,
        values: handoffCandidateValues,
      }),
      headers: {
        ...jsonHeaders,
        "X-Prefill-Handoff-Secret": prefillHandoffSecret,
      },
      method: "POST",
    })
  );
  expect(restartHandoffCreate.status).toBe(200);
  const restartHandoffCode = (
    (await restartHandoffCreate.json()) as { code: string }
  ).code;
  const restartHandoffLaunch = await app.handle(
    new Request("http://test.local/prefill/handoff", {
      body: new URLSearchParams({ code: restartHandoffCode }),
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      method: "POST",
    })
  );
  expect(restartHandoffLaunch.status).toBe(303);
  const restartPendingCookie = restartHandoffLaunch.headers
    .get("set-cookie")
    ?.split(";", 1)[0];
  if (!restartPendingCookie) {
    throw new Error("The restart handoff did not set a pending claim cookie");
  }
  const restartAfterDiscard = await app.handle(
    new Request(`http://test.local/api/forms/${formPublicId}/start`, {
      headers: {
        Authorization: `Bearer ${archivedNoResponseBearer}`,
        Cookie: restartPendingCookie,
      },
      method: "POST",
    })
  );
  expect(restartAfterDiscard.status).toBe(200);
  const restartAfterDiscardBody = (await restartAfterDiscard.json()) as {
    response?: { id?: string };
  };
  const restartedResponseId = restartAfterDiscardBody.response?.id;
  if (!restartedResponseId) {
    throw new Error("The response was not restartable after discard");
  }
  expect(restartedResponseId).not.toBe(discardedResponseId);
  const discardRestartResponse = await app.handle(
    new Request(`http://test.local/api/responses/${restartedResponseId}`, {
      headers: { Authorization: `Bearer ${archivedNoResponseBearer}` },
      method: "DELETE",
    })
  );
  expect(discardRestartResponse.status).toBe(200);
  const unarchivedFormListResponse = await app.handle(
    new Request("http://test.local/api/admin/forms", {
      headers: { Authorization: `Bearer ${adminBearer}` },
    })
  );
  const unarchivedFormListBody = (await unarchivedFormListResponse.json()) as {
    forms?: {
      activeDraftCount: number;
      publicId: string;
      status: string;
      submissionCount: number;
    }[];
  };
  expect(
    unarchivedFormListBody.forms?.find((form) => form.publicId === formPublicId)
  ).toMatchObject({
    activeDraftCount: 0,
    publicId: formPublicId,
    status: "published",
    submissionCount: 1,
  });
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "unarchive_form",
        outcome: "success",
        targetId: formPublicId,
      },
    })
  ).toMatchObject({ actorId: adminId, targetId: formPublicId });
  expect(
    await prisma.publishedTemplate.findUnique({
      include: {
        manifest: { include: { fields: { orderBy: { tag: "asc" } } } },
        prefillConfiguration: {
          include: { fields: { orderBy: { tag: "asc" } } },
        },
      },
      where: { formId },
    })
  ).toEqual(publishedContractBefore);
  const staleOperationId = crypto.randomUUID();
  const staleStagingObjectKey = objectKey(
    "operations",
    staleOperationId,
    "staged.docx"
  );
  const staleFinalObjectKey = objectKey(
    "operations",
    staleOperationId,
    "final.docx"
  );
  const staleOperationUserdata = createCallbackUserdata({
    documentKey: templateDraft.documentKey,
    expiresAt: Math.floor(Date.now() / 1000) - 1,
    operationId: staleOperationId,
    operationType: "save_template_draft",
  });
  await Promise.all([
    putObject(staleStagingObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
    putObject(staleFinalObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
  ]);
  await prisma.operation.create({
    data: {
      actorId: adminId,
      documentKey: templateDraft.documentKey,
      errorCode: null,
      formId,
      id: staleOperationId,
      metadata: {
        action: "save-template",
        finalObjectKey: staleFinalObjectKey,
        formId,
        stagedObjectKey: staleStagingObjectKey,
      },
      ownerUserId: adminId,
      stagingObjectKey: staleStagingObjectKey,
      status: "processing",
      targetId: templateDraft.id,
      targetType: "template_draft",
      type: "save_template_draft",
      updatedAt: new Date(0),
    },
  });
  await persistCallbackClaim({
    expiresAt: new Date(0),
    operationId: staleOperationId,
    userdata: staleOperationUserdata,
  });
  const staleOperationCapability = createEditorCapability({
    action: "poll-operation",
    actorId: adminId,
    documentKey: templateDraft.documentKey,
    formId,
    operationId: staleOperationId,
    role: "admin",
    targetId: templateDraft.id,
    targetType: "template-draft",
  });

  const publishedTemplate = await prisma.publishedTemplate.findUnique({
    select: { id: true, version: true },
    where: { formId },
  });
  if (!publishedTemplate) {
    throw new Error("The published template was not found");
  }
  const staleResponseId = crypto.randomUUID();
  const staleResponseDocumentKey = `response-${staleResponseId}-${crypto.randomUUID()}`;
  const staleResponseObjectKey = objectKey(
    "responses",
    staleResponseId,
    "draft",
    crypto.randomUUID(),
    "docx"
  );
  const staleSubmitOperationId = crypto.randomUUID();
  const staleSubmitStagingObjectKey = objectKey(
    "operations",
    staleSubmitOperationId,
    "staged.docx"
  );
  const staleSubmitFinalObjectKey = objectKey(
    "operations",
    staleSubmitOperationId,
    "final.docx"
  );
  const staleSubmissionId = crypto.randomUUID();
  const staleSubmitUserdata = createCallbackUserdata({
    documentKey: staleResponseDocumentKey,
    expiresAt: Math.floor(Date.now() / 1000) - 1,
    operationId: staleSubmitOperationId,
    operationType: "submit_response",
  });
  await Promise.all([
    putObject(staleResponseObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
    putObject(staleSubmitStagingObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
    putObject(staleSubmitFinalObjectKey, sourceDocument, DOCX_CONTENT_TYPE),
  ]);
  await prisma.response.create({
    data: {
      draftData: {},
      draftDocumentKey: staleResponseDocumentKey,
      draftObjectKey: staleResponseObjectKey,
      form: { connect: { id: formId } },
      id: staleResponseId,
      owner: { connect: { id: otherUser.id } },
      publishedTemplate: { connect: { id: publishedTemplate.id } },
      publishedVersion: publishedTemplate.version,
      status: "submitting",
    },
  });
  await prisma.operation.create({
    data: {
      actorId: otherUser.id,
      documentKey: staleResponseDocumentKey,
      errorCode: null,
      formId,
      id: staleSubmitOperationId,
      metadata: {
        action: "submit",
        finalObjectKey: staleSubmitFinalObjectKey,
        formId,
        responseId: staleResponseId,
        stagedObjectKey: staleSubmitStagingObjectKey,
        submissionDocumentKey: `submission-${staleSubmissionId}-${crypto.randomUUID()}`,
        submissionId: staleSubmissionId,
      },
      ownerUserId: otherUser.id,
      responseId: staleResponseId,
      stagingObjectKey: staleSubmitStagingObjectKey,
      status: "processing",
      targetId: staleResponseId,
      targetType: "response",
      type: "submit_response",
      updatedAt: new Date(0),
    },
  });
  await persistCallbackClaim({
    expiresAt: new Date(0),
    operationId: staleSubmitOperationId,
    userdata: staleSubmitUserdata,
  });

  const expiredLease = await prisma.editorLease.findUnique({
    where: {
      targetType_targetId: {
        targetId: templateDraft.id,
        targetType: "template_draft",
      },
    },
  });
  if (!expiredLease) {
    throw new Error("The active template lease was not found");
  }
  await prisma.editorLease.update({
    data: { createdAt: new Date(0), expiresAt: new Date(1) },
    where: { id: expiredLease.id },
  });
  await prisma.operation.update({
    data: { updatedAt: new Date() },
    where: { id: staleOperationId },
  });
  const cleanupIntentObjectKey = objectKey(
    "cleanup-intents",
    crypto.randomUUID(),
    "orphan.docx"
  );
  await putObject(
    cleanupIntentObjectKey,
    docxFixture("ticket-08-cleanup-intent"),
    DOCX_CONTENT_TYPE
  );
  await prisma.objectCleanupIntent.create({
    data: { objectKey: cleanupIntentObjectKey },
  });
  expect(await objectExists(cleanupIntentObjectKey)).toBe(true);
  const inFlightCleanupObjectKey = objectKey(
    "cleanup-intents",
    crypto.randomUUID(),
    "in-flight.docx"
  );
  await putObject(
    inFlightCleanupObjectKey,
    docxFixture("ticket-08-in-flight-cleanup"),
    DOCX_CONTENT_TYPE
  );
  await prisma.objectCleanupIntent.create({
    data: {
      cleanupAfter: new Date(Date.now() + 60_000),
      objectKey: inFlightCleanupObjectKey,
    },
  });
  await reconcileRecoverableState();
  expect(
    await prisma.editorLease.findUnique({ where: { id: expiredLease.id } })
  ).not.toBeNull();
  expect(await objectExists(cleanupIntentObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: cleanupIntentObjectKey },
    })
  ).toBeNull();
  expect(await objectExists(inFlightCleanupObjectKey)).toBe(true);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: inFlightCleanupObjectKey },
    })
  ).not.toBeNull();
  await prisma.objectCleanupIntent.update({
    data: { cleanupAfter: new Date(0) },
    where: { objectKey: inFlightCleanupObjectKey },
  });
  await prisma.operation.update({
    data: { updatedAt: new Date(0) },
    where: { id: staleOperationId },
  });
  await reconcileRecoverableState();
  expect(await objectExists(inFlightCleanupObjectKey)).toBe(false);
  expect(
    await prisma.objectCleanupIntent.findUnique({
      where: { objectKey: inFlightCleanupObjectKey },
    })
  ).toBeNull();

  expect(
    await prisma.operation.findUnique({
      select: { errorCode: true, status: true },
      where: { id: staleOperationId },
    })
  ).toEqual({ errorCode: "operation_timeout", status: "failed" });
  expect(
    await prisma.operation.findUnique({
      select: { errorCode: true, status: true },
      where: { id: staleSubmitOperationId },
    })
  ).toEqual({ errorCode: "operation_timeout", status: "failed" });
  expect(
    await prisma.response.findUnique({
      select: {
        draftDocumentKey: true,
        draftObjectKey: true,
        status: true,
      },
      where: { id: staleResponseId },
    })
  ).toEqual({
    draftDocumentKey: staleResponseDocumentKey,
    draftObjectKey: staleResponseObjectKey,
    status: "draft",
  });
  expect(
    await prisma.editorLease.findUnique({ where: { id: expiredLease.id } })
  ).toBeNull();
  expect(
    await prisma.callbackClaim.findUnique({
      where: { operationId: staleOperationId },
    })
  ).toBeNull();
  expect(
    await prisma.callbackClaim.findUnique({
      where: { operationId: staleSubmitOperationId },
    })
  ).toBeNull();
  expect(await objectExists(staleStagingObjectKey)).toBe(false);
  expect(await objectExists(staleFinalObjectKey)).toBe(false);
  expect(await objectExists(staleSubmitStagingObjectKey)).toBe(false);
  expect(await objectExists(staleSubmitFinalObjectKey)).toBe(false);
  expect(await objectExists(staleResponseObjectKey)).toBe(true);
  expect(await objectExists(templateDraft.objectKey)).toBe(true);
  expect(await readObject(templateDraft.objectKey)).toEqual(sourceDocument);

  const reconciledOperationResponse = await app.handle(
    new Request(`http://test.local/api/operations/${staleOperationId}`, {
      headers: { "X-Editor-Capability": staleOperationCapability },
    })
  );
  expect(reconciledOperationResponse.status).toBe(200);
  expect(await reconciledOperationResponse.json()).toMatchObject({
    operation: {
      error: "operation_timeout",
      id: staleOperationId,
      status: "failed",
    },
  });
  const reclaimedAfterReconcileResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/forms/${formPublicId}/editor-config`,
      {
        headers: { Authorization: `Bearer ${adminBearer}` },
      }
    )
  );
  expect(reclaimedAfterReconcileResponse.status).toBe(409);
  expect(await reclaimedAfterReconcileResponse.json()).toMatchObject({
    error: "published_immutable",
  });

  const callbackFormId = crypto.randomUUID();
  const callbackFormPublicId = crypto.randomUUID().replaceAll("-", "");
  const callbackTemplateDocumentKey = `callback-template-${crypto.randomUUID()}`;
  const callbackTemplateObjectKey = objectKey(
    "forms",
    callbackFormId,
    "template-draft",
    "callback.docx"
  );
  await putObject(callbackTemplateObjectKey, sourceDocument, DOCX_CONTENT_TYPE);
  await prisma.form.create({
    data: {
      createdBy: adminId,
      description: "Callback operation test form",
      id: callbackFormId,
      publicId: callbackFormPublicId,
      templateDraft: {
        create: {
          contentHash: createHash("sha256")
            .update(sourceDocument)
            .digest("hex"),
          documentKey: callbackTemplateDocumentKey,
          id: crypto.randomUUID(),
          objectKey: callbackTemplateObjectKey,
        },
      },
      title: "Callback operation test form",
    },
  });
  const callbackTemplateDraft = await prisma.templateDraft.findUnique({
    where: { formId: callbackFormId },
  });
  if (!callbackTemplateDraft) {
    throw new Error("The callback template draft was not created");
  }

  const callbackDocument = docxFixture(
    `ticket-08-callback-${crypto.randomUUID()}`
  );
  const externalCallbackDocument = docxXmlFixture({
    additionalParts: {
      "word/_rels/document.xml.rels": strToU8(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="external" Target="http://127.0.0.1:80/" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/></Relationships>'
      ),
    },
    document:
      '<?xml version="1.0"?><word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><word:body/></word:document>',
  });
  const callbackMaximumBytes = Math.max(
    callbackDocument.byteLength,
    externalCallbackDocument.byteLength
  );
  const callbackDownloadPaths = new Set<string>();
  const callbackDocumentServer = Bun.serve({
    fetch(request) {
      const url = new URL(request.url);
      if (
        !verifyOnlyOfficeAuthorization(request.headers.get("authorization"), {
          url: request.url,
        })
      ) {
        return Response.json({ error: "unauthorized" }, { status: 401 });
      }
      callbackDownloadPaths.add(url.pathname);
      if (url.pathname === "/redirect.docx") {
        return new Response(null, {
          headers: { Location: `${url.origin}/ok.docx` },
          status: 302,
        });
      }
      if (url.pathname === "/large.docx") {
        return new Response(new Uint8Array(callbackMaximumBytes + 1));
      }
      if (url.pathname === "/malformed.docx") {
        return new Response("not a DOCX package");
      }
      if (url.pathname === "/external-relationship.docx") {
        return new Response(externalCallbackDocument, {
          headers: { "Content-Type": DOCX_CONTENT_TYPE },
        });
      }
      return new Response(callbackDocument, {
        headers: { "Content-Type": DOCX_CONTENT_TYPE },
      });
    },
    port: 0,
  });
  try {
    const callbackOrigin = callbackDocumentServer.url.origin;
    const callbackApp = createApp({
      onlyOffice: {
        convertDocxToPdf: () =>
          Promise.resolve(new TextEncoder().encode("%PDF-test")),
        forceSave: () => Promise.resolve(false),
      },
      onlyOfficeCallbackMaxBytes: callbackMaximumBytes,
      onlyOfficeCallbackOrigins: [callbackOrigin],
    });
    const callbackAppReplica = createApp({
      onlyOffice: {
        convertDocxToPdf: () =>
          Promise.resolve(new TextEncoder().encode("%PDF-test")),
        forceSave: () => Promise.resolve(false),
      },
      onlyOfficeCallbackMaxBytes: callbackMaximumBytes,
      onlyOfficeCallbackOrigins: [callbackOrigin],
    });
    const createCallbackOperation =
      async (): Promise<CallbackOperationFixture> => {
        const id = crypto.randomUUID();
        const stagedObjectKey = objectKey(
          "operations",
          id,
          "callback-staged.docx"
        );
        const finalObjectKey = objectKey(
          "operations",
          id,
          "callback-final.docx"
        );
        const userdata = createCallbackUserdata({
          documentKey: callbackTemplateDraft.documentKey,
          operationId: id,
          operationType: "save_template_draft",
        });
        await prisma.operation.create({
          data: {
            actorId: adminId,
            documentKey: callbackTemplateDraft.documentKey,
            errorCode: null,
            formId: callbackFormId,
            id,
            metadata: {
              action: "save-template",
              finalObjectKey,
              formId: callbackFormId,
              stagedObjectKey,
            },
            ownerUserId: adminId,
            stagingObjectKey: stagedObjectKey,
            status: "processing",
            targetId: id,
            targetType: "template_draft",
            type: "save_template_draft",
          },
        });
        await persistCallbackClaim({ operationId: id, userdata });
        return {
          documentKey: callbackTemplateDraft.documentKey,
          finalObjectKey,
          id,
          userdata,
        };
      };
    const postCallback = (
      payload: Record<string, unknown>,
      authorization = createOnlyOfficeAuthorization(payload),
      bodyToken = createOnlyOfficeBodyToken(payload),
      callbackHandler = callbackApp
    ): Promise<Response> =>
      callbackHandler.handle(
        new Request("http://test.local/onlyoffice/callback", {
          body: JSON.stringify({ ...payload, token: bodyToken }),
          headers: {
            Authorization: authorization,
            ...jsonHeaders,
          },
          method: "POST",
        })
      );

    const trustOperation = await createCallbackOperation();
    const persistedTrustClaim = await prisma.callbackClaim.findUnique({
      where: { operationId: trustOperation.id },
    });
    if (!persistedTrustClaim) {
      throw new Error("The callback claim was not persisted");
    }
    expect(persistedTrustClaim.tokenDigest).toMatch(/^[0-9a-f]{64}$/u);
    expect(persistedTrustClaim.tokenDigest).toBe(
      createHash("sha256").update(trustOperation.userdata).digest("hex")
    );
    const trustedPayload = callbackPayload(
      trustOperation,
      `${callbackOrigin}/ok.docx`
    );
    const invalidJwtResponse = await postCallback(
      trustedPayload,
      tamperAuthorization(createOnlyOfficeAuthorization(trustedPayload))
    );
    expect(invalidJwtResponse.status).toBe(401);
    const invalidBodyTokenResponse = await postCallback(
      trustedPayload,
      createOnlyOfficeAuthorization(trustedPayload),
      `${createOnlyOfficeBodyToken(trustedPayload)}x`
    );
    expect(invalidBodyTokenResponse.status).toBe(401);
    const ordinaryCallbackResponse = await postCallback({
      actions: [],
      key: trustOperation.documentKey,
      status: 1,
    });
    expect(await ordinaryCallbackResponse.json()).toEqual({ error: 0 });
    const browserSessionBoundaryResponse = await postCallback(
      trustedPayload,
      `Bearer ${publishCapability}`
    );
    expect(browserSessionBoundaryResponse.status).toBe(401);

    const wrongKeyPayload = {
      ...trustedPayload,
      key: `wrong-${trustOperation.documentKey}`,
    };
    const wrongKeyResponse = await postCallback(wrongKeyPayload);
    expect(await wrongKeyResponse.json()).toEqual({ error: 1 });
    const wrongOperationPayload = {
      ...trustedPayload,
      userdata: createCallbackUserdata({
        documentKey: trustOperation.documentKey,
        operationId: crypto.randomUUID(),
        operationType: "save_template_draft",
      }),
    };
    const wrongOperationResponse = await postCallback(wrongOperationPayload);
    expect(await wrongOperationResponse.json()).toEqual({ error: 1 });
    const wrongOperationTypePayload = {
      ...trustedPayload,
      userdata: createCallbackUserdata({
        documentKey: trustOperation.documentKey,
        operationId: trustOperation.id,
        operationType: "submit_response",
      }),
    };
    const wrongOperationTypeResponse = await postCallback(
      wrongOperationTypePayload
    );
    expect(await wrongOperationTypeResponse.json()).toEqual({ error: 1 });
    const expiredCallbackPayload = {
      ...trustedPayload,
      userdata: createCallbackUserdata({
        documentKey: trustOperation.documentKey,
        expiresAt: Math.floor(Date.now() / 1000) - 1,
        operationId: trustOperation.id,
        operationType: "save_template_draft",
      }),
    };
    const expiredCallbackResponse = await postCallback(expiredCallbackPayload);
    expect(await expiredCallbackResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: trustOperation.id },
      })
    ).toEqual({ status: "processing" });

    const oversizedCallbackPayload = {
      padding: "x".repeat(65 * 1024),
      status: 6,
    };
    const oversizedCallbackResponse = await postCallback(
      oversizedCallbackPayload
    );
    expect(oversizedCallbackResponse.status).toBe(413);

    const forbiddenOriginOperation = await createCallbackOperation();
    const forbiddenOriginResponse = await postCallback(
      callbackPayload(
        forbiddenOriginOperation,
        "https://attacker.example/document.docx"
      )
    );
    expect(await forbiddenOriginResponse.json()).toEqual({ error: 0 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: forbiddenOriginOperation.id },
      })
    ).toEqual({ status: "failed" });

    const redirectOperation = await createCallbackOperation();
    const redirectResponse = await postCallback(
      callbackPayload(redirectOperation, `${callbackOrigin}/redirect.docx`)
    );
    expect(await redirectResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: redirectOperation.id },
      })
    ).toEqual({ status: "failed" });

    const oversizedDocumentOperation = await createCallbackOperation();
    const oversizedDocumentResponse = await postCallback(
      callbackPayload(
        oversizedDocumentOperation,
        `${callbackOrigin}/large.docx`
      )
    );
    expect(await oversizedDocumentResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: oversizedDocumentOperation.id },
      })
    ).toEqual({ status: "failed" });
    const templateBeforeMalformedSave = await prisma.templateDraft.findUnique({
      select: { documentKey: true, objectKey: true },
      where: { formId: callbackFormId },
    });
    if (!templateBeforeMalformedSave) {
      throw new Error("The Template Draft rollback baseline was not found");
    }
    const templateBytesBeforeMalformedSave = await readObject(
      templateBeforeMalformedSave.objectKey
    );
    const malformedSaveOperation = await createCallbackOperation();
    const malformedSaveResponse = await postCallback(
      callbackPayload(
        malformedSaveOperation,
        `${callbackOrigin}/malformed.docx`
      )
    );
    expect(await malformedSaveResponse.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: malformedSaveOperation.id },
      })
    ).toEqual({ status: "failed" });
    expect(
      await prisma.templateDraft.findUnique({
        select: { documentKey: true, objectKey: true },
        where: { formId: callbackFormId },
      })
    ).toEqual(templateBeforeMalformedSave);
    expect(await readObject(templateBeforeMalformedSave.objectKey)).toEqual(
      templateBytesBeforeMalformedSave
    );
    const externalResponseId = crypto.randomUUID();
    const externalResponseDocumentKey = `external-response-${crypto.randomUUID()}`;
    const externalResponseObjectKey = objectKey(
      "responses",
      externalResponseId,
      "draft",
      "baseline.docx"
    );
    await putObject(
      externalResponseObjectKey,
      callbackDocument,
      DOCX_CONTENT_TYPE
    );
    await prisma.response.create({
      data: {
        draftData: {},
        draftDocumentKey: externalResponseDocumentKey,
        draftObjectKey: externalResponseObjectKey,
        formId,
        id: externalResponseId,
        publishedTemplateId: publishedManifestRecord.id,
        publishedVersion: publishedManifestRecord.version,
        status: "draft",
        userId: adminId,
      },
    });
    const externalResponseOperationId = crypto.randomUUID();
    const externalResponseStagedObjectKey = objectKey(
      "operations",
      externalResponseOperationId,
      "external-response-staged.docx"
    );
    const externalResponseFinalObjectKey = objectKey(
      "operations",
      externalResponseOperationId,
      "external-response-final.docx"
    );
    const externalResponseUserdata = createCallbackUserdata({
      documentKey: externalResponseDocumentKey,
      operationId: externalResponseOperationId,
      operationType: "save_draft",
    });
    await prisma.operation.create({
      data: {
        actorId: adminId,
        documentKey: externalResponseDocumentKey,
        errorCode: null,
        formId,
        id: externalResponseOperationId,
        metadata: {
          action: "save-draft",
          data: {},
          finalObjectKey: externalResponseFinalObjectKey,
          formId,
          publicId: formPublicId,
          responseId: externalResponseId,
          stagedObjectKey: externalResponseStagedObjectKey,
        },
        ownerUserId: adminId,
        responseId: externalResponseId,
        stagingObjectKey: externalResponseStagedObjectKey,
        status: "processing",
        targetId: externalResponseId,
        targetType: "response",
        type: "save_draft",
      },
    });
    await persistCallbackClaim({
      operationId: externalResponseOperationId,
      userdata: externalResponseUserdata,
    });
    const externalResponseOperation: CallbackOperationFixture = {
      documentKey: externalResponseDocumentKey,
      finalObjectKey: externalResponseFinalObjectKey,
      id: externalResponseOperationId,
      userdata: externalResponseUserdata,
    };
    const externalResponseBefore = await prisma.response.findUnique({
      select: {
        draftData: true,
        draftDocumentKey: true,
        draftObjectKey: true,
        status: true,
      },
      where: { id: externalResponseId },
    });
    const externalResponseCallback = await postCallback(
      callbackPayload(
        externalResponseOperation,
        `${callbackOrigin}/external-relationship.docx`
      )
    );
    expect(await externalResponseCallback.json()).toEqual({ error: 1 });
    expect(
      await prisma.operation.findUnique({
        select: { errorCode: true, status: true },
        where: { id: externalResponseOperationId },
      })
    ).toEqual({
      errorCode: "invalid_template",
      status: "failed",
    });
    expect(
      await prisma.response.findUnique({
        select: {
          draftData: true,
          draftDocumentKey: true,
          draftObjectKey: true,
          status: true,
        },
        where: { id: externalResponseId },
      })
    ).toEqual(externalResponseBefore);
    expect(await objectExists(externalResponseObjectKey)).toBe(true);
    expect(await objectExists(externalResponseStagedObjectKey)).toBe(false);
    expect(await objectExists(externalResponseFinalObjectKey)).toBe(false);
    await prisma.operation.delete({
      where: { id: externalResponseOperationId },
    });
    await prisma.response.delete({ where: { id: externalResponseId } });

    const validCallbackResponse = await postCallback(trustedPayload);
    expect(await validCallbackResponse.json()).toEqual({ error: 0 });
    expect(
      await prisma.operation.findUnique({
        select: { status: true },
        where: { id: trustOperation.id },
      })
    ).toEqual({ status: "completed" });
    expect(await readObject(trustOperation.finalObjectKey)).toEqual(
      callbackDocument
    );
    const consumedTrustClaim = await prisma.callbackClaim.findUnique({
      where: { operationId: trustOperation.id },
    });
    expect(consumedTrustClaim?.consumedAt).not.toBeNull();
    const changedEditorResponse = await callbackApp.handle(
      new Request(
        `http://test.local/api/admin/forms/${callbackFormPublicId}/editor-config`,
        { headers: { Authorization: `Bearer ${adminBearer}` } }
      )
    );
    expect(changedEditorResponse.status).toBe(200);
    const changedEditor =
      (await changedEditorResponse.json()) as EditorConfigBody;
    expect(JSON.stringify(changedEditor)).not.toContain(callbackFormId);
    const changedDocumentUrl = changedEditor.config.document.url;
    const changedDocumentResponse = await callbackApp.handle(
      new Request(changedDocumentUrl, {
        headers: {
          Authorization: createOnlyOfficeAuthorization({
            url: changedDocumentUrl,
          }),
        },
      })
    );
    expect(changedDocumentResponse.status).toBe(200);
    expect(Buffer.from(await changedDocumentResponse.arrayBuffer())).toEqual(
      Buffer.from(callbackDocument)
    );

    const concurrentCallbackOperation = await createCallbackOperation();
    const concurrentCallbackPayload = callbackPayload(
      concurrentCallbackOperation,
      `${callbackOrigin}/ok.docx`
    );
    const [concurrentCallbackA, concurrentCallbackB] = await Promise.all([
      postCallback(concurrentCallbackPayload),
      postCallback(
        concurrentCallbackPayload,
        undefined,
        undefined,
        callbackAppReplica
      ),
    ]);
    expect(await concurrentCallbackA.json()).toEqual({ error: 0 });
    expect(await concurrentCallbackB.json()).toEqual({ error: 0 });
    const concurrentCallbackState = await prisma.operation.findUnique({
      select: { result: true, status: true },
      where: { id: concurrentCallbackOperation.id },
    });
    expect(concurrentCallbackState).toMatchObject({ status: "completed" });
    const concurrentCallbackClaim = await prisma.callbackClaim.findUnique({
      where: { operationId: concurrentCallbackOperation.id },
    });
    expect(concurrentCallbackClaim?.consumedAt).not.toBeNull();
    expect(
      await readObject(concurrentCallbackOperation.finalObjectKey)
    ).toEqual(callbackDocument);
    const replayedCallbackResponse = await postCallback(
      concurrentCallbackPayload
    );
    expect(await replayedCallbackResponse.json()).toEqual({ error: 0 });
    expect(
      await prisma.operation.findUnique({
        select: { result: true, status: true },
        where: { id: concurrentCallbackOperation.id },
      })
    ).toEqual(concurrentCallbackState);

    expect(callbackDownloadPaths).toEqual(
      new Set([
        "/external-relationship.docx",
        "/large.docx",
        "/malformed.docx",
        "/ok.docx",
        "/redirect.docx",
      ])
    );
  } finally {
    callbackDocumentServer.stop(true);
  }

  const passwordEmail = `ticket-04-password-${crypto.randomUUID()}@example.com`;
  const currentPassword = "Ticket04-current-password";
  const passwordUser = await createCredentialFixture({
    email: passwordEmail,
    mustChangePassword: true,
    name: "Ticket 04 Password User",
    password: currentPassword,
  });
  const mandatoryToken = await bearerFor(
    passwordEmail,
    currentPassword,
    `ticket-04-password-${crypto.randomUUID()}`
  );
  const mandatorySessionResponse = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${mandatoryToken}` },
    })
  );
  expect(mandatorySessionResponse.status).toBe(200);
  expect(await mandatorySessionResponse.json()).toMatchObject({
    user: { mustChangePassword: true },
  });
  const restrictedProductResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}`, {
      headers: { Authorization: `Bearer ${mandatoryToken}` },
    })
  );
  expect(restrictedProductResponse.status).toBe(403);
  expect(await restrictedProductResponse.json()).toMatchObject({
    error: "password_change_required",
  });

  const replacePassword = (
    token: string,
    oldPassword: string,
    newPassword: string
  ): Promise<Response> =>
    app.handle(
      new Request("http://test.local/api/account/password", {
        body: JSON.stringify({
          currentPassword: oldPassword,
          newPassword,
        }),
        headers: { ...jsonHeaders, Authorization: `Bearer ${token}` },
        method: "POST",
      })
    );

  const tooShortResponse = await replacePassword(
    mandatoryToken,
    currentPassword,
    "x".repeat(11)
  );
  expect(tooShortResponse.status).toBe(400);
  expect(await tooShortResponse.json()).toMatchObject({
    error: "password_too_short",
  });

  const spacesPassword = " ".repeat(12);
  const spacesReplacementResponse = await replacePassword(
    mandatoryToken,
    currentPassword,
    spacesPassword
  );
  expect(spacesReplacementResponse.status).toBe(200);
  expect(await spacesReplacementResponse.json()).toEqual({ ok: true });
  const invalidatedMandatorySession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${mandatoryToken}` },
    })
  );
  expect(invalidatedMandatorySession.status).toBe(401);

  const spacesToken = await bearerFor(
    passwordEmail,
    spacesPassword,
    `ticket-04-spaces-${crypto.randomUUID()}`
  );
  const persistedSession = await prisma.session.findFirst({
    orderBy: { createdAt: "desc" },
    where: { userId: passwordUser.id },
  });
  if (!persistedSession) {
    throw new Error("The replacement password session was not persisted");
  }
  const sessionDurationMs = 60 * 60 * 1000;
  expect(
    Math.abs(
      persistedSession.expiresAt.getTime() -
        persistedSession.createdAt.getTime() -
        sessionDurationMs
    )
  ).toBeLessThanOrEqual(1000);

  interface SessionBody {
    session?: { expiresAt?: string };
    user?: { mustChangePassword?: boolean };
  }
  const readSession = async (token: string): Promise<SessionBody> => {
    const response = await app.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${token}` },
      })
    );
    expect(response.status).toBe(200);
    return (await response.json()) as SessionBody;
  };
  const firstSessionRead = await readSession(spacesToken);
  const secondSessionRead = await readSession(spacesToken);
  const firstExpiresAt = firstSessionRead.session?.expiresAt;
  const secondExpiresAt = secondSessionRead.session?.expiresAt;
  expect(firstSessionRead.user?.mustChangePassword).toBe(false);
  expect(firstExpiresAt).toBe(secondExpiresAt);
  expect(firstExpiresAt).toBe(persistedSession.expiresAt.toISOString());

  const maximumPassword = "x".repeat(128);
  const maximumReplacementResponse = await replacePassword(
    spacesToken,
    spacesPassword,
    maximumPassword
  );
  expect(maximumReplacementResponse.status).toBe(200);
  expect(await maximumReplacementResponse.json()).toEqual({ ok: true });
  const invalidatedSpacesSession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${spacesToken}` },
    })
  );
  expect(invalidatedSpacesSession.status).toBe(401);

  const maximumToken = await bearerFor(
    passwordEmail,
    maximumPassword,
    `ticket-04-maximum-${crypto.randomUUID()}`
  );
  const tooLongResponse = await replacePassword(
    maximumToken,
    maximumPassword,
    "x".repeat(129)
  );
  expect(tooLongResponse.status).toBe(400);
  expect(await tooLongResponse.json()).toMatchObject({
    error: "password_too_long",
  });
  const maximumSession = await readSession(maximumToken);
  expect(maximumSession.user?.mustChangePassword).toBe(false);

  const logoutResponse = await app.handle(
    new Request("http://test.local/api/auth/sign-out", {
      headers: { Authorization: `Bearer ${maximumToken}` },
      method: "POST",
    })
  );
  expect(logoutResponse.status).toBe(200);
  const revokedLogoutSession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${maximumToken}` },
    })
  );
  expect(revokedLogoutSession.status).toBe(401);

  const throttleEmail = `ticket-04-throttle-${crypto.randomUUID()}@example.com`;
  const throttlePassword = "Ticket04-throttle-password";
  await createCredentialFixture({
    email: throttleEmail,
    name: "Ticket 04 Throttle User",
    password: throttlePassword,
  });
  const knownWrongResponse = await signIn(
    throttleEmail,
    "wrong-password",
    `ticket-04-known-${crypto.randomUUID()}`
  );
  const knownWrongBody = (await knownWrongResponse.json()) as {
    error?: string;
  };
  const unknownWrongResponse = await signIn(
    `ticket-04-unknown-${crypto.randomUUID()}@example.com`,
    "wrong-password",
    `ticket-04-unknown-${crypto.randomUUID()}`
  );
  const unknownWrongBody = (await unknownWrongResponse.json()) as {
    error?: string;
  };
  expect(knownWrongResponse.status).toBe(401);
  expect(unknownWrongResponse.status).toBe(401);
  expect(knownWrongBody.error).toBe("invalid_credentials");
  expect(unknownWrongBody.error).toBe(knownWrongBody.error);

  const parallelEmail = `ticket-04-parallel-${crypto.randomUUID()}@example.com`;
  const parallelIp = `ticket-04-parallel-ip-${crypto.randomUUID()}`;
  const parallelResponses = await Promise.all(
    Array.from({ length: 8 }, () =>
      signIn(parallelEmail, "wrong-password", parallelIp)
    )
  );
  const parallelStatuses = parallelResponses
    .map((response) => response.status)
    .toSorted((left, right) => left - right);
  expect(parallelStatuses).toEqual([401, 401, 401, 401, 401, 429, 429, 429]);

  const normalizedIp = `ticket-04-normalized-ip-${crypto.randomUUID()}`;
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const emailVariant =
      attempt % 2 === 0 ? ` ${throttleEmail.toUpperCase()} ` : throttleEmail;
    const response = await signIn(emailVariant, "wrong-password", normalizedIp);
    expect(response.status).toBe(401);
    expect(await response.json()).toMatchObject({
      error: "invalid_credentials",
    });
  }
  const normalizedSixthResponse = await signIn(
    ` ${throttleEmail.toUpperCase()} `,
    "wrong-password",
    normalizedIp
  );
  expect(normalizedSixthResponse.status).toBe(429);
  expect(await normalizedSixthResponse.json()).toMatchObject({
    error: "login_throttled",
  });

  const isolatedIp = `ticket-04-isolated-ip-${crypto.randomUUID()}`;
  const isolatedResponse = await signIn(
    throttleEmail,
    "wrong-password",
    isolatedIp
  );
  expect(isolatedResponse.status).toBe(401);
  expect(await isolatedResponse.json()).toMatchObject({
    error: "invalid_credentials",
  });

  const clearingIp = `ticket-04-clearing-ip-${crypto.randomUUID()}`;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const response = await signIn(throttleEmail, "wrong-password", clearingIp);
    expect(response.status).toBe(401);
  }
  const successfulLoginToken = await bearerFor(
    throttleEmail,
    throttlePassword,
    clearingIp
  );
  expect(successfulLoginToken).toBeTruthy();
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const response = await signIn(throttleEmail, "wrong-password", clearingIp);
    expect(response.status).toBe(401);
  }
  const clearedSixthResponse = await signIn(
    throttleEmail,
    "wrong-password",
    clearingIp
  );
  expect(clearedSixthResponse.status).toBe(429);
  expect(await clearedSixthResponse.json()).toMatchObject({
    error: "login_throttled",
  });
  const accountRequest = (
    method: string,
    pathname: string,
    token: string,
    body?: Record<string, unknown>
  ): Promise<Response> => {
    const headers =
      body === undefined
        ? { Authorization: `Bearer ${token}` }
        : { ...jsonHeaders, Authorization: `Bearer ${token}` };
    return app.handle(
      new Request(`http://test.local${pathname}`, {
        body: body === undefined ? undefined : JSON.stringify(body),
        headers,
        method,
      })
    );
  };
  const sessionStatus = async (token: string): Promise<number> => {
    const response = await app.handle(
      new Request("http://test.local/api/session", {
        headers: { Authorization: `Bearer ${token}` },
      })
    );
    return response.status;
  };

  const managedEmail = `ticket-07-managed-${crypto.randomUUID()}@example.com`;
  const managedCreateResponse = await accountRequest(
    "POST",
    "/api/admin/users",
    adminBearer,
    {
      email: ` ${managedEmail.toUpperCase()} `,
      name: "Ticket 07 Managed User",
      role: "user",
    }
  );
  expect(managedCreateResponse.status).toBe(200);
  const managedCreateBody = (await managedCreateResponse.json()) as {
    temporaryPassword?: unknown;
    user?: Record<string, unknown>;
  };
  const managedTemporaryPassword = managedCreateBody.temporaryPassword;
  const managedUser = managedCreateBody.user;
  if (
    typeof managedTemporaryPassword !== "string" ||
    !managedUser ||
    typeof managedUser.id !== "string"
  ) {
    throw new TypeError("The account creation response was malformed");
  }
  const managedId = managedUser.id;
  expect(Object.keys(managedUser).toSorted()).toEqual([
    "createdAt",
    "email",
    "enabled",
    "id",
    "mustChangePassword",
    "name",
    "role",
    "updatedAt",
  ]);
  expect(managedUser).toMatchObject({
    email: managedEmail,
    enabled: true,
    mustChangePassword: true,
    role: "user",
  });
  expect(managedTemporaryPassword).toMatch(/^[A-Za-z0-9_-]+$/u);
  expect(managedTemporaryPassword.length).toBe(32);
  const managedAccount = await prisma.account.findFirst({
    where: { providerId: "credential", userId: managedId },
  });
  expect(managedAccount?.password).not.toBe(managedTemporaryPassword);

  const cursorSeed = Array.from({ length: 21 }, (_, index) => ({
    email: `ticket-07-page-${crypto.randomUUID()}-${index}@example.com`,
    emailVerified: true,
    id: crypto.randomUUID(),
    name: `Ticket 07 Page ${index}`,
    role: "user" as const,
  }));
  await prisma.user.createMany({ data: cursorSeed });
  const firstPageResponse = await accountRequest(
    "GET",
    "/api/admin/users",
    adminBearer
  );
  expect(firstPageResponse.status).toBe(200);
  const firstPageBody = (await firstPageResponse.json()) as {
    nextCursor?: unknown;
    users?: Record<string, unknown>[];
  };
  if (
    !Array.isArray(firstPageBody.users) ||
    typeof firstPageBody.nextCursor !== "string"
  ) {
    throw new TypeError("The first account page was malformed");
  }
  expect(firstPageBody.users).toHaveLength(20);
  const firstPageIds = new Set(
    firstPageBody.users.map((pageUser) => String(pageUser.id))
  );
  const secondPageResponse = await accountRequest(
    "GET",
    `/api/admin/users?cursor=${encodeURIComponent(firstPageBody.nextCursor)}`,
    adminBearer
  );
  expect(secondPageResponse.status).toBe(200);
  const secondPageBody = (await secondPageResponse.json()) as {
    users?: Record<string, unknown>[];
  };
  if (!Array.isArray(secondPageBody.users)) {
    throw new TypeError("The second account page was malformed");
  }
  expect(
    secondPageBody.users.every(
      (pageUser) => !firstPageIds.has(String(pageUser.id))
    )
  ).toBe(true);
  const invalidCursorResponse = await accountRequest(
    "GET",
    "/api/admin/users?cursor=------------------------------------",
    adminBearer
  );
  expect(invalidCursorResponse.status).toBe(400);
  const normalizedListResponse = await accountRequest(
    "GET",
    `/api/admin/users?email=${encodeURIComponent(` ${managedEmail.toUpperCase()} `)}&role=user&enabled=true`,
    adminBearer
  );
  expect(normalizedListResponse.status).toBe(200);
  const normalizedListBody = (await normalizedListResponse.json()) as {
    users?: Record<string, unknown>[];
  };
  expect(normalizedListBody.users).toHaveLength(1);
  expect(normalizedListBody.users?.[0]).toMatchObject({
    email: managedEmail,
    id: managedId,
    role: "user",
  });
  expect(JSON.stringify(normalizedListBody)).not.toContain(
    managedTemporaryPassword
  );

  const managedTemporaryToken = await bearerFor(
    managedEmail,
    managedTemporaryPassword,
    `ticket-07-managed-temporary-${crypto.randomUUID()}`
  );
  const managedTemporarySession = await app.handle(
    new Request("http://test.local/api/session", {
      headers: { Authorization: `Bearer ${managedTemporaryToken}` },
    })
  );
  expect(managedTemporarySession.status).toBe(200);
  expect(await managedTemporarySession.json()).toMatchObject({
    user: { mustChangePassword: true },
  });
  const managedPassword = "Ticket07-managed-permanent-password";
  const managedPasswordChange = await replacePassword(
    managedTemporaryToken,
    managedTemporaryPassword,
    managedPassword
  );
  expect(managedPasswordChange.status).toBe(200);
  expect(await sessionStatus(managedTemporaryToken)).toBe(401);
  let managedToken = await bearerFor(
    managedEmail,
    managedPassword,
    `ticket-07-managed-live-${crypto.randomUUID()}`
  );

  const managedDisableResponse = await accountRequest(
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { enabled: false }
  );
  expect(managedDisableResponse.status).toBe(200);
  expect(await sessionStatus(managedToken)).toBe(401);
  const managedEnableResponse = await accountRequest(
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { enabled: true }
  );
  expect(managedEnableResponse.status).toBe(200);
  managedToken = await bearerFor(
    managedEmail,
    managedPassword,
    `ticket-07-managed-enabled-${crypto.randomUUID()}`
  );
  const changedManagedEmail = `ticket-07-managed-renamed-${crypto.randomUUID()}@example.com`;
  const managedEmailResponse = await accountRequest(
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { email: ` ${changedManagedEmail.toUpperCase()} ` }
  );
  expect(managedEmailResponse.status).toBe(200);
  expect(await sessionStatus(managedToken)).toBe(401);
  managedToken = await bearerFor(
    changedManagedEmail,
    managedPassword,
    `ticket-07-managed-renamed-${crypto.randomUUID()}`
  );
  const managedPromoteResponse = await accountRequest(
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { role: "admin" }
  );
  expect(managedPromoteResponse.status).toBe(200);
  expect(await sessionStatus(managedToken)).toBe(401);
  const promotedManagedToken = await bearerFor(
    changedManagedEmail,
    managedPassword,
    `ticket-07-managed-promoted-${crypto.randomUUID()}`
  );
  const promotedListResponse = await accountRequest(
    "GET",
    "/api/admin/users",
    promotedManagedToken
  );
  expect(promotedListResponse.status).toBe(200);
  const managedDemoteResponse = await accountRequest(
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { role: "user" }
  );
  expect(managedDemoteResponse.status).toBe(200);
  expect(await sessionStatus(promotedManagedToken)).toBe(401);
  managedToken = await bearerFor(
    changedManagedEmail,
    managedPassword,
    `ticket-07-managed-demoted-${crypto.randomUUID()}`
  );

  const managedResetResponse = await accountRequest(
    "POST",
    `/api/admin/users/${managedId}/password-reset`,
    adminBearer
  );
  expect(managedResetResponse.status).toBe(200);
  const managedResetBody = (await managedResetResponse.json()) as {
    temporaryPassword?: unknown;
    user?: Record<string, unknown>;
  };
  if (
    typeof managedResetBody.temporaryPassword !== "string" ||
    !managedResetBody.user
  ) {
    throw new Error("The password reset response was malformed");
  }
  const resetTemporaryPassword = managedResetBody.temporaryPassword;
  expect(resetTemporaryPassword).not.toBe(managedTemporaryPassword);
  expect(managedResetBody.user).toMatchObject({
    id: managedId,
    mustChangePassword: true,
    role: "user",
  });
  expect(await sessionStatus(managedToken)).toBe(401);
  const resetTemporaryToken = await bearerFor(
    changedManagedEmail,
    resetTemporaryPassword,
    `ticket-07-managed-reset-${crypto.randomUUID()}`
  );
  expect(await sessionStatus(resetTemporaryToken)).toBe(200);
  const resetPassword = "Ticket07-managed-reset-permanent-password";
  const resetPasswordChange = await replacePassword(
    resetTemporaryToken,
    resetTemporaryPassword,
    resetPassword
  );
  expect(resetPasswordChange.status).toBe(200);
  expect(await sessionStatus(resetTemporaryToken)).toBe(401);
  managedToken = await bearerFor(
    changedManagedEmail,
    resetPassword,
    `ticket-07-managed-reset-live-${crypto.randomUUID()}`
  );

  const duplicateCreateResponse = await accountRequest(
    "POST",
    "/api/admin/users",
    adminBearer,
    {
      email: ` ${changedManagedEmail.toUpperCase()} `,
      name: "Ticket 07 Duplicate",
      role: "user",
    }
  );
  expect(duplicateCreateResponse.status).toBe(409);
  expect(await duplicateCreateResponse.json()).toMatchObject({
    error: "email_in_use",
  });
  const malformedPatchResponse = await accountRequest(
    "PATCH",
    `/api/admin/users/${managedId}`,
    adminBearer,
    { email: "another@example.com", role: "user" }
  );
  expect(malformedPatchResponse.status).toBe(400);
  expect(await malformedPatchResponse.json()).toMatchObject({
    error: "invalid_request",
  });
  const unknownResetResponse = await accountRequest(
    "POST",
    `/api/admin/users/${crypto.randomUUID()}/password-reset`,
    adminBearer
  );
  expect(unknownResetResponse.status).toBe(404);
  const invalidTargetSecret = `ticket-07-target-secret-${crypto.randomUUID()}`;
  const invalidTargetResponse = await accountRequest(
    "POST",
    `/api/admin/users/${encodeURIComponent(invalidTargetSecret)}/password-reset`,
    adminBearer
  );
  expect(invalidTargetResponse.status).toBe(404);
  const userForbiddenResponse = await accountRequest(
    "POST",
    "/api/admin/users",
    managedToken,
    {
      email: `ticket-07-forbidden-${crypto.randomUUID()}@example.com`,
      name: "Ticket 07 Forbidden",
      role: "user",
    }
  );
  expect(userForbiddenResponse.status).toBe(403);
  expect(await userForbiddenResponse.json()).toMatchObject({
    error: "forbidden",
  });

  const provisionAdmin = async (label: string) => {
    const email = `ticket-07-${label}-${crypto.randomUUID()}@example.com`;
    const provisionResponse = await accountRequest(
      "POST",
      "/api/admin/users",
      adminBearer,
      { email, name: `Ticket 07 ${label}`, role: "admin" }
    );
    expect(provisionResponse.status).toBe(200);
    const body = (await provisionResponse.json()) as {
      temporaryPassword?: unknown;
      user?: Record<string, unknown>;
    };
    if (
      typeof body.temporaryPassword !== "string" ||
      !body.user ||
      typeof body.user.id !== "string"
    ) {
      throw new Error("The Admin provisioning response was malformed");
    }
    const temporaryToken = await bearerFor(
      email,
      body.temporaryPassword,
      `ticket-07-${label}-temporary-${crypto.randomUUID()}`
    );
    const newPassword = `Ticket07-${label}-permanent-password`;
    const passwordChange = await replacePassword(
      temporaryToken,
      body.temporaryPassword,
      newPassword
    );
    expect(passwordChange.status).toBe(200);
    return {
      email,
      id: body.user.id,
      password: newPassword,
      temporaryPassword: body.temporaryPassword,
      token: await bearerFor(
        email,
        newPassword,
        `ticket-07-${label}-live-${crypto.randomUUID()}`
      ),
    };
  };
  const authorityAdminA = await provisionAdmin("authority-a");
  const authorityAdminB = await provisionAdmin("authority-b");
  const authorityListResponses = await Promise.all([
    accountRequest("GET", "/api/admin/users", authorityAdminA.token),
    accountRequest("GET", "/api/admin/users", authorityAdminB.token),
  ]);
  expect(authorityListResponses.map((response) => response.status)).toEqual([
    200, 200,
  ]);
  const enabledAdminsBeforeRace = await prisma.user.findMany({
    select: { id: true },
    where: { enabled: true, role: "admin" },
  });
  for (const existingAdmin of enabledAdminsBeforeRace) {
    if (
      existingAdmin.id === authorityAdminA.id ||
      existingAdmin.id === authorityAdminB.id
    ) {
      continue;
    }
    const demoteResponse = await accountRequest(
      "PATCH",
      `/api/admin/users/${existingAdmin.id}`,
      authorityAdminA.token,
      { role: "user" }
    );
    expect(demoteResponse.status).toBe(200);
  }
  expect(
    await prisma.user.count({ where: { enabled: true, role: "admin" } })
  ).toBe(2);
  const [disableRaceResponse, demoteRaceResponse] = await Promise.all([
    accountRequest(
      "PATCH",
      `/api/admin/users/${authorityAdminA.id}`,
      authorityAdminA.token,
      { enabled: false }
    ),
    accountRequest(
      "PATCH",
      `/api/admin/users/${authorityAdminB.id}`,
      authorityAdminB.token,
      { role: "user" }
    ),
  ]);
  expect(
    [disableRaceResponse.status, demoteRaceResponse.status].toSorted()
  ).toEqual([200, 409]);
  expect(
    [disableRaceResponse, demoteRaceResponse].some(
      (response) => response.status === 409
    )
  ).toBe(true);
  const raceBodies = await Promise.all([
    disableRaceResponse.json(),
    demoteRaceResponse.json(),
  ]);
  expect(
    raceBodies.some(
      (body) =>
        (body as Record<string, unknown>).error === "final_admin_required"
    )
  ).toBe(true);
  expect(
    await prisma.user.count({ where: { enabled: true, role: "admin" } })
  ).toBe(1);
  if (disableRaceResponse.status === 200) {
    expect(await sessionStatus(authorityAdminA.token)).toBe(401);
    expect(await sessionStatus(authorityAdminB.token)).toBe(200);
  } else {
    expect(await sessionStatus(authorityAdminA.token)).toBe(200);
    expect(await sessionStatus(authorityAdminB.token)).toBe(401);
  }

  const erasureAdmin =
    disableRaceResponse.status === 200 ? authorityAdminB : authorityAdminA;
  const erasureAdminBearer = erasureAdmin.token;
  const erasureAdminId = erasureAdmin.id;
  const finalAdminDelete = await accountRequest(
    "DELETE",
    `/api/admin/users/${erasureAdmin.id}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(finalAdminDelete.status).toBe(409);
  expect(await finalAdminDelete.json()).toMatchObject({
    error: "final_admin_required",
  });
  const releaseLeaseIfPresent = async (
    targetId: string,
    releaseUrl: string,
    token: string
  ) => {
    const leaseCount = await prisma.editorLease.count({
      where: { targetId },
    });
    if (leaseCount === 0) {
      return;
    }
    const leaseReleaseResponse = await app.handle(
      new Request(new URL(releaseUrl, "http://test.local").toString(), {
        headers: { Authorization: `Bearer ${token}` },
        method: "DELETE",
      })
    );
    expect(leaseReleaseResponse.status).toBe(200);
  };

  const pictureResponseBeforeDeletion = await prisma.response.findUniqueOrThrow(
    {
      select: {
        draftObjectKey: true,
        externalReferenceDigest: true,
      },
      where: { id: pictureResponseId },
    }
  );
  const pictureSubmissionBeforeDeletion =
    await prisma.submission.findUniqueOrThrow({
      select: { id: true, objectKey: true },
      where: { responseId: pictureResponseId },
    });
  const pictureCorrectionBeforeDeletion =
    await prisma.correction.findFirstOrThrow({
      select: { id: true, objectKey: true },
      where: { responseId: pictureResponseId },
    });
  await releaseLeaseIfPresent(
    pictureResponseId,
    pictureCorrectionEditor.bridge.lease.releaseUrl,
    erasureAdminBearer
  );
  const forbiddenResponseDeletion = await app.handle(
    new Request(`http://test.local/api/admin/responses/${pictureResponseId}`, {
      body: JSON.stringify({ confirm: true }),
      headers: { ...jsonHeaders, Authorization: `Bearer ${userBearer}` },
      method: "DELETE",
    })
  );
  expect(forbiddenResponseDeletion.status).toBe(403);
  const malformedResponseDeletion = await accountRequest(
    "DELETE",
    `/api/admin/responses/${pictureResponseId}`,
    erasureAdminBearer,
    { confirm: false }
  );
  expect(malformedResponseDeletion.status).toBe(400);
  const pictureDeletionResponse = await accountRequest(
    "DELETE",
    `/api/admin/responses/${pictureResponseId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(pictureDeletionResponse.status).toBe(200);
  expect(await pictureDeletionResponse.json()).toEqual({ deleted: true });
  expect(
    await prisma.response.findUnique({ where: { id: pictureResponseId } })
  ).toBeNull();
  expect(
    await prisma.submission.findUnique({
      where: { id: pictureSubmissionBeforeDeletion.id },
    })
  ).toBeNull();
  expect(
    await prisma.correction.findUnique({
      where: { id: pictureCorrectionBeforeDeletion.id },
    })
  ).toBeNull();
  expect(
    await prisma.prefillSnapshot.count({
      where: { responseId: pictureResponseId },
    })
  ).toBe(0);
  expect(
    await prisma.editorLease.count({ where: { targetId: pictureResponseId } })
  ).toBe(0);
  expect(
    await prisma.operation.count({
      where: {
        OR: [
          { responseId: pictureResponseId },
          { submissionId: pictureSubmissionBeforeDeletion.id },
          { correctionId: pictureCorrectionBeforeDeletion.id },
        ],
      },
    })
  ).toBe(0);
  if (pictureResponseBeforeDeletion.draftObjectKey) {
    expect(
      await objectExists(pictureResponseBeforeDeletion.draftObjectKey)
    ).toBe(false);
  }
  expect(await objectExists(pictureSubmissionBeforeDeletion.objectKey)).toBe(
    false
  );
  expect(await objectExists(pictureCorrectionBeforeDeletion.objectKey)).toBe(
    false
  );
  const pictureDeletionAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "delete_response",
      outcome: "success",
      targetId: pictureResponseId,
    },
  });
  expect(pictureDeletionAudit.safeMetadata).toEqual({});
  expect(JSON.stringify(pictureDeletionAudit)).not.toContain("ตรวจสอบรูปภาพ");
  expect(JSON.stringify(pictureDeletionAudit)).not.toContain(
    pictureSubmissionBeforeDeletion.objectKey
  );
  const pictureTombstone = await prisma.deletionTombstone.findUniqueOrThrow({
    where: {
      responseLookupDigest: createHash("sha256")
        .update(pictureResponseId)
        .digest("hex"),
    },
  });
  expect(pictureTombstone).toMatchObject({
    actorId: erasureAdminId,
    externalReferenceDigest:
      pictureResponseBeforeDeletion.externalReferenceDigest,
    outcome: "success",
  });

  const blockedUserDeletion = await accountRequest(
    "DELETE",
    `/api/admin/users/${userId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(blockedUserDeletion.status).toBe(409);
  expect(await blockedUserDeletion.json()).toMatchObject({
    error: "personal_data_remains",
  });
  await releaseLeaseIfPresent(responseId, userLease.releaseUrl, userBearer);
  const mainResponseBeforeDeletion = await prisma.response.findUniqueOrThrow({
    select: {
      draftObjectKey: true,
      externalReferenceDigest: true,
    },
    where: { id: responseId },
  });
  const mainSubmissionBeforeDeletion =
    await prisma.submission.findUniqueOrThrow({
      select: { id: true, objectKey: true },
      where: { responseId },
    });
  const mainDeletionResponse = await accountRequest(
    "DELETE",
    `/api/admin/responses/${responseId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(mainDeletionResponse.status).toBe(200);
  expect(await mainDeletionResponse.json()).toEqual({ deleted: true });
  expect(await sessionStatus(userBearer)).toBe(401);
  expect(
    await prisma.response.findUnique({ where: { id: responseId } })
  ).toBeNull();
  expect(
    await prisma.submission.findUnique({
      where: { id: mainSubmissionBeforeDeletion.id },
    })
  ).toBeNull();
  expect(await prisma.prefillSnapshot.count({ where: { responseId } })).toBe(0);
  expect(
    await prisma.editorLease.count({ where: { targetId: responseId } })
  ).toBe(0);
  expect(
    await prisma.operation.count({
      where: {
        OR: [{ responseId }, { submissionId: mainSubmissionBeforeDeletion.id }],
      },
    })
  ).toBe(0);
  if (mainResponseBeforeDeletion.draftObjectKey) {
    expect(await objectExists(mainResponseBeforeDeletion.draftObjectKey)).toBe(
      false
    );
  }
  expect(await objectExists(mainSubmissionBeforeDeletion.objectKey)).toBe(
    false
  );
  const deletedHandoff = await prisma.handoff.findUniqueOrThrow({
    where: { id: postSubmitHandoffRecord.id },
  });
  expect(deletedHandoff).toMatchObject({
    codeDigest: null,
    configurationHash: null,
    deletionResponseLookupDigest: createHash("sha256")
      .update(responseId)
      .digest("hex"),
    filteredValues: null,
    formId: null,
    normalizedEmail: null,
    responseId: null,
    status: "deleted",
  });
  const deletedExternalStatus = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(deletedExternalStatus.status).toBe(200);
  const deletedExternalBody = await deletedExternalStatus.json();
  expect(deletedExternalBody).toMatchObject({
    deletedAt: expect.any(String),
    status: "deleted",
  });
  const staleCapabilityResponse = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/draft`, {
      body: JSON.stringify({
        data: {},
        documentKey: responseDocumentKey,
        responseId,
      }),
      headers: capabilityHeaders(saveDraftCapability),
      method: "POST",
    })
  );
  expect([401, 404]).toContain(staleCapabilityResponse.status);
  const pendingStatusExternalReference = `ticket-22-pending-status-${crypto.randomUUID()}`;
  const pendingStatusResponseDigest = createHash("sha256")
    .update(`ticket-22-pending-response-${crypto.randomUUID()}`)
    .digest("hex");
  await prisma.deletionTombstone.create({
    data: {
      actorId: erasureAdminId,
      externalReferenceDigest: createHash("sha256")
        .update(pendingStatusExternalReference)
        .digest("hex"),
      id: crypto.randomUUID(),
      outcome: "success",
      responseLookupDigest: pendingStatusResponseDigest,
    },
  });
  const pendingStatusObjectKey = `ticket-22-pending-status/${crypto.randomUUID()}`;
  await prisma.objectCleanupIntent.create({
    data: {
      deletionResponseLookupDigest: pendingStatusResponseDigest,
      objectKey: pendingStatusObjectKey,
    },
  });
  const pendingExternalStatus = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({
      externalReference: pendingStatusExternalReference,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(pendingExternalStatus.status).toBe(404);
  expect(await pendingExternalStatus.json()).toMatchObject({
    error: "handoff_unavailable",
  });
  await prisma.objectCleanupIntent.delete({
    where: { objectKey: pendingStatusObjectKey },
  });
  const completedExternalStatus = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({
      externalReference: pendingStatusExternalReference,
    }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(completedExternalStatus.status).toBe(200);
  expect(await completedExternalStatus.json()).toMatchObject({
    status: "deleted",
  });
  expect(JSON.stringify(deletedExternalBody)).not.toContain(userEmail);
  expect(JSON.stringify(deletedExternalBody)).not.toContain(
    "Ticket 17 Prefill"
  );
  const mainDeletionAudit = await prisma.auditEvent.findFirstOrThrow({
    orderBy: { createdAt: "desc" },
    where: {
      action: "delete_response",
      outcome: "success",
      targetId: responseId,
    },
  });
  expect(mainDeletionAudit.safeMetadata).toEqual({});
  expect(JSON.stringify(mainDeletionAudit)).not.toContain(
    JSON.stringify(savedDraftData)
  );
  expect(JSON.stringify(mainDeletionAudit)).not.toContain(
    postSubmitExternalReference
  );
  const mainTombstone = await prisma.deletionTombstone.findUniqueOrThrow({
    where: {
      responseLookupDigest: createHash("sha256")
        .update(responseId)
        .digest("hex"),
    },
  });
  expect(mainTombstone).toMatchObject({
    actorId: erasureAdminId,
    externalReferenceDigest: mainResponseBeforeDeletion.externalReferenceDigest,
    outcome: "success",
  });
  const repeatedMainDeletion = await accountRequest(
    "DELETE",
    `/api/admin/responses/${responseId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(repeatedMainDeletion.status).toBe(200);
  expect(await repeatedMainDeletion.json()).toEqual({ deleted: true });
  const deletedOwnerAccount = await accountRequest(
    "DELETE",
    `/api/admin/users/${userId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(deletedOwnerAccount.status).toBe(200);
  expect(await deletedOwnerAccount.json()).toEqual({ deleted: true });
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_user",
        outcome: "success",
        targetId: userId,
      },
    })
  ).toMatchObject({ actorId: erasureAdminId, targetId: userId });
  expect(await prisma.user.findUnique({ where: { id: userId } })).toBeNull();
  const deletedExternalStatusAfterAccount = await fetch(`${mock.url}/status`, {
    body: JSON.stringify({ externalReference: postSubmitExternalReference }),
    headers: jsonHeaders,
    method: "POST",
  });
  expect(deletedExternalStatusAfterAccount.status).toBe(200);
  expect(await deletedExternalStatusAfterAccount.json()).toMatchObject({
    deletedAt: expect.any(String),
    status: "deleted",
  });

  const cleanupStart = await app.handle(
    new Request(`http://test.local/api/forms/${formRecord.publicId}/start`, {
      headers: { Authorization: `Bearer ${otherUserBearer}` },
      method: "POST",
    })
  );
  expect(cleanupStart.status).toBe(200);
  const cleanupStartBody = (await cleanupStart.json()) as {
    response?: { id?: string };
  };
  const cleanupResponseId = cleanupStartBody.response?.id;
  if (!cleanupResponseId) {
    throw new Error("The cleanup failure response was not created");
  }
  const cleanupResponse = await prisma.response.findUniqueOrThrow({
    select: { draftObjectKey: true },
    where: { id: cleanupResponseId },
  });
  if (!cleanupResponse.draftObjectKey) {
    throw new Error("The cleanup failure object was not created");
  }
  const databaseFailureResponseId = cleanupResponseId;
  await prisma.$executeRawUnsafe(`
    CREATE OR REPLACE FUNCTION ticket22_abort_response_delete()
    RETURNS trigger
    LANGUAGE plpgsql
    AS $function$
    BEGIN
      RAISE EXCEPTION 'ticket22 database failure';
    END;
    $function$;
  `);
  await prisma.$executeRawUnsafe(`
    CREATE TRIGGER ticket22_abort_response_delete
    BEFORE DELETE ON "responses"
    FOR EACH ROW
    EXECUTE FUNCTION ticket22_abort_response_delete();
  `);
  let databaseFailureDeletion: Response;
  try {
    databaseFailureDeletion = await app.handle(
      new Request(
        `http://test.local/api/admin/responses/${databaseFailureResponseId}`,
        {
          body: JSON.stringify({ confirm: true }),
          headers: {
            ...jsonHeaders,
            Authorization: `Bearer ${erasureAdminBearer}`,
          },
          method: "DELETE",
        }
      )
    );
  } finally {
    await prisma.$executeRawUnsafe(`
      DROP TRIGGER IF EXISTS ticket22_abort_response_delete ON "responses";
    `);
    await prisma.$executeRawUnsafe(`
      DROP FUNCTION IF EXISTS ticket22_abort_response_delete();
    `);
  }
  expect(databaseFailureDeletion.status).toBe(500);
  expect(await databaseFailureDeletion.json()).toMatchObject({
    error: "internal_error",
  });
  expect(
    await prisma.response.findUnique({
      where: { id: databaseFailureResponseId },
    })
  ).not.toBeNull();
  expect(
    await prisma.objectCleanupIntent.count({
      where: {
        deletionResponseLookupDigest: createHash("sha256")
          .update(databaseFailureResponseId)
          .digest("hex"),
      },
    })
  ).toBe(0);
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_response",
        outcome: "failure",
        targetId: databaseFailureResponseId,
      },
    })
  ).toMatchObject({ safeMetadata: { errorCode: "internal_error" } });
  let failCleanup = true;
  const cleanupFailureApp = createApp({
    deleteObject: async (key) => {
      if (failCleanup) {
        throw new Error(`deterministic cleanup failure: ${key}`);
      }
      await deleteObject(key);
    },
  });
  const failedCleanupDeletion = await cleanupFailureApp.handle(
    new Request(`http://test.local/api/admin/responses/${cleanupResponseId}`, {
      body: JSON.stringify({ confirm: true }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${erasureAdminBearer}`,
      },
      method: "DELETE",
    })
  );
  expect(failedCleanupDeletion.status).toBe(503);
  expect(await failedCleanupDeletion.json()).toMatchObject({
    error: "deletion_cleanup_failed",
  });
  expect(
    await prisma.response.findUnique({ where: { id: cleanupResponseId } })
  ).toBeNull();
  expect(
    await prisma.objectCleanupIntent.count({
      where: {
        deletionResponseLookupDigest: createHash("sha256")
          .update(cleanupResponseId)
          .digest("hex"),
      },
    })
  ).toBeGreaterThan(0);
  expect(
    await prisma.auditEvent.findFirst({
      where: {
        action: "delete_response",
        outcome: "success",
        targetId: cleanupResponseId,
      },
    })
  ).toBeNull();
  expect(
    await prisma.auditEvent.findFirst({
      where: {
        action: "delete_response_pending",
        outcome: "success",
        targetId: cleanupResponseId,
      },
    })
  ).toMatchObject({ safeMetadata: {} });
  const blockedCleanupOwner = await accountRequest(
    "DELETE",
    `/api/admin/users/${otherUser.id}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(blockedCleanupOwner.status).toBe(409);
  expect(await blockedCleanupOwner.json()).toMatchObject({
    error: "personal_data_remains",
  });
  failCleanup = false;
  const retriedCleanupDeletion = await cleanupFailureApp.handle(
    new Request(`http://test.local/api/admin/responses/${cleanupResponseId}`, {
      body: JSON.stringify({ confirm: true }),
      headers: {
        ...jsonHeaders,
        Authorization: `Bearer ${erasureAdminBearer}`,
      },
      method: "DELETE",
    })
  );
  expect(retriedCleanupDeletion.status).toBe(200);
  expect(await retriedCleanupDeletion.json()).toEqual({ deleted: true });
  expect(
    await prisma.objectCleanupIntent.count({
      where: {
        deletionResponseLookupDigest: createHash("sha256")
          .update(cleanupResponseId)
          .digest("hex"),
      },
    })
  ).toBe(0);
  expect(
    await prisma.auditEvent.findFirst({
      where: {
        action: "delete_response",
        outcome: "success",
        targetId: cleanupResponseId,
      },
    })
  ).toMatchObject({ safeMetadata: {} });
  expect(await objectExists(cleanupResponse.draftObjectKey)).toBe(false);
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_response",
        outcome: "failure",
        targetId: cleanupResponseId,
      },
    })
  ).toMatchObject({
    safeMetadata: { errorCode: "deletion_cleanup_failed" },
  });

  const managedDeleteResponse = await accountRequest(
    "DELETE",
    `/api/admin/users/${managedId}`,
    erasureAdminBearer,
    { confirm: true }
  );
  expect(managedDeleteResponse.status).toBe(200);
  expect(await managedDeleteResponse.json()).toEqual({ deleted: true });
  expect(await sessionStatus(managedToken)).toBe(401);
  expect(await prisma.user.findUnique({ where: { id: managedId } })).toBeNull();
  expect(
    await prisma.auditEvent.findFirst({
      orderBy: { createdAt: "desc" },
      where: {
        action: "delete_user",
        outcome: "success",
        targetId: managedId,
      },
    })
  ).toMatchObject({ actorId: erasureAdminId, targetId: managedId });

  const accountAuditEvents = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: { targetType: "user" },
  });
  const auditText = JSON.stringify(accountAuditEvents);
  for (const secret of [
    managedTemporaryPassword,
    resetTemporaryPassword,
    authorityAdminA.temporaryPassword,
    authorityAdminB.temporaryPassword,
    managedPassword,
    resetPassword,
    authorityAdminA.password,
    authorityAdminB.password,
    invalidTargetSecret,
  ]) {
    expect(auditText).not.toContain(secret);
  }
  for (const event of accountAuditEvents) {
    const metadata =
      event.safeMetadata &&
      typeof event.safeMetadata === "object" &&
      !Array.isArray(event.safeMetadata)
        ? (event.safeMetadata as Record<string, unknown>)
        : null;
    expect(
      Object.keys(metadata ?? {}).every(
        (key) => key === "change" || key === "errorCode"
      )
    ).toBe(true);
  }
  const managedAuditActions = new Set(
    accountAuditEvents
      .filter((event) => event.targetId === managedId)
      .map((event) => event.action)
  );
  for (const action of [
    "create_user",
    "enable_user",
    "disable_user",
    "change_user_email",
    "promote_user",
    "demote_user",
    "reset_user_password",
    "delete_user",
  ]) {
    expect(managedAuditActions.has(action)).toBe(true);
  }
  expect(
    accountAuditEvents.some(
      (event) =>
        event.action === "create_user" &&
        event.outcome === "failure" &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "email_in_use"
    )
  ).toBe(true);
  expect(
    accountAuditEvents.some(
      (event) =>
        event.outcome === "failure" &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "final_admin_required"
    )
  ).toBe(true);
  expect(
    accountAuditEvents.some(
      (event) =>
        event.action === "update_user" &&
        event.outcome === "failure" &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "invalid_request"
    )
  ).toBe(true);
  const formAuditEvents = await prisma.auditEvent.findMany({
    orderBy: { createdAt: "asc" },
    where: { targetType: "form" },
  });
  const formAuditText = JSON.stringify(formAuditEvents);
  for (const secret of [
    adminBearer,
    competingAdminBearer,
    secretFormTitle,
    secretFormDescription,
    templateDocumentKey,
    createdFormRecord.templateDraft.objectKey,
    publishCapability,
    saveTemplateCapability,
  ]) {
    expect(formAuditText).not.toContain(secret);
  }
  for (const event of formAuditEvents) {
    const metadata =
      event.safeMetadata &&
      typeof event.safeMetadata === "object" &&
      !Array.isArray(event.safeMetadata)
        ? (event.safeMetadata as Record<string, unknown>)
        : null;
    expect(
      Object.keys(metadata ?? {}).every(
        (key) =>
          key === "errorCode" ||
          key === "source" ||
          key === "sourcePublicId" ||
          key === "status"
      )
    ).toBe(true);
  }
  for (const expected of [
    {
      action: "create_form",
      outcome: "success",
      targetId: formPublicId,
    },
    {
      action: "create_form",
      outcome: "failure",
      targetId: null,
    },
    {
      action: "delete_form",
      outcome: "success",
      targetId: uploadPublicId,
    },
    {
      action: "delete_form",
      outcome: "failure",
      targetId: formPublicId,
    },
    {
      action: "save_template_draft",
      outcome: "success",
      targetId: formPublicId,
    },
    {
      action: "save_template_draft",
      outcome: "failure",
      targetId: formPublicId,
    },
    {
      action: "publish_form",
      outcome: "success",
      targetId: formPublicId,
    },
    {
      action: "publish_form",
      outcome: "failure",
      targetId: formPublicId,
    },
  ] as const) {
    expect(
      formAuditEvents.some(
        (event) =>
          event.action === expected.action &&
          event.outcome === expected.outcome &&
          event.targetId === expected.targetId
      )
    ).toBe(true);
  }
  expect(
    formAuditEvents.some(
      (event) =>
        event.action === "delete_form" &&
        event.actorId === competingAdmin.id &&
        event.outcome === "failure" &&
        event.targetId === uploadPublicId &&
        (event.safeMetadata as Record<string, unknown> | null)?.errorCode ===
          "editor_in_use"
    )
  ).toBe(true);
  const auditUnauthenticatedResponse = await app.handle(
    new Request("http://test.local/api/admin/audit-events")
  );
  expect(auditUnauthenticatedResponse.status).toBe(401);
  const auditUserBearer = await bearerFor(
    otherUserEmail,
    password,
    `ticket-23-audit-user-${crypto.randomUUID()}`
  );
  const auditForbiddenResponse = await app.handle(
    new Request("http://test.local/api/admin/audit-events", {
      headers: { Authorization: `Bearer ${auditUserBearer}` },
    })
  );
  expect(auditForbiddenResponse.status).toBe(403);
  expect(await auditForbiddenResponse.json()).toMatchObject({
    error: "forbidden",
  });
  const filteredAuditResponse = await app.handle(
    new Request(
      `http://test.local/api/admin/audit-events?action=delete_response&target=${encodeURIComponent(pictureResponseId)}&outcome=success`,
      { headers: { Authorization: `Bearer ${erasureAdminBearer}` } }
    )
  );
  expect(filteredAuditResponse.status).toBe(200);
  const filteredAuditBody = (await filteredAuditResponse.json()) as {
    events?: {
      action: string;
      outcome: string;
      safeMetadata: Record<string, unknown>;
      targetId: string | null;
    }[];
    nextCursor?: string | null;
  };
  expect(filteredAuditBody.events).toHaveLength(1);
  expect(filteredAuditBody.events?.[0]).toMatchObject({
    action: "delete_response",
    outcome: "success",
    safeMetadata: {},
    targetId: pictureResponseId,
  });
  const auditPageResponse = await app.handle(
    new Request("http://test.local/api/admin/audit-events", {
      headers: { Authorization: `Bearer ${erasureAdminBearer}` },
    })
  );
  expect(auditPageResponse.status).toBe(200);
  const auditPageBody = (await auditPageResponse.json()) as {
    events?: {
      action: string;
      createdAt: string;
      id: string;
      safeMetadata: Record<string, unknown>;
    }[];
    nextCursor?: string | null;
  };
  if (!auditPageBody.events || auditPageBody.events.length === 0) {
    throw new Error("The audit API returned no events");
  }
  for (let index = 1; index < auditPageBody.events.length; index += 1) {
    const previous = auditPageBody.events[index - 1];
    const current = auditPageBody.events[index];
    if (!previous || !current) {
      continue;
    }
    expect(
      new Date(previous.createdAt).getTime() >=
        new Date(current.createdAt).getTime()
    ).toBe(true);
  }
  if (auditPageBody.nextCursor) {
    const nextAuditPageResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/audit-events?cursor=${encodeURIComponent(auditPageBody.nextCursor)}`,
        { headers: { Authorization: `Bearer ${erasureAdminBearer}` } }
      )
    );
    expect(nextAuditPageResponse.status).toBe(200);
    const nextAuditPageBody = (await nextAuditPageResponse.json()) as {
      events?: { id: string }[];
    };
    const auditPageIds = new Set(auditPageBody.events.map((event) => event.id));
    for (const event of nextAuditPageBody.events ?? []) {
      expect(auditPageIds.has(event.id)).toBe(false);
    }
  }
  for (const invalidAuditQuery of [
    "?actor=not-a-uuid",
    "?from=2026-09-20T00:00:00.000Z&to=2026-09-19T00:00:00.000Z",
  ]) {
    const invalidAuditResponse = await app.handle(
      new Request(
        `http://test.local/api/admin/audit-events${invalidAuditQuery}`,
        { headers: { Authorization: `Bearer ${erasureAdminBearer}` } }
      )
    );
    expect(invalidAuditResponse.status).toBe(400);
  }
  const auditApiText = JSON.stringify({
    events: auditPageBody.events,
  });
  for (const secret of [
    adminBearer,
    userBearer,
    saveDraftCapability,
    submitCapability,
    postSubmitExternalReference,
    savedDraftData,
    password,
    managedPassword,
    resetPassword,
  ]) {
    expect(auditApiText).not.toContain(
      typeof secret === "string" ? secret : JSON.stringify(secret)
    );
  }
  let auditMutationBlocked = false;
  const auditProbe = await prisma.auditEvent.findFirstOrThrow();
  try {
    await prisma.auditEvent.update({
      data: { action: "tampered" },
      where: { id: auditProbe.id },
    });
  } catch {
    auditMutationBlocked = true;
  }
  expect(auditMutationBlocked).toBe(true);
  let tombstoneMutationBlocked = false;
  try {
    await prisma.deletionTombstone.update({
      data: { outcome: "failure" },
      where: { responseLookupDigest: mainTombstone.responseLookupDigest },
    });
  } catch {
    tombstoneMutationBlocked = true;
  }
  expect(tombstoneMutationBlocked).toBe(true);
});
