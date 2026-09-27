// oxlint-disable no-await-in-loop avoid-new -- Polling and delay are intentionally sequential.
export const API_ORIGIN =
  import.meta.env.VITE_API_ORIGIN ?? "http://localhost:8080";
export const SESSION_KEY = "onlyoffice.sessionToken";

export class ApiError extends Error {
  readonly status: number;
  readonly code: string;

  constructor(status: number, code: string, message: string) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.code = code;
  }
}

interface ApiErrorBody {
  error?: unknown;
  message?: unknown;
  code?: unknown;
}

export type Role = "admin" | "user";
export interface SessionUser {
  id: string;
  name?: string;
  email: string;
  role?: Role;
  mustChangePassword: boolean;
}
export interface Session {
  user: SessionUser;
  session: { expiresAt: string };
}
export interface SignInResponse {
  error?: string;
  code?: string;
  token?: string;
  session?: { token?: string };
  user?: SessionUser;
}
export interface PasswordReplacementResponse {
  ok: true;
}
export interface AdminUser {
  id: string;
  name: string;
  email: string;
  role: Role;
  enabled: boolean;
  mustChangePassword: boolean;
  createdAt: string;
  updatedAt: string;
}
export interface AdminUserListResponse {
  users: AdminUser[];
  nextCursor: string | null;
}
export interface AdminUserMutationResponse {
  user: AdminUser;
}
export interface AdminUserCredentialResponse extends AdminUserMutationResponse {
  temporaryPassword: string;
}

export type FormStatus = "draft" | "published" | "archived";
export interface FormSummary {
  publicId: string;
  title: string;
  description: string;
  status: FormStatus;
  version: number;
  createdAt: string;
  updatedAt: string;
  hasTemplateDraft: boolean;
  activeDraftCount: number;
  submissionCount: number;
}
export type FormDetail = FormSummary & {
  editorConfigUrl: string;
};
export interface Submission {
  id: string;
  responseId?: string;
  formPublicId?: string;
  formTitle?: string;
  userEmail?: string;
  submissionId?: string;
  latestCorrectionNumber?: number | null;
  status?: string;
  createdAt?: string;
  submittedAt?: string;
  updatedAt?: string;
}
export type ReceiptFieldType =
  | "text"
  | "checkbox"
  | "date"
  | "dropdown"
  | "combo"
  | "picture";
export interface ReceiptField {
  label: string;
  options: { displayText: string; value: string }[];
  placeholder: string | null;
  position: number;
  tag: string;
  type: ReceiptFieldType;
}
export interface ResponseRevision {
  actorEmail: string | null;
  actorName: string | null;
  createdAt: string;
  data: Record<string, unknown>;
  document: {
    available: boolean;
    state: "submission" | "correction";
  };
  id: string | null;
  reason: string | null;
  revision: number;
}
export interface ResponseRevisionsResponse {
  latestRevision: number;
  revisions: ResponseRevision[];
}
export type AdminResultState = "draft" | "submitted";
export interface AdminResult {
  createdAt: string;
  formPublicId: string;
  formTitle: string;
  id: string;
  latestCorrectionNumber: number | null;
  state: AdminResultState;
  submissionId: string | null;
  submittedAt: string | null;
  updatedAt: string;
  userEmail: string;
}
export interface AdminResultListResponse {
  nextCursor: string | null;
  results: AdminResult[];
}
export type AuditOutcome = "failure" | "success";
export type AuditMetadataValue = boolean | number | string | null;
export interface AdminAuditEvent {
  action: string;
  actorId: string | null;
  createdAt: string;
  id: string;
  outcome: AuditOutcome;
  safeMetadata: Record<string, AuditMetadataValue>;
  targetId: string | null;
  targetType: string;
}
export interface AdminAuditListResponse {
  events: AdminAuditEvent[];
  nextCursor: string | null;
}
export interface AdminResultDetail extends AdminResult {
  correction: {
    createdAt: string;
    reason: string;
    revision: number;
  } | null;
  data: Record<string, unknown>;
  document: {
    available: boolean;
    state: "draft" | "submission" | "correction";
  };
  revision: number | null;
}
export interface AdminResultDetailResponse {
  result: AdminResultDetail;
}
export interface Operation {
  id: string;
  status: "pending" | "processing" | "completed" | "failed";
  error?: string;
}
export const formatDate = (value: string | Date | undefined) => {
  if (!value) {
    return "—";
  }
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime())
    ? "—"
    : new Intl.DateTimeFormat("th-TH", {
        dateStyle: "medium",
      }).format(date);
};
export const formatDateTime = (value: string | Date | undefined) => {
  if (!value) {
    return "—";
  }
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime())
    ? "—"
    : new Intl.DateTimeFormat("th-TH", {
        dateStyle: "medium",
        timeStyle: "short",
      }).format(date);
};

export const getToken = () => localStorage.getItem(SESSION_KEY);

export const setToken = (token: string) => {
  localStorage.setItem(SESSION_KEY, token);
};

export const clearToken = () => {
  localStorage.removeItem(SESSION_KEY);
};
const AUTH_ROUTE_PREFIXES = ["/login", "/change-password"] as const;
const SAFE_RETURN_QUERY_KEYS = new Set(["responseid"]);
const RETURN_PATH_MAX_LENGTH = 2048;
const hasSafeReturnQuery = (url: URL): boolean => {
  for (const key of url.searchParams.keys()) {
    if (!SAFE_RETURN_QUERY_KEYS.has(key.toLowerCase())) {
      return false;
    }
  }
  return true;
};

const containsControlCharacter = (value: string): boolean => {
  for (const character of value) {
    const codePoint = character.codePointAt(0) ?? 0;
    if (codePoint <= 31 || codePoint === 127) {
      return true;
    }
  }
  return false;
};

const errorCodeFor = (
  body: ApiErrorBody | null | undefined,
  fallback: string
): string => {
  if (typeof body?.error === "string") {
    return body.error;
  }
  if (typeof body?.code === "string") {
    return body.code;
  }
  return fallback;
};

// oxlint-disable-next-line complexity -- Rejects malformed return paths at every URL decoding boundary.
export const safeReturnPath = (value: unknown): string | null => {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > RETURN_PATH_MAX_LENGTH ||
    !value.startsWith("/") ||
    value.startsWith("//") ||
    value.includes("\\") ||
    containsControlCharacter(value)
  ) {
    return null;
  }

  let decodedValue: string;
  try {
    decodedValue = decodeURIComponent(value);
  } catch {
    return null;
  }
  if (
    decodedValue.startsWith("//") ||
    decodedValue.includes("\\") ||
    containsControlCharacter(decodedValue)
  ) {
    return null;
  }

  const origin =
    typeof window === "undefined" ? "http://localhost" : window.location.origin;
  let parsed: URL;
  try {
    parsed = new URL(value, origin);
  } catch {
    return null;
  }
  if (parsed.origin !== origin) {
    return null;
  }
  if (!hasSafeReturnQuery(parsed)) {
    return null;
  }

  const { pathname, search } = parsed;
  let decodedPathname: string;
  try {
    decodedPathname = decodeURIComponent(pathname);
    if (containsControlCharacter(decodeURIComponent(search))) {
      return null;
    }
  } catch {
    return null;
  }
  if (
    !decodedPathname.startsWith("/") ||
    decodedPathname.startsWith("//") ||
    decodedPathname.includes("\\")
  ) {
    return null;
  }

  const normalizedPathname = decodedPathname.toLowerCase();
  if (
    AUTH_ROUTE_PREFIXES.some(
      (route) =>
        normalizedPathname === route ||
        normalizedPathname.startsWith(`${route}/`)
    )
  ) {
    return null;
  }
  return `${pathname}${search}`;
};

const request = async <T>(path: string, init: RequestInit = {}): Promise<T> => {
  const headers = new Headers(init.headers);
  if (
    !(typeof FormData !== "undefined" && init.body instanceof FormData) &&
    !headers.has("Content-Type")
  ) {
    headers.set("Content-Type", "application/json");
  }
  const token = getToken();
  if (token) {
    headers.set("Authorization", `Bearer ${token}`);
  }
  const response = await fetch(`${API_ORIGIN}${path}`, {
    ...init,
    credentials: "include",
    headers,
  });
  const text = await response.text();
  let body: ApiErrorBody | T | string | null = null;
  if (text) {
    try {
      body = JSON.parse(text) as ApiErrorBody | T;
    } catch {
      body = text;
    }
  }
  if (!response.ok) {
    const fallbackMessage = `Request failed (${response.status})`;
    const errorBody =
      body && typeof body === "object" ? (body as ApiErrorBody) : undefined;
    const code = errorCodeFor(errorBody, "request_failed");
    throw new ApiError(response.status, code, fallbackMessage);
  }
  return body as T;
};

export const apiGet = <T>(path: string) => request<T>(path);

export const apiPost = <T>(
  path: string,
  body?: unknown,
  editorCapability?: string
) =>
  request<T>(path, {
    body: JSON.stringify(body ?? {}),
    headers: editorCapability
      ? { "X-Editor-Capability": editorCapability }
      : undefined,
    method: "POST",
  });

export const apiPostFormData = <T>(path: string, body: FormData) =>
  request<T>(path, {
    body,
    method: "POST",
  });
export const apiPatch = <T>(path: string, body?: unknown) =>
  request<T>(path, {
    body: JSON.stringify(body ?? {}),
    method: "PATCH",
  });
export const apiDelete = <T>(
  path: string,
  body?: unknown,
  options?: Pick<RequestInit, "keepalive">
) =>
  request<T>(path, {
    body: body === undefined ? undefined : JSON.stringify(body),
    method: "DELETE",
    ...options,
  });
export const downloadArtifact = async (path: string, filename: string) => {
  const token = getToken();
  const response = await fetch(`${API_ORIGIN}${path}`, {
    credentials: "include",
    headers: token ? { Authorization: `Bearer ${token}` } : undefined,
  });
  if (!response.ok) {
    throw new ApiError(
      response.status,
      "download_failed",
      `Download failed (${response.status})`
    );
  }
  const blob = await response.blob();
  const href = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = href;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(href);
};

export const getSession = () => request<Session>("/api/session");

export const signIn = async (
  email: string,
  password: string
): Promise<SignInResponse> => {
  const response = await fetch(`${API_ORIGIN}/api/auth/sign-in/email`, {
    body: JSON.stringify({ email, password }),
    credentials: "include",
    headers: { "Content-Type": "application/json" },
    method: "POST",
  });
  let body: SignInResponse | null = null;
  try {
    body = (await response.json()) as SignInResponse;
  } catch {
    body = null;
  }
  if (!response.ok) {
    const code = errorCodeFor(body, "sign_in_failed");
    throw new ApiError(
      response.status,
      code,
      `Request failed (${response.status})`
    );
  }
  const headerToken = response.headers
    .get("set-auth-token")
    ?.replace(/^Bearer\s+/iu, "");
  const token = headerToken ?? body?.token ?? body?.session?.token;
  if (token) {
    setToken(token);
  }
  return body ?? {};
};

export const replacePassword = (currentPassword: string, newPassword: string) =>
  apiPost<PasswordReplacementResponse>("/api/account/password", {
    currentPassword,
    newPassword,
  });

export const signOut = async () => {
  await apiPost("/api/auth/sign-out");
  clearToken();
};

export const waitForOperation = async (
  operationId: string,
  onUpdate?: (operation: Operation) => void
) => {
  for (let attempt = 0; attempt < 300; attempt += 1) {
    const payload = await apiGet<Operation | { operation: Operation }>(
      `/api/operations/${operationId}`
    );
    const operation = "operation" in payload ? payload.operation : payload;
    onUpdate?.(operation);
    if (operation.status === "completed") {
      return operation;
    }
    if (operation.status === "failed") {
      throw new Error(
        operation.error ?? "The operation failed. Your draft is still safe."
      );
    }
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 1000);
    });
  }
  throw new Error(
    "The operation is taking longer than expected. Check back shortly."
  );
};
