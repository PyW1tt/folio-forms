# Folio Forms

Production-oriented MMVP for designing, publishing, and completing DOCX forms with ONLYOFFICE Docs.

The application provides:

- Admin form design and publishing.
- Public, shareable form links.
- Authenticated user form completion.
- User-specific prefill with locked and editable field policies.
- Explicit draft save and resume.
- Immutable submissions with extracted JSON, canonical filled DOCX, and on-demand PDF exports.
- Admin submission review and authorized downloads.

The accepted deployment target is one private single-host Docker Compose stack.

The production Compose topology is defined by `compose.yaml`: Caddy publishes one browser-facing Forms HTTPS host and serves ONLYOFFICE at `https://FORMS_HOST/office`; PostgreSQL, RustFS, the API, and the web container remain private. Every production credential is injected through an uncommitted environment file; no database, RustFS, bootstrap, auth, editor, Handoff, or ONLYOFFICE JWT value is committed.

```bash
docker compose --env-file .env.production -f compose.yaml up -d --build
curl -f https://forms.example.test/ready
```

Startup applies Prisma migrations before serving traffic, bootstraps only the configured first Admin, and reconciles stale Handoffs, Leases, Operations, callback claims, and object cleanup intents before readiness. The verification-only `prefill-mock` service is gated behind `--profile verification`.

This is a single-host, single-disk deployment with no application or off-host backup. Disk loss, ransomware, and regional loss are unrecoverable. The intended ceiling is about 100 accounts, 100 Forms, and 20 concurrent editors; queues, distributed locks, and horizontal scaling are intentionally absent.

## Contents

- [Stack and ports](#stack-and-ports)
- [Quick start](#quick-start)
- [Open the application](#open-the-application)
- [Run modes](#run-modes)
- [Roles and workflow](#roles-and-workflow)
- [DOCX template requirements](#docx-template-requirements)
- [Persistence and artifact storage](#persistence-and-artifact-storage)
- [Authentication and security](#authentication-and-security)
- [HTTP API](#http-api)
- [Frontend routes](#frontend-routes)
- [Repository map](#repository-map)
- [Commands](#commands)
- [Troubleshooting](#troubleshooting)
- [Production hardening](#production-hardening)

## Stack and ports

| Component | Technology | Production exposure |
| --- | --- | --- |
| Forms web/API | React, Vite, Bun, Elysia, `joseluisq/static-web-server:2.44.0-alpine` | `https://FORMS_HOST` (web listens on internal port `80`) |
| Document editor | ONLYOFFICE Docs Community Edition 9.4.0.1 | `https://FORMS_HOST/office` |
| Database | PostgreSQL 18 | Private Compose network only |
| ORM and migrations | Prisma | Applied by the server before readiness |
| Authentication | Better Auth opaque bearer sessions | Forms host only |
| Object storage | RustFS 1.0.0-rc.6, private S3 bucket | Private Compose network only |

The web image keeps the pinned Bun builder and serves its bundle with digest-pinned `joseluisq/static-web-server:2.44.0-alpine` on internal port `80`. `apps/web/sws.toml` configures the `/index.html` SPA fallback and `Content-Security-Policy: frame-ancestors 'self'` plus `X-Frame-Options: SAMEORIGIN`; gateway routes remain unchanged.

## Quick start

This is the private single-host production stack. It requires Docker Compose, one DNS record for the Forms host, and an uncommitted `.env.production` containing deployment values.

### 1. Install dependencies

```bash
bun install
```

### 2. Configure deployment

Set `DATABASE_URL`, `POSTGRES_DB`, `POSTGRES_USER`, `POSTGRES_PASSWORD`, `BETTER_AUTH_SECRET`, `EDITOR_CAPABILITY_SECRET`, `PREFILL_HANDOFF_SECRET`, `ONLYOFFICE_JWT_SECRET`, `RUSTFS_ACCESS_KEY_ID`, `RUSTFS_SECRET_ACCESS_KEY`, `RUSTFS_BUCKET`, `FORMS_HOST`, `CADDY_EMAIL`, and `PREFILL_RETURN_URL` in `.env.production`. Set bootstrap values for the first empty database only. Never commit this file.

### 3. Start the stack

```bash
docker compose --env-file .env.production -f compose.yaml up -d --build
docker compose --env-file .env.production -f compose.yaml ps
curl -f https://forms.example.test/ready
```

Prisma migrations run before the API listens. The configured bootstrap values create the first Admin only; after the first sign-in replaces its password, remove those three values from the deployment environment. No demo seed runs.

### 4. Verify restart recovery

```bash
docker compose --env-file .env.production -f compose.yaml restart server
docker compose --env-file .env.production -f compose.yaml ps
```

The server reconciles stale Handoffs, pending claims, Editor Leases, Operations, callback claims, and object cleanup intents before readiness. PostgreSQL and RustFS data, immutable forms/responses/audit events, and Caddy certificates persist in named volumes.

The deterministic external mock is disabled unless explicitly requested:

```bash
docker compose --env-file .env.production -f compose.yaml --profile verification up prefill-mock
```

This topology has no application or off-host backup. Disk, ransomware, or regional loss is unrecoverable.

## Open the application

Production application:

```text
https://<FORMS_HOST>
```

ONLYOFFICE:

```text
https://<FORMS_HOST>/office
```

For local development only:

```text
http://localhost:5173
```

Create and publish a Form from the Admin interface, then use its generated opaque share link. No seeded Form or fixed share ID exists.

Main screens:

- `/login` — sign in.
- `/dashboard` — user responses and draft resume links.
- `/admin` — admin form list.
- `/admin/users` — search, provision, and administer accounts.
- `/admin/forms/new` — create a form.
- `/admin/forms/:formId` — edit and publish a DOCX template.
- `/admin/forms/:formId/submissions` — inspect submissions for a form.
- `/receipt/:submissionId` — read-only receipt and artifact downloads.

## Run modes

The production run mode is the private Compose topology described above. Do not run a second host API on port `3000` beside it.

Check service state and English startup diagnostics with:

```bash
docker compose --env-file .env.production -f compose.yaml ps
docker compose --env-file .env.production -f compose.yaml logs --tail=100 server
```

The liveness endpoint `/health` is shallow. The readiness endpoint `/ready` stays unavailable until PostgreSQL, RustFS, and the bundled editor preparation are available.

## Roles and workflow

### Account administration

Every Admin has the same account authority. Use `/admin/users` to search by normalized email, filter by role or enabled state, provision an account, correct its email, enable or disable access, change its role, or reset its password. Creation and reset show a generated temporary password only in that response; copy it before leaving the result. The final enabled Admin cannot be disabled or demoted.

Each privileged account attempt appends an immutable Audit Event with its actor, target, action, outcome, and safe metadata. Passwords, hashes, tokens, and other credential material are excluded.

### Admin workflow

1. Sign in with an Admin account.
2. Select **New form**, enter a title and description, then choose the bundled starter DOCX or upload a `.docx` no larger than 25 MiB.
3. The API validates the file type, size, and required DOCX package parts before creating the Form.
4. Open the leased DOCX editor. A second Admin sees a blocked, non-editable state until the active lease is released or expires.
5. Add tagged content controls to the template. While the Template Draft is open, the ONLYOFFICE right-side Form Bridge panel is available only to this authorized Admin editor. Select a content control to inspect its exact tag, mark it required, and choose `editable` or `lock-when-available` Prefill policy. The panel searches the one deterministic external mock schema with bounded cursor paging and shows only nested scalar RFC 6901 JSON Pointers; it never shows external record values, User records, or a mapping canvas. Copy a pointer or apply that exact pointer as the selected control tag before saving the policy.
6. Select **Save Template** and wait for its Operation. A completed save is the exact Draft reopened later; a failed save leaves the prior Draft available for retry.
7. Select **Publish** when the template is ready.
8. Copy the generated share link.
9. Review submitted responses from the form's **View submissions** page.

The Admin Form list shows lifecycle state plus active Draft and Submission counts without exposing database IDs or RustFS object keys. Only a never-published Draft with no Response data can be hard-deleted; deletion removes its Template Draft objects. Create, save, and delete outcomes append attributable, secret-free Form Audit Events. Published Forms keep the same opaque share ID, Published Template, Field Manifest, and Prefill Configuration when an Admin changes only the title or description. Structural or policy changes use **Duplicate Form** instead: the source DOCX and rules are copied into a new editable Draft with a new share ID, while Responses, Submissions, Operations, Leases, and source Audit Events remain with the source. The duplicate can be edited or deleted independently; a Published Form cannot be hard-deleted or returned to Draft.

Publishing:

- Validates the DOCX package and requires at least one tagged supported content control.
- Records each supported control's tag, type, dropdown/combo options, required state, and Prefill policy in an immutable Field Manifest. Picture fields publish as native embedded-image fields with JPEG/PNG, one-image, 10 MiB, and 4096×4096 limits.
- Rejects blank or duplicate tags, malformed controls/options, unknown or unsupported control types, and invalid Prefill policies.
- Creates one immutable Published Template and opaque generated share ID. Save and Publish cannot replace that structural or policy contract in place.
- Failed publication leaves the prior Template Draft editable and exposes no shareable Published Form.
- Structural or policy changes require a new Form and share ID; existing Responses remain reproducible.

### User workflow

1. Open the share link.
2. Sign in.
3. Wait for the user-specific prefill to appear.
4. Open the ONLYOFFICE **Form** tab.
5. Select **Save Draft** to persist a resumable response.
6. Return through the dashboard to resume the same response.
7. Select **Submit** to create the immutable receipt.

The prefill is snapshotted when the response starts. Resuming a draft does not refresh the profile. Starting again after publication invalidates an old draft and creates a new snapshot for the new published version. The User dashboard lists the current User's Draft with the Form title and last-saved time. Save Draft is explicit: incomplete scalar and Picture values are allowed, but unknown tags, wrong types/options/dates, text over 10,000 characters, and response JSON over 256 KiB are rejected. Native Picture controls are never remote-prefilled or serialized into response JSON; submission requires a required Picture to contain one embedded JPEG/PNG within the published byte and dimension limits. Resume reopens the same Response and document under the owning User's Editor Lease; another User cannot access its data, Operation, editor configuration, or document.

Static document images remain non-editable. Picture input uses ONLYOFFICE's native control; the canonical DOCX is authoritative, with no separate image upload or image object. Submission is complete only after the extracted field JSON and canonical filled DOCX are persisted. PDF is an on-demand export and is not durable submission state.

### External editable Prefill handoff

An external system uses the deterministic mock's schema pointers, then calls `POST /api/integrations/prefill/handoffs` with `X-Prefill-Handoff-Secret` (the deployment-only `PREFILL_HANDOFF_SECRET`, distinct from the editor capability, auth, and ONLYOFFICE JWT secrets) and the published Form public ID, normalized email, external reference, and candidate scalar object. Folio filters the candidate through the immutable Prefill Configuration and retains only configured tag values. Invalid credentials and unavailable Forms return the same non-enumerating failure. The runnable deterministic mock lives in `apps/prefill-mock`: set `FOLIO_ORIGIN` and the distinct `PREFILL_HANDOFF_SECRET`, then run `bun run --cwd apps/prefill-mock start`. Its `/schema` endpoint returns the fixed scalar-pointer catalog, `/handoffs` is the narrow server-to-server connector, and `/launch` renders a form whose browser submits the code directly to Folio so the pending cookie remains scoped to the Folio host.

The response contains a 32-byte random, base64url one-time code and the `/prefill/handoff` launch path. The external system must submit that code in the body of a top-level `POST` form; codes in URLs are not accepted. Folio stores only SHA-256 digests, expires the launch code after 120 seconds, and exchanges it for a ten-minute `__Host-folio-pending-claim` cookie (`Secure`, `HttpOnly`, `SameSite=Lax`, `Path=/`). The pending claim survives login and mandatory password replacement; no code or Prefill value is written to a URL, localStorage, or sessionStorage.

After authentication, the clean Form path redeems the claim once. Folio verifies the email, Form, published Prefill Configuration, and expiry inside one serializable transaction before creating the user's Response, copying the published DOCX, and snapshotting editable or `lock-when-available` Prefill values. A configured Form rejects an ordinary share-link start. Discarding its Draft invalidates the claim and requires a new external Handoff. Invalid, expired, replayed, or mismatched claims return an accessible Thai retry state, and Handoff Audit Events contain only safe target references. The same mock polls `POST /status` with only the External Reference; Folio returns only `pending`, `draft`, `submitted`, `deleted`, or `expired` plus lifecycle timestamps and latest Correction number. Status failures and unknown references are non-enumerating. Pending or reserved Handoffs are swept every 60 seconds after expiry: claim, code digest, Prefill, and normalized identity are purged while the lookup digest and timestamps remain.

Receipts expose one keyboard-accessible Thai **กลับไปยังระบบต้นทาง** action using the deployment-configured `PREFILL_RETURN_URL`; the Handoff payload cannot override it. No webhook, email, SMS, or inbound callback is used.

## DOCX template requirements

Fields are ONLYOFFICE content controls. Their tags are the stable field keys used in JSON and prefill data.

Each Form defines its own unique, non-empty tags. Supported controls include text, checkbox, date, dropdown, combo box, and picture fields. The Form Bridge panel uses the exact content-control tag as the Field identity. JSON Pointer tags such as `/person/name` are stored literally (RFC 6901 escaping applies to `/` and `~` inside a segment); the panel's external schema search returns pointer keys and scalar types only. Supported authoring controls are text, checkbox, date, dropdown, combo box, and Picture; publication remains the authority that validates their final types and options.

Picture fields accept exactly one native ONLYOFFICE embedded JPEG or PNG up to 10 MiB and 4096×4096 pixels. Required Picture fields must contain an image. Picture bytes stay in the canonical Draft/Submission DOCX and are preserved by authorized DOCX/PDF export; they never appear in scalar Response JSON or as independent image objects.

The plugin extracts:

- Inline and block text.
- Checkbox values as booleans.
- Date picker values as `YYYY-MM-DD`.
- Dropdown and combo-box stored values.
- Picture controls are deliberately omitted from scalar extraction; their embedded DOCX bytes are authoritative.

The plugin applies prefill through ONLYOFFICE's command API, then restricts respondent sessions to form editing. The action flow temporarily switches the editor to view restriction while extracting and force-saving a draft or submission, then restores form editing.

## Persistence and artifact storage

### PostgreSQL tables

`packages/db/prisma/schema.prisma` defines:

- Better Auth `User`, `Session`, `Account`, and `Verification` models.
- Forms, Template Drafts, immutable Published Templates, Field Manifests, and Prefill Configuration.
- One Response per User/Form, immutable Submissions, Prefill snapshots, and append-only Corrections.
- Handoffs, pending claims, Operations, callback claims, Editor Leases, and durable Object Cleanup Intents.
- immutable Audit Events, login-failure windows, and Deletion Tombstones.

Database constraints enforce:

- Normalized unique account email and opaque Form public IDs.
- One Response per User/Form and one Submission per Response.
- Immutable published contracts and monotonic Correction revisions.
- One active Operation and one Editor Lease per target.
- Accepted lifecycle, ownership, document-reference, and expiry invariants.

### Private object storage

RustFS stores canonical DOCX objects under opaque keys such as:

```text
forms/<formId>/template-draft/<uuid>/docx
forms/<formId>/published/<version>/<uuid>/docx
responses/<responseId>/draft/<uuid>/docx
submissions/<submissionId>/filled/<uuid>/docx
operations/<operationId>/<kind>/<uuid>/docx
```

Form creation and hard deletion record Object Cleanup Intents before an object can become unreachable. A committed create clears its intent atomically with the canonical database reference; a new create has a 15-minute in-flight grace so another server cannot clean it prematurely. Failed RustFS deletion leaves the intent durable, and startup plus minute reconciliation retries eligible cleanup.

- Better Auth handles email/password sign-in and PostgreSQL-backed sessions.
- Sessions use opaque Better Auth tokens, not JWTs.
- The frontend sends `Authorization: Bearer <session-token>`.
- The local web app stores the token at `localStorage["onlyoffice.sessionToken"]`.
- Opening an editor configuration atomically claims or renews a 90-second lease for that Template Draft or Response under the live browser Session. A competing Session receives no editable configuration until the holder releases the lease or it expires.
- The browser Session token and editor capabilities are never included in the signed ONLYOFFICE or plugin configuration. Before each action, the browser parent obtains a fresh five-minute capability bound to the actor, role, Form, document target, action, and active lease, then sends only that capability over the source/origin-pinned bridge.
- ONLYOFFICE Document Server uses `ONLYOFFICE_JWT_SECRET` for signed editor configuration, command/conversion requests, private document downloads, and callbacks. It never accepts a Better Auth Session or editor capability at that boundary.
- Each session expires one hour after issuance; sessions do not refresh or slide. The client clears the token at expiry and warns during the final five minutes.
- Logout and password replacement revoke sessions immediately. Password replacement revokes every session, so the Admin must sign in again.
- Admin account creation, disablement, email/role changes, and password resets are attributable Audit Events. Account disablement and credential or identity changes revoke all target Sessions immediately.
- A serialized database invariant prevents concurrent requests from disabling or demoting the final enabled Admin.
- The server derives identity and role from Better Auth, never from client-supplied role fields.
- Admin routes require the `admin` role.
- User response and submission routes enforce ownership.
- ONLYOFFICE document URLs use a five-minute HMAC token bound to the document key.
- ONLYOFFICE callback userdata uses a server-signed HMAC envelope containing the Operation ID. Its exact token digest and single-use consumption are persisted, so replay protection survives process restarts.
- Callback document downloads are restricted to configured ONLYOFFICE origins, disallow redirects, and enforce a response-size limit.
- Response JSON is restricted to published content-control tags, scalar values, and a bounded payload size.
- Operations expire with stable error codes and atomically roll back submitting Responses when they remain active too long. Startup reconciliation expires stale Operations, callback claims, editor leases, and leases backed by expired Sessions without replacing committed object references.

These protections cover the current local stack, but the deployment defaults below are not suitable for the public internet.

## HTTP API

Protected browser endpoints use the opaque bearer Session header:

```http
Authorization: Bearer <opaque-better-auth-session-token>
```

Mutating editor requests require the active lease-bound action capability returned through the validated editor bridge. Plugin Operation polling uses a separate Operation-bound capability:

```http
X-Editor-Capability: <signed-action-or-operation-capability>
```

### Authentication and health

| Method | Path | Purpose |
| --- | --- | --- |
| `POST` | `/api/auth/sign-in/email` | Provisioned email/password sign-in |
| `POST` | `/api/auth/sign-out` | Revoke the current live Session |
| `GET` | `/api/session` | Read the current live Session and absolute expiry |
| `POST` | `/api/account/password` | Replace the authenticated account password and revoke its Sessions |
| `POST` | `/api/editor-leases/:id/renew` | Renew the current Session's active editor lease |
| `DELETE` | `/api/editor-leases/:id` | Release the current Session's editor lease |
| `GET` | `/health` | API health check |
| `GET` | `/ready` | Readiness check for PostgreSQL, RustFS, and editor preparation |

### Authenticated user form operations

| Method | Path | Purpose |
| --- | --- | --- |
| `GET` | `/api/forms/:publicId` | Read Form metadata after authentication |
| `GET` | `/api/forms/:publicId/editor-config` | Get protected user editor config |
| `POST` | `/api/forms/:publicId/start` | Start or resume one user response |
| `POST` | `/api/forms/:publicId/draft` | Save a draft through an asynchronous operation |
| `POST` | `/api/forms/:publicId/submit` | Submit a response through an asynchronous operation |
| `GET` | `/api/responses/me` | List the current user's responses |
| `GET` | `/api/operations/:id` | Poll an owned or Admin operation |

### Admin account operations

| Method | Path | Purpose |
| --- | --- | --- |
| `GET` | `/api/admin/users` | Cursor-page and filter accounts by normalized email, role, and enabled state |
| `POST` | `/api/admin/users` | Provision an account and disclose its temporary password once |
| `PATCH` | `/api/admin/users/:id` | Enable/disable or change one email/role value and revoke affected Sessions |
| `POST` | `/api/admin/users/:id/password-reset` | Rotate the credential, revoke Sessions, and disclose a temporary password once |

### Admin form operations

| Method | Path | Purpose |
| --- | --- | --- |
| `GET` | `/api/admin/forms` | List lifecycle state plus active Draft and Submission counts |
| `DELETE` | `/api/admin/forms/:publicId` | Remove a never-published Draft without Response data |
| `POST` | `/api/admin/forms` | Create from the starter or an uploaded validated DOCX |
| `GET` | `/api/admin/forms/:publicId` | Read safe form detail and counts |
| `GET` | `/api/admin/forms/:publicId/editor-config` | Claim the exclusive lease and get template editor config |
| `GET` | `/api/admin/forms/:publicId/schema?q=&cursor=` | Search the deterministic scalar-only external schema with cursor paging |
| `GET` | `/api/admin/forms/:publicId/field-rules` | Read the current Draft Field required and Prefill policies through the editor capability |
| `PATCH` | `/api/admin/forms/:publicId/field-rules` | Save one exact Field tag, required state, and Prefill pointer/policy |
| `POST` | `/api/admin/forms/:publicId/save` | Save the Template Draft through an asynchronous Operation |
| `POST` | `/api/admin/forms/:publicId/publish` | Publish a validated template |
| `GET` | `/api/admin/forms/:publicId/submissions` | List all submissions for a form |

### Submission artifacts

| Method | Path | Purpose |
| --- | --- | --- |
| `GET` | `/api/submissions/:id/data` | Read authorized submission JSON and metadata |
| `GET` | `/api/submissions/:id/docx` | Download an authorized filled DOCX |
| `GET` | `/api/submissions/:id/pdf` | Download an authorized PDF |

### ONLYOFFICE integration

| Method | Path | Purpose |
| --- | --- | --- |
| `GET` | `/onlyoffice/document/:key` | Serve a signed document to ONLYOFFICE |
| `GET` | `/onlyoffice-plugin/config.json` | Serve plugin registration metadata |
| `GET` | `/onlyoffice-plugin/index.html` | Serve the plugin iframe |
| `GET` | `/onlyoffice-plugin/plugin.js` | Serve the form bridge |
| `POST` | `/onlyoffice/callback` | Receive signed operation callbacks |

## Frontend routes

| Route | Access | Purpose |
| --- | --- | --- |
| `/` | Authenticated | Redirect Admin to `/admin`, User to `/dashboard` |
| `/login` | Public | Sign in and preserve a safe return path |
| `/dashboard` | User | View and resume responses |
| `/admin` | Admin | View lifecycle state and Draft/Submission counts; remove eligible Draft forms |
| `/admin/audit` | Admin | Read-only cursor-filtered immutable audit events |
| `/admin/forms/new` | Admin | Create from the starter or upload a validated DOCX |
| `/admin/forms/:formId` | Admin | Edit under an exclusive lease, save, publish, and share a template |
| `/admin/forms/:formId/submissions` | Admin | Review form submissions |
| `/admin/forms/:formId/submissions/:submissionId` | Admin | Read a submission and download artifacts |
| `/forms/:publicId/fill` | Authenticated | Fill, save, or submit a shared form |
| `/receipt/:submissionId` | Authorized | Read a completed submission receipt |

ONLYOFFICE is a desktop-oriented editor. The dashboard and administrative shell are responsive; the fill/editor screen displays a desktop recommendation.

## Repository map

```text
apps/
  onlyoffice-plugin/       ONLYOFFICE plugin manifest, UI, extraction, prefill, actions
  server/
    src/app.ts             Elysia routes and asynchronous operation orchestration
    src/index.ts           Production listen entrypoint
    src/onlyoffice.ts      ONLYOFFICE URLs, HMAC tokens, force-save, PDF conversion
    src/storage.ts         Private RustFS object primitives
    Dockerfile             Production API image
  web/
    src/routes/            TanStack Router screens
    src/components/        App shell, editor wrapper, UI primitives
    src/lib/               API client and auth provider
    Dockerfile             Production static web image
    sws.toml               Port 80, SPA fallback, same-origin frame policy
packages/
  auth/                    Better Auth configuration and bearer plugin
  db/                      Prisma schema, generated client, and checked-in migration
  env/                     Server environment validation
  config/                  Shared TypeScript/project configuration
onlyoffice-templates/
  template.docx            Tracked tagged template
compose.yaml               Canonical private single-host Compose topology
Caddyfile                  HTTPS host routing
CONTEXT.md                 Canonical domain glossary
```

## Commands

Install dependencies:

```bash
bun install
```

Database:

```bash
bun run --cwd apps/server db:migrate
bun run --cwd packages/db db:generate
```

Development:

```bash
bun run dev
bun run --cwd apps/server dev
bun run --cwd apps/web dev
```

Quality:

```bash
bun x ultracite fix
bun x ultracite check
bun run check-types
bun run build
```

Docker:

```bash
docker compose --env-file .env.production -f compose.yaml config --quiet
docker compose --env-file .env.production -f compose.yaml up -d --build
docker compose --env-file .env.production -f compose.yaml ps
docker compose --env-file .env.production -f compose.yaml logs --tail=100 server
docker compose --env-file .env.production -f compose.yaml down
```

Run `bun run --cwd apps/server test:http` and `bun run --cwd apps/server test:storage` against an isolated PostgreSQL database and disposable private RustFS bucket.

## Troubleshooting

### API does not start

Check ordering, migration output, and readiness:

```bash
docker compose --env-file .env.production -f compose.yaml ps
docker compose --env-file .env.production -f compose.yaml logs --tail=100 server
curl -f https://forms.example.test/health
curl -f https://forms.example.test/ready
```

`/health` is liveness only. `/ready` returns `503` until the database, RustFS health endpoint, and bundled template are available.

### PostgreSQL or RustFS does not start

Inspect the private service logs:

```bash
docker compose --env-file .env.production -f compose.yaml logs --tail=100 postgres rustfs rustfs-init
```

Do not publish database or RustFS ports. Confirm the injected `DATABASE_URL`, `RUSTFS_ENDPOINT`, bucket, and credentials match the Compose services.

### The editor says that a document is unavailable

Confirm that PostgreSQL, RustFS, ONLYOFFICE, and the server are healthy; `TEMPLATE_PATH` exists in the server image; `ONLYOFFICE_DOCUMENT_BASE_URL` is reachable from the ONLYOFFICE container; and the Forms host serves `/office` through Caddy.

### Form actions are missing

Open the document's **Form** tab inside ONLYOFFICE. The custom actions are served through the Forms host by the Form Bridge plugin. The plugin receives only action-scoped editor capabilities and never receives the browser Session bearer token.

## Operational ceiling

The stack intentionally has one API, one web container, one ONLYOFFICE instance, one PostgreSQL instance, and one RustFS volume set. It adds no queue, distributed lock, horizontal scaling, or high-availability machinery. Plan capacity around 100 accounts, 100 Forms, and 20 concurrent editors.

The stack has no application or off-host backup. Attached-disk loss, ransomware, and regional loss are unrecoverable. Add an explicit backup/restore decision before operating beyond this MMVP.

## Related documentation

- [Domain glossary](CONTEXT.md)
- [Frontend setup research](docs/research/frontend-setup.md)
- [Authentication setup research](docs/research/auth-setup.md)
- [Database setup notes](packages/db/README.md)
