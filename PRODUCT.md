# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

- **Admin** creates, edits, publishes, and shares Forms; Admin can inspect every Draft and Submission.
- **User** fills Forms and can inspect only their own Drafts and Submissions.
- This work gives Admin and User equal priority.

## Product Purpose

Folio Forms turns DOCX templates with tagged content controls into shareable Forms. Admins publish templates; authenticated Users submit responses. A completed Submission retains extracted field data, a filled DOCX, and a PDF.

## Positioning

DOCX remains the canonical Form template and output. A Google-Forms-like page adds another way to fill the same Fields; it does not replace the DOCX workflow.

## Operating Context

- Admin prepares a Template Draft, then publishes it for new Responses.
- A User's Response can hold a Draft and produce one Submission.
- An Admin uploads a PDF into the AI chat or describes a Form in text, iterates on the result in the same chat, and previews a generated DOCX read-only. The source PDF remains available to the agent throughout the active session, including later edit requests. The final DOCX contains static text and tagged content controls; the Admin downloads it, uploads it through the existing Form flow, and finishes editing it in ONLYOFFICE.

## Capabilities and Constraints

- The server calls OmniRoute as the sole model gateway; provider and model selection stay there. Keep the endpoint and API key server-side.
- Pi manages AI sessions and tool orchestration. AI authoring is Admin-only.
- Send the original PDF to the model through OmniRoute; Folio performs no OCR. The configured route must support PDF input and tool calls.
- Personal Codex credentials through OmniRoute are for the account owner's local, single-user tests only. Keep shared AI unavailable until a permitted service credential is configured.
- Before the first external send, tell Admin that PDF and chat data go to OmniRoute and its configured provider. Upstream retention follows their policies; Folio deletes only its own copies.
- Allow model-generated Python only in an isolated, ephemeral document worker. Restrict files to the active session; deny shell and network access; enforce resource and time limits. Model instructions are not a security boundary.
- Admin selects one filling method per Form. Method changes affect existing and new Responses; all Field types and Prefill rules remain available. DOCX stays canonical.
- The native web form shows all Fields from the published DOCX in document order on one scrollable page. It displays the Field Title/Alias and Placeholder, not the DOCX's static text, headings, or layout.
- Use content-control Title/Alias as the display label, Tag as stable Field identity and fallback label, and Placeholder as entry guidance.
- Admin can inspect Drafts and Submissions. Result screens use readable Field data instead of JSON; keep JSON internal.
- Do not allow external websites to frame the web app.
- Keep template-only Field settings hidden during respondent fill. Keep the settings panel scrollable for Admins.
- Provide an in-page, body-sized ONLYOFFICE mode on every editor surface; do not invoke browser fullscreen.
- Render User receipts as a stacked label/value list, never JSON. Admin result detail has read-only DOCX and Field-table tabs, one revision selector for both, and a separate Correction action. Support Drafts and Submissions; retain DOCX/PDF downloads but remove JSON UI actions.
- Default User Receipt and Admin results to the latest revision; let both select every stored revision, including the original and intermediate Corrections. Field values, read-only DOCX view, and downloads must match the selected revision.
- Let Admins copy temporary passwords after user creation and password reset.
- Serve the SPA with `joseluisq/static-web-server` on internal port 80 and preserve the SPA fallback.
- Breaking changes are acceptable during development.
- AI preview shows the generated DOCX read-only.
- Keep PDF, chat, and generated files only for the active AI authoring session. Refreshing the page preserves it; explicit end, logout, or two hours of inactivity ends it and deletes Folio's copies. Do not promise upstream deletion.

## Legacy Login

- Add an optional one-time-code SSO bridge from the old application. Its frontend uses its existing login and token refresh; its backend exchanges a short-lived, single-use code with Folio's backend. Folio never accepts an old access/refresh token as its own session credential.
- After an authenticated old-system User clicks a Form link, Folio creates its own session for the linked local User and opens the requested Form without another login. The Folio login page also offers a legacy SSO button beside local password login; preserve a safe requested Form return path.
- A first-time verified Legacy Identity with a verified email creates a local role=user account without a local password. Its stable provider/subject determines future sign-ins; neither email nor name is a linking key. If verified identity/email is unavailable, do not provision.
- If the legacy email matches an existing local account without a link, do not merge automatically. Store a verified, non-login Pending Account Link for an Admin to review and explicitly approve; the User retries SSO after approval. Never approve a link to an Admin account.
- SSO is for Users only. Admins continue using local login. SSO-only Users have no local password, and self-registration stays disabled.
- Do not synchronize name or email from the old system after a link is established. A linked local account that still requires a temporary password change may use SSO; local password login still requires that change.
- If Folio already has a different signed-in User, stop and require confirmation before switching accounts. When its one-hour session expires during a Form, save the Draft and repeat legacy SSO, then return to the Form; if the old login expired, the old system handles reauthentication.
- The new SSO entry supports ordinary Form links, not the existing email-bound prefill handoff. Existing prefill links retain their current login behavior. Seamless login does not imply synchronized logout or immediate revocation when the old account is disabled.
- Implement Folio's side first. The old application adds its authorize page and code issue/exchange APIs later; cross-system SSO cannot be verified end-to-end until those endpoints exist.
