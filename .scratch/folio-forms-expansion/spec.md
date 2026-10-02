# Extend Folio Forms with native filling, readable results, AI authoring, and legacy SSO

## Problem Statement

Folio Forms currently makes Users fill a DOCX in ONLYOFFICE and exposes raw JSON on result screens. People who want a familiar web questionnaire cannot use one; Drafts, Submissions, and Corrections are hard to read. Admins lack a clear read-only document/table result view, an in-page expanded editor, a scrollable Field-settings panel, a copy action for temporary passwords, and an AI-assisted way to turn a PDF or description into a DOCX Template Draft.

Users already authenticated in the old application also have to authenticate separately in Folio to follow an ordinary Form link. That creates friction and risks assigning a Response to the wrong local User if identities are linked by email alone. The old application will add its part of the login bridge later; Folio must implement its side without sharing the old application's tokens, passwords, or user database.

## Solution

Preserve DOCX as the canonical Form document and Submission artifact. Let an Admin choose ONLYOFFICE or a single-page native web form per Form; both use the same Published Template, Fields, Prefill rules, Draft, and Submission lifecycle. Make result screens human-readable: a stacked Field list on the User Receipt and a read-only DOCX/table pair on Admin result detail, with a shared revision selector that can open every Correction. Retain DOCX/PDF downloads but remove JSON displays and JSON download actions from the UI.

Add Admin-only AI authoring: attach a PDF in Folio's chat or describe a Form; converse and revise; inspect a read-only DOCX preview; download a DOCX with static text and tagged content controls, then upload and finish it in the existing Template Draft editor. Pi manages the agent; Folio sends model requests through OmniRoute. Permit model-generated Python only inside an isolated, bounded document worker. Preserve the source PDF for reinspection during the temporary authoring session, disclose external data processing, and delete Folio's session copies when the session ends.

Add an optional, browser-bound, one-time-code SSO Bridge for ordinary Users arriving from the old application. Verify identity by a backend-to-backend code exchange, bind the stable Legacy Identity to one local User, then create a normal Folio Session and open the intended Form. Retain local password login and Admin-only local account creation. Replace the web image's Nginx server with a pinned static web server without breaking the existing deployment routes or SPA fallback.

## User Stories

1. As a User, I want to follow an ordinary Form link from the old application while logged in there, so that I can open the correct Form without another login.
2. As a User, I want the old application to refresh my old login when possible during this handoff, so that an expired access token does not require an unnecessary new login.
3. As a User, I want to log in at the old application if its login has fully expired, so that my Form link still works after I authenticate there.
4. As a User, I want Folio's login page to offer both local login and old-system login, so that I can choose the account route appropriate to me.
5. As a User, I want an old-system login initiated from a direct Form link to return to that Form, so that I do not have to find it again.
6. As a User, I want a first successful Legacy Identity login to create my Folio User automatically when its verified email is unused, so that I can begin filling immediately.
7. As a User, I want my Legacy Identity to return to the same Folio User on later visits, so that my Drafts and Submissions remain mine.
8. As a User, I want my old-system password and tokens to remain with the old application, so that the two systems do not share credentials.
9. As a User, I want a login failure to leave me unauthenticated rather than create a partly linked account, so that I cannot enter another person's Response.
10. As a User, I want a matching email to trigger a review instead of automatically merging accounts, so that my existing Folio data cannot be attached to the wrong Legacy Identity.
11. As a User, I want to see that a Pending Account Link requires Admin approval, so that I know why I cannot enter yet and when to retry.
12. As an Admin, I want to review verified pending Legacy Identities and their candidate Folio Users, so that I can approve only a genuinely matching owner.
13. As an Admin, I want an approved link to take effect on the User's next SSO attempt, so that a stale one-time code is not reused.
14. As an Admin, I want to keep Legacy Identities away from Admin accounts, so that an old-system login never grants Folio Admin access.
15. As an Admin, I want a disabled Folio User blocked even with a valid old-system login, so that Folio's account controls remain authoritative.
16. As a User, I want an SSO-created Folio account without a second password, so that the new path does not add credentials to manage.
17. As a User, I want a locally provisioned account to keep its local login, so that adding SSO does not replace existing access.
18. As a User, I want Folio to retain my existing name and email after a Legacy Identity is linked, so that old-system profile changes do not silently rewrite my local account.
19. As a User, I want a valid SSO login to work when my unused local temporary password still requires a change, so that I need not know that password to use SSO.
20. As a User, I want local login with that temporary password still to require a password change, so that SSO does not weaken local credential rules.
21. As a User, I want Folio to warn me when it is logged into a different User than my Legacy Identity, so that I do not save a Response under the wrong account.
22. As a User, I want to confirm any switch between those Folio accounts, so that an existing session is not silently replaced.
23. As a User, I want to save my Draft and repeat SSO when my one-hour Folio Session expires, so that I can resume without local password login if my old login is still valid.
24. As a User, I want old-system logout not to be presented as Folio logout, so that I understand the two Sessions are independent.
25. As an Admin, I want to choose one Fill Method per Form, so that respondents have a predictable way to complete it.
26. As a User, I want a native web form built from the Published Template's Fields, so that I can fill a Form without editing a DOCX in ONLYOFFICE.
27. As a User, I want every Field on one scrollable page in document order, so that I can see all questions and revisit an earlier answer.
28. As a User, I want a Field's Title/Alias used as its label and its Placeholder used as guidance, so that I can understand what to enter.
29. As a User, I want the Field Tag used as the label when Title/Alias is blank, so that every question remains identifiable.
30. As a User, I want static DOCX text and layout omitted from the native web form, so that the web page remains a focused Field questionnaire.
31. As a User, I want text, checkbox, date, dropdown, combo, and picture Fields to work in the native page, so that my Form does not lose a Field because of its type.
32. As a User, I want required and Prefill-locked Fields enforced in either Fill Method, so that the Form has one consistent rule set.
33. As a User, I want to save and resume a native Draft, so that leaving the web page does not lose successfully saved answers.
34. As a User, I want a native Submission to produce the same authoritative DOCX and PDF outcomes as an ONLYOFFICE Submission, so that the method does not change the resulting record.
35. As a User, I want an existing Draft to reopen in the Admin's current Fill Method after a method change, so that my saved answers remain usable.
36. As a User, I want my Receipt to show each Field label above its value, so that I can read the result on narrow or wide screens without interpreting JSON.
37. As a User, I want the latest effective revision opened on my Receipt by default, so that I see the current record first.
38. As a User, I want to select the original or any intermediate Correction, so that I can see how my Submission changed.
39. As a User, I want DOCX/PDF downloads to match the revision I selected, so that downloaded documents agree with displayed answers.
40. As an Admin, I want Result lists and detail screens to present readable Draft and Submission information, so that I can find and inspect work without raw JSON.
41. As an Admin, I want a read-only ONLYOFFICE document tab in a Result detail, so that I can inspect the filled DOCX without modifying it.
42. As an Admin, I want a second tab with Field/value rows, so that I can scan answers without reading the whole document.
43. As an Admin, I want one revision selector for both result tabs and downloads, so that document and Field values always represent the same original or Correction.
44. As an Admin, I want the latest revision opened by default while every stored revision remains selectable, so that I can review both current and historical results.
45. As an Admin, I want to inspect a saved Draft read-only as well as a Submission, so that I can support a User before completion.
46. As an Admin, I want Correction kept as a separate action rather than an editable result tab, so that inspection does not accidentally change a record.
47. As a User, I want no JSON viewer or JSON download button on result screens, so that I see only meaningful document and Field representations.
48. As an Admin, I want no JSON viewer or JSON download button on result screens, so that operational work uses readable representations while internal JSON remains intact.
49. As an Admin, I want Title/Alias labels retained with the applicable Template/Response data, so that older Results remain understandable after the Form changes.
50. As a User, I want Field configuration controls hidden while filling in ONLYOFFICE, so that I see only respondent actions.
51. As an Admin, I want the Field-settings panel to scroll when its contents exceed the available height, so that I can reach every setting.
52. As a User, I want an in-page button to expand the ONLYOFFICE iframe across the page body, so that I can work without surrounding page chrome.
53. As an Admin, I want the same expanded mode on every ONLYOFFICE editor/viewer surface, so that Form design, Result inspection, and Correction feel consistent.
54. As a User, I want to leave expanded mode without invoking browser fullscreen, so that the browser remains under my control.
55. As an Admin, I want to copy the temporary password shown after User creation, so that I can deliver it accurately.
56. As an Admin, I want the same copy action after resetting a User password, so that I do not have to transcribe it.
57. As an Admin, I want an AI chat inside Folio that accepts a PDF or a text description, so that I can begin authoring a Form from source material.
58. As an Admin, I want to request revisions repeatedly in the same chat, so that the generated Form can improve without starting over.
59. As an Admin, I want the AI to revisit the original PDF during later instructions, so that corrections do not depend only on a first-pass summary.
60. As an Admin, I want the AI to produce a DOCX with static text and valid tagged content controls, so that I can use it as a Folio Template Draft.
61. As an Admin, I want a read-only DOCX preview beside the chat, so that I can inspect the current generation before downloading.
62. As an Admin, I want to download the generated DOCX and upload it through Folio's existing Form creation flow, so that I can finish it in ONLYOFFICE before publication.
63. As an Admin, I want a notice before PDF or chat content is sent outside Folio, so that I understand OmniRoute/provider processing and their separate retention policies.
64. As an Admin, I want a refreshed chat page to restore my active authoring session, so that an accidental reload does not lose the draft document.
65. As an Admin, I want an explicit end action and automatic end after logout or two hours of inactivity, so that Folio does not retain temporary files and chat longer than intended.
66. As an Admin, I want the authoring feature unavailable when no permitted service credential or reachable gateway exists, so that it never silently uses a personal login as a shared backend.
67. As a User, I want AI authoring unavailable to my role, so that only Form authors can use document generation tools.
68. As an Operator, I want model-generated Python confined to an ephemeral document workspace without shell or network access, so that untrusted code cannot reach application data or the internet.
69. As an Operator, I want bounded worker time, memory, and file access, so that document generation cannot exhaust or escape the host.
70. As an Operator, I want uploaded PDFs sent to a configured PDF-capable OmniRoute route rather than OCR run in Folio, so that the chosen model handles the original document.
71. As an Operator, I want a pinned static web server image with SPA fallback and internal port 80, so that deployment continues to serve every existing route through the gateway.
72. As a User, I want Folio pages protected against framing by external sites, so that another site cannot embed the application around my interactions.
73. As an Operator, I want the old-system backend API contract documented while that repo remains unchanged, so that Folio's SSO side can be tested now and integrated later.

## Implementation Decisions

- **Baseline and scope:** Extend the currently implemented Folio Forms lifecycle rather than rebuilding issue #1. The current Product decisions and accepted AI/SSO ADRs take precedence where the older baseline issue differs, especially SSO inclusion, all-revision viewing, and removing JSON actions from the UI.
- **Domain identity:** Retain Folio User IDs as owners of Responses. Represent a Legacy Identity by a configured provider ID and non-reassignable old-system subject using the existing unique provider/account mapping. Never identify or attach an account using only an email match.
- **SSO endpoints:** Folio initiates a browser-bound login transaction at `/auth/legacy/start`, preserves only a validated same-origin Form return path, receives `/auth/legacy/callback`, and exchanges a short-lived code server-to-server with the old backend. The old frontend later adds a same-origin authorize page using its existing bearer/refresh helper; its backend later adds authenticated code issuance and client-authenticated atomic exchange.
- **SSO security:** Bind code to client ID, exact allowlisted callback, and PKCE S256 challenge. Use state plus a secure pre-login cookie to bind browser and transaction; keep the verifier and client secret server-side. Expire codes after about 60 seconds and consume once atomically. Use HTTPS; do not log codes, passwords, credentials, or full callback URLs. Never put an old credential or new Folio session bearer in a URL.
- **Session delivery:** Create a normal Folio Session using its existing authentication system. Add a secure browser handoff for that Session: current API authentication expects a bearer token and current UI stores it locally, so a callback redirect alone cannot complete login safely. Preserve existing authorization checks and regenerate the Session upon successful authentication.
- **Provisioning:** Require a stable subject and verified email from old-system exchange. If no local email collision exists, create a role=user Folio User with no local password and bind the Legacy Identity in one durable transaction. If identity or verified email is absent, deny without partial account creation.
- **Account collision:** A verified identity whose email belongs to an unlinked local User creates a Pending Account Link but gains no Session. Expose the verified request to an Admin for an explicit identity/owner check and approval; require a fresh SSO attempt after approval. Prevent linking to an Admin or disabled User. Keep profile name/email unchanged on subsequent SSO logins.
- **Authorization:** Allow SSO only for role=user accounts. Keep Admin accounts on local login, local self-registration disabled, and existing local credential login intact. Distinguish SSO from local password login when enforcing temporary-password change; never bypass other account, Form, Prefill, ownership, or enabled-state checks.
- **Browser account switch:** If the browser has a live Session for a different User, display both identities and require a deliberate switch. Do not grant or replace the Session silently. If the same User is already signed in, preserve their intended Form destination.
- **Expiry:** On Folio's one-hour Session expiry, save the active Draft before leaving, repeat SSO if old login remains valid, then return to the same Form. An expired old login is handled at the old origin. SSO does not synchronize logout or immediately revoke Folio Sessions after old-system disable.
- **Prefill boundary:** SSO supports ordinary Form links only. Existing Prefill Handoffs remain email-bound and retain their current login and claim behavior; do not silently treat a prefill code as an SSO code or bypass Prefill requirements.
- **Fill Method:** Store one Admin-controlled method per Form. The active method applies to both existing Drafts and new Responses; map the same persisted Field data and Prefill snapshot between native and ONLYOFFICE editing. Preserve the canonical DOCX for every save and Submission.
- **Native presentation:** Build one scrollable page from the Published Template's ordered Field Manifest. Support text, checkbox, date, dropdown, combo, and picture; required, option, size, and locked-Prefill rules match the existing server contract. Render Title/Alias as label, Tag as fallback and stable identity, Placeholder as help; omit static DOCX paragraphs and layout from the web page.
- **Native persistence:** Native Save Draft and Submit use the existing Response/Operation authorization and validation boundaries and create coherent Field data plus DOCX artifacts. A failed document update cannot leave a Draft whose values and DOCX disagree. Keep PDF generation consistent with the existing on-demand export model.
- **Field metadata:** Capture display label, placeholder, type, options, and document order at publication and keep enough revision-specific metadata to label historical Drafts, Submissions, and Corrections without reading the current mutable Template Draft. Do not change the exact Tag identity used by Prefill or flat stored data.
- **Results:** Update User Receipt and all Admin result/detail surfaces, including the Form-specific Submission detail, to remove JSON UI. Render a stacked label/value Receipt and a read-only ONLYOFFICE DOCX plus Field table in Admin detail; keep Correction as an independent edit action. Preserve owner/Admin access and audit behavior.
- **Revision contract:** Default Receipt and Admin Results to the latest effective revision; select original and every intermediate Correction as well. Extend current original/latest-only data/document/export selection so selected Field values, read-only document, and DOCX/PDF downloads always match. Drafts have only their current saved state, not Correction history. Preserve internal JSON storage/API where needed; remove only JSON UI actions.
- **ONLYOFFICE surfaces:** Respondent fill mode does not show Admin Field settings. Make the Admin settings panel scroll within available height. Provide a reversible in-page body-sized iframe mode on every ONLYOFFICE surface without requesting browser fullscreen.
- **Admin credentials:** Add clipboard Copy to the existing one-time temporary-password display after account creation and reset; do not make those passwords retrievable later.
- **AI orchestration:** Use Pi for the Admin-only conversation and bounded document tools. Use OmniRoute as Folio's sole model gateway; Folio does not select upstream providers. Keep gateway endpoint and API key on the server. Send original PDF directly to a PDF-capable OmniRoute endpoint and allow reinspection in later turns; Pi's standard prompt does not accept raw PDFs. Do not replace the PDF with a one-time text summary or local OCR.
- **AI output:** Build the generated DOCX with static text and supported tagged content controls using document tooling rather than accepting arbitrary model prose as a valid Form. Validate it against Folio's existing upload/publication constraints before presenting it as ready for Admin refinement; allow iterative chat edits, read-only DOCX preview, and manual download/upload into the normal Template Draft flow.
- **Python boundary:** Model-generated Python executes only in a separate ephemeral document worker with access restricted to active-session files, no shell or network access, and enforceable CPU/time/memory limits. Prompt instructions are not a sandbox. No unrestricted execution in the web/API process.
- **AI privacy and lifecycle:** Disclose external PDF/chat processing before first send; the upstream gateway/provider manages its own retention. Keep Folio's chat, PDF, and generated files for the active session only, resume after refresh, and remove Folio copies on explicit end, logout, or two hours of inactivity. Do not promise upstream deletion.
- **AI credential:** Shared Admin AI stays unavailable until a permitted service credential and PDF/tool-capable OmniRoute route are configured. Personal Codex through OmniRoute is limited to the account owner's local, single-user test; it is not a shared backend credential.
- **Static deployment:** Replace Nginx with a pinned `joseluisq/static-web-server` image, keep the existing internal port 80 and SPA fallback, maintain proxy routing, and set frame policy so external sites cannot embed Folio. Do not copy sample wildcard `frame-ancestors` or `ALLOWALL` headers.

## Testing Decisions

- A good permanent test observes an authorized person's HTTP result, Form/Response state transition, document content, browser-visible action, or security rejection. It fails for a plausible user-visible bug. Do not test private helper names, copied wiring, HTML source text, mock echoes, CSS implementation, or exact UI wording.
- **Primary seam:** Extend the existing complete Elysia HTTP application tests, with real Better Auth, Prisma/PostgreSQL, and private document storage, for SSO, linking, native Draft/Submission, AI session routes, revised Results, and arbitrary revision exports. This is the highest reusable seam; avoid new lower-level suites for the same behavior.
- For SSO, a deterministic local implementation of the old backend's issue/exchange contract exercises success, expired/replayed/wrong-client/wrong-callback/wrong-verifier codes, mismatched state or cookie, cross-browser callback, local email collision, pending-link approval, wrong account, disabled/Admin accounts, safe redirects, session expiry, and owner isolation. Assert that no session is granted before verified exchange or approval, and that no bearer appears in redirect URLs/logs.
- For native filling, publish controlled DOCX fixtures with every supported Field type and metadata; save and reopen across method changes, preserve locked Prefill and picture content, reject invalid required/option/size values, and prove selected DOCX/PDF outputs agree with stored Field values. Extend the existing HTTP fixture/ONLYOFFICE fake path, not a mock-only form renderer test.
- For Results, exercise Draft and Submission reads, owner/Admin authorization, all original/intermediate/latest Corrections, document/data/export consistency, latest default, and JSON UI removal. Behavior of the actual Receipt, tabs, selectors, Field labels, pictures, and download buttons requires a browser smoke run; a test of raw JSON alone does not prove those surfaces.
- For AI, test the HTTP authoring flow with a controlled OmniRoute-compatible service that consumes real PDF bytes and returns changing document instructions and tool calls. Assert repeated PDF inspection, valid generated DOCX with static text and content controls, revision through chat, preview/download, consent, Admin-only access, service-unavailable behavior, refresh recovery, and timely deletion after explicit end/logout/inactivity. Do not present a canned DOCX or mock echo as generation proof.
- Exercise the actual isolated Python worker with generated document tasks and adversarial file/network/shell/resource attempts. Its failure must leave other Users, Sessions, files, and canonical Forms untouched. Run at least one real generated DOCX through the existing Template Draft upload/publication flow.
- Extend the existing ONLYOFFICE plugin VM contract only where its behavior changes: Field-panel visibility, metadata extraction, and secure bridge behavior. Reuse the existing web lifecycle test patterns for safe return paths and save-before-reauthentication. Verify scrolling and in-page expanded mode on the actual browser surface rather than testing CSS text.
- Run a browser smoke with the real application to observe a User filling natively, switching methods on a Draft, reading Receipt revisions, an Admin opening both Result tabs, generating and refining a DOCX, and the old-system login button. Verify the web image serves deep SPA routes on internal port 80 and rejects external framing. A full old-system end-to-end smoke requires its future authorize and issue/exchange implementation; until then prove Folio's external contract against the controlled service and report that limit.
- Existing prior art: the current black-box HTTP acceptance suite already exercises authenticated routes, documents, Prefill, and Operations; the plugin VM tests exercise ONLYOFFICE callbacks/controls; focused web lifecycle tests cover return paths and save ordering. Reuse these seams before introducing a browser test framework or a second application test harness.

## Out of Scope

- Moving both applications to a central identity provider, sharing user tables, passing the old access/refresh token to Folio, or merging accounts on email match alone.
- Implementing the old application's frontend bridge or code issue/exchange APIs in this repository; they are a future change in its separate repository. Do not claim a cross-repository live login smoke until that side exists.
- SSO for Admins, public self-registration, automatic profile synchronization, local passwords for SSO-created Users, and synchronized logout or instant revocation of a Folio Session when the old account changes.
- Combining SSO with the existing email-bound Prefill Handoff. Existing configured Prefill requirements and their login behavior remain in place.
- Anonymous Form filling, a second Response per User/Form, changing the immutable original Submission, or replacing DOCX as the canonical Form/output.
- Showing full DOCX static text or document layout in the native web form; a Typeform-style one-question-at-a-time wizard; letting respondents select a Fill Method contrary to the Form setting.
- JSON displays and download buttons in the new result UI; deleting internal JSON records or required API compatibility solely to hide JSON from people.
- General provider selection in Folio, per-Admin AI credentials, unrestricted agent shell/network access, local OCR, guaranteed upstream deletion, or a shared backend using a personal Codex login.
- External-site framing, browser fullscreen, changing gateway upstream ports, or deploying a new static app origin.
- Reimplementing the accepted MMVP baseline in issue #1; only changes required by these new behaviors belong here.

## Further Notes

- Domain terms follow the current Folio glossary. Respect the accepted AI gateway/worker and legacy one-time-code ADRs. This issue adds behavior to the existing app; it supersedes issue #1 only where current accepted product decisions differ (SSO, all-revision viewing, JSON UI removal, and the Admin-selectable native Fill Method).
- Verified old-system subject and email, backend client authentication, and fixed callback agreement are prerequisites for live SSO. A reachable PDF/tool-capable OmniRoute route and permitted service credential are prerequisites for live shared AI. External services were not available during planning; do not claim their full end-to-end paths were exercised until they exist.
- The single SSO transaction is not the existing Prefill Handoff. Their codes, cookies, trust boundaries, and expiry are distinct.
