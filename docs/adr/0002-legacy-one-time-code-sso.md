# ADR 0002: Add a one-time-code bridge for legacy login

**Status:** Accepted

## Context

The old application owns its existing login, including a browser-held access token and refresh cookie. Folio Forms already has its own Users, password login, one-hour sessions, and Form ownership by local User ID. A centralized identity provider and a shared user database are not available. Sending the old token to Folio or linking accounts by matching email would blur the trust boundary and risk assigning a Response to the wrong person.

## Decision

- Add legacy login as an optional path for ordinary Folio Users. Keep local password login and Admin-only account creation; do not enable public self-registration. Do not grant Folio Admin rights through SSO.
- The old application's same-origin bridge page uses its existing login or refresh helper. Its backend issues a short-lived, single-use code after authenticating the old User. Folio starts a browser-bound transaction, receives the callback, and exchanges the code directly with the old backend using a server-held client credential. Bind the code to the client, exact callback, and PKCE S256 challenge; validate state and pre-login cookie; consume the code atomically. Use HTTPS and do not log credentials or put Folio's session bearer in a URL.
- The old backend returns a stable, non-reassignable subject and a verified email. Folio pairs that subject with its configured legacy provider ID, maps the pair to one local User, and issues its own session through the existing authentication system. The existing Account model already has a unique `(providerId, accountId)` pair; use that identity boundary rather than matching on email. A first-time identity with a free, verified email gets a role=user account without a local credential.
- If the email belongs to an unlinked local User, deny the login and create a Pending Account Link from the verified old identity. An Admin verifies the ownership of the local User and approves the link; the User starts SSO again. Never attach a Legacy Identity to an Admin account. A disabled Folio User stays blocked even if the old login works.
- Keep Folio's name and email unchanged on subsequent legacy logins. A local User with a temporary-password-change requirement may enter through SSO without changing that password; local password login still enforces the change. Keep the two login methods distinguishable without bypassing other authorization checks.
- Warn and require confirmation if the browser already has a Folio session for a different User. Provide a legacy SSO button on Folio's login page. On expiry of Folio's one-hour session, save the active Form Draft and repeat the bridge if the old login remains valid; otherwise the old application handles login before returning to the Form.
- Support direct Form entry only. Do not merge this flow with the existing email-bound prefill handoff. Synchronizing logout or immediate old-system revocation is out of scope.

## Consequences

- Browser navigation must preserve only validated, local Form or dashboard return paths. On success, the callback redirects to the configured `CORS_ORIGIN` web origin, not the API callback origin. The redirect origin comes only from server configuration, never request `Host`, forwarded headers, or user input. Failures return to the fixed login failure path on that same web origin. The callback must not expose the Folio session bearer in a redirect URL; current Folio API requests require a bearer token, so the browser session handoff needs a secure delivery path.
- Folio may implement its side first, but end-to-end verification requires the old application's authorize page and issue/exchange APIs. The old-system API contract must supply a stable subject and verified email; no self-registration, automatic email linking, or placeholder identity may silently replace them.
- The existing prefill handoff still requires its own Folio login state and email match. SSO for a direct Form link does not remove that requirement.

## Required old-backend HTTP contract

Configure `LEGACY_SSO_AUTHORIZE_URL`, `LEGACY_SSO_EXCHANGE_URL`, `LEGACY_SSO_CALLBACK_URL`, `LEGACY_SSO_CLIENT_ID`, `LEGACY_SSO_CLIENT_SECRET`, and `LEGACY_SSO_PROVIDER_ID` together on Folio's server. The callback URL must be the exact public Folio API URI. `CORS_ORIGIN` identifies the trusted Folio web origin for redirects; direct local web development uses `http://localhost:5173`, while same-origin deployments use their public origin. Use HTTPS outside localhost and keep the client secret server-side.

The authorization bridge accepts a top-level GET with `client_id`, exact `redirect_uri`, unpredictable `state`, `code_challenge`, and `code_challenge_method=S256`. It authenticates the old User with the old application's existing login or refresh flow. It issues a cryptographically random, short-lived, single-use authorization code bound to client ID, exact callback URI, and PKCE challenge. Treat the code as opaque; Folio accepts standard Base64 characters such as `+`, `/`, and `=` and form-URL-encodes the exact value during exchange. The bridge redirects to the exact callback with only `code` and `state`; it never returns an access token or refresh token in the browser.

Folio redeems the code directly from its server using an HTTPS POST to the exchange URL. The request uses HTTP Basic authentication for the configured client ID and secret, plus form fields `grant_type=authorization_code`, `client_id`, `code`, `redirect_uri`, and `code_verifier`. The backend atomically rejects expired, replayed, unknown, or mismatched codes. It returns JSON with a stable, non-reassignable `sub`, `email`, and `email_verified: true`. It does not return legacy access or refresh credentials. Neither side logs codes, credentials, or tokens.

Ticket 12 accepts only identities already mapped by `(providerId, accountId)` to an enabled `role=user` Folio account whose email matches the verified provider email. First-time account provisioning and pending links remain Tickets 13–14. Account-switch confirmation and draft re-entry remain Tickets 15–16.
