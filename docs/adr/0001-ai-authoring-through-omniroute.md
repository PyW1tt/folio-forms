# ADR 0001: Route AI authoring through OmniRoute

**Status:** Accepted

## Context

Folio Forms needs an Admin-only authoring flow that accepts chat instructions or a PDF, supports iterative edits, and produces a DOCX for the existing upload flow. The flow needs an agent session and document tools, while upstream model and provider selection should remain outside Folio Forms.

The original PDF must go to a model without OCR in Folio. Model-generated Python is allowed for document work, but uploaded files and generated code are untrusted. The app must delete its session data without promising deletion by external services.

## Decision

- Use Pi for agent sessions and tool orchestration. Send model requests through OmniRoute; Folio Forms does not select an upstream provider or store provider credentials.
- Treat OmniRoute as an externally managed gateway. The server holds its endpoint and API key. Provider and model selection remain in OmniRoute.
- Accept PDF uploads in Folio's AI chat and keep the original in the active authoring session. Send it from the Folio server to a PDF-capable OmniRoute endpoint for model inspection, including later requests that need to revisit the PDF; do not perform OCR or rely on a one-time summary. Pass findings into the Pi conversation for document work. Pi's standard `prompt()` accepts images, not raw PDF attachments; the configured route must support PDF input and subsequent tool calls.
- Run model-generated Python only in an isolated, ephemeral document worker. Limit file access to the active session, deny shell and network access, and enforce resource and time limits. Model instructions are not a security boundary.
- Restrict AI authoring to Admins. Before the first external send, disclose that chat and PDF content go to OmniRoute and its configured provider. Let an Admin resume after page refresh; end the session on explicit close, logout, or two hours of inactivity, then delete Folio's copies. Upstream retention follows OmniRoute and provider policies.
- Permit personal Codex credentials only for the account owner's local, single-user tests. Keep shared AI unavailable until OmniRoute has a permitted service credential.

## Consequences

- Folio Forms has one model-gateway integration, but its configured route must support both direct PDF input and the subsequent Pi tool-calling conversation. Verify both capabilities end-to-end against the selected OmniRoute route.
- AI authoring depends on a reachable OmniRoute service and its configuration. No service listened at `localhost:20128` in the development environment during planning; end-to-end AI verification requires a reachable endpoint and permitted credential.
- Folio can delete its own session files but cannot guarantee deletion of upstream logs or files.
- A personal Codex subscription cannot serve as the shared credential for multiple Admins.

## References

- [OmniRoute API Reference](https://github.com/diegosouzapw/OmniRoute/wiki/API-Reference)
- [Pi SDK prompt attachments](https://github.com/earendil-works/pi/blob/main/packages/coding-agent/src/core/agent-session.ts)
- [OpenAI file inputs](https://developers.openai.com/api/docs/guides/file-inputs)
- [OpenAI Account Sharing Policy](https://help.openai.com/en/articles/10471989-openai-account-sharing-policy)
