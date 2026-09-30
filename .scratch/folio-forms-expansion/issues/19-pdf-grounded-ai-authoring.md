# 19: PDF-grounded AI Authoring

**What to build:** An Admin attaches a PDF to an AI Authoring Session and can ask later questions that cause the model to inspect that same original PDF again while refining the generated DOCX.

**Blocked by:** 18: Revise and Resume AI Authoring.

**Status:** ready-for-human

- [x] After informed consent, Folio sends original PDF bytes directly to a configured PDF-capable OmniRoute route; no local OCR or one-time text-only summary substitutes for the PDF.
- [x] The active session retains the source PDF for later model reinspection, including after refresh, and removes it on end/logout/inactivity.
- [x] Admin-only access, unavailable-route behavior, file limits, and safe preview/download remain enforced.
- [x] A controlled service observes real PDF bytes on initial and later turns; a browser smoke refines the resulting DOCX.

## Verification

Reset run, attempt 1: the scoped Ticket 19 HTTP regression passed with 64 assertions, and server/web typechecks passed. The prose-refinement regression failed with HTTP 502 before the fix and passed after replacing English keyword inference with validated document-tool edit declarations. Invalid declarations leave the last valid DOCX and restored session unchanged. The HTTP flow also verifies consent, Admin ownership, unavailable routes, MIME/size limits, authorized download, end/logout/expiry, and blocked-inspection cancellation.

A real browser attached a valid 625-byte PDF, restored the active session after refresh, added a tagged Serial Number field, shortened the static instructions, uploaded the resulting DOCX as a Template Draft, and ended the session. The controlled OmniRoute-compatible service received identical original PDF bytes on all three turns (SHA-256 `7be7576bd76e0ed5d83146b53fda433ee222aa2bf50e74e6a28d96535cefd9b6`). Visual evidence: `/tmp/folio19-reset-declarations-final.png`.

Final parallel reviews: Spec found no remaining defects; Standards found no hard violations and retained four non-blocking duplication/request-context heuristics. Ticket-related code and documentation are committed in `6ac069c`, `31f9a49`, and `5044a20`. Selective staging includes the minimum current-session/revision state and refinement UI blocks structurally required for repeated PDF inspection. Independent prerequisite changes, Python-worker integration, and other tickets remain untouched in the worktree.

The final filtered source run passed 52 tests with 1,522 assertions and excluded three previously observed unrelated failures: native picture ZIP-byte equality, Ticket 17 cleanup-error propagation, and Ticket 17 shutdown fixtures missing `toolState.pending`. The earlier unfiltered run had 52 passing and 3 failing tests; those failures were not rerun solely to confirm them or changed here. This is not a green unfiltered-suite claim.

**Remaining:** No Ticket 19 blocker. No live third-party provider verification is claimed. The isolated browser API did not configure ONLYOFFICE, so Template Draft upload was verified but its editor was not. Controlled-service acceptance is complete. Temporary smoke code, browser sessions, and isolated services were removed after verification.
