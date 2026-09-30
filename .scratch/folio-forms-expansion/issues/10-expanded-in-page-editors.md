# 10: Expanded In-page Editors

**What to build:** Users and Admins expand and collapse ONLYOFFICE within the page body on fill, design, Result review, and Correction surfaces without browser fullscreen.

**Blocked by:** 02: Readable Admin Submission Results.

**Status:** ready-for-human

**Verification:** Real-browser smoke loaded live ONLYOFFICE documents on User fill, Admin design, read-only Result, and Correction. Pointer restoration and keyboard restoration from document focus retained the same iframe on all four surfaces; focus returned to the expand control, Result remained read-only, and background controls recovered. Correction zoom stayed at 110% and its entered reason survived restoration. Browser fullscreen remained inactive. The expanded control and keyboard guidance fit at 320, 375, 414, and 768 px. Web typecheck, scoped Ultracite, and UI detector passed; both required review axes found no blocking issues.

**Keyboard exit:** With focus inside the document, press Alt/Option, then F, then the displayed ONLYOFFICE key tip for “คืนค่าขนาดปกติ” (D on the tested editable surfaces; C on Result). This uses the documented DocsAPI `onRequestClose` event and `customization.close` control to restore the page without closing or remounting the document. Escape still restores when focus is on the parent page; no unsupported vendor keyboard API is used.

- [x] One accessible control expands and restores the iframe on every editor or read-only viewer surface.
- [x] Expansion preserves editing/view state, permissions, and ability to leave the mode from keyboard or pointer.
- [x] Browser smoke covers User fill, Admin design, read-only Result, and Correction without invoking browser fullscreen.

**Suite note:** Final source-suite verification passed 53 tests with 1,598 assertions. Three previously reported, unrelated HTTP failures were filtered out rather than rerun: the broad authenticated workflow's native-picture ZIP equality, Ticket 17 sign-out cleanup, and Ticket 17 shutdown fixture. The existing pending-cleanup diagnostic also appeared; no unrelated code was changed. No ticket-specific blockers remain.
