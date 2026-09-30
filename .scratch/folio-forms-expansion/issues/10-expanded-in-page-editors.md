# 10: Expanded In-page Editors

**What to build:** Users and Admins expand and collapse ONLYOFFICE within the page body on fill, design, Result review, and Correction surfaces without browser fullscreen.

**Blocked by:** 02: Readable Admin Submission Results.

**Status:** ready-for-human

**Verification:** Real browser smoke with valid local ONLYOFFICE configuration loaded Admin design, User fill, read-only Admin Result, and Correction documents. The same iframe stayed mounted through expand and restore on all four surfaces; Correction zoom remained at 110% after restoration, and Result stayed read-only. The restore button worked by pointer and by Enter when focused. Browser fullscreen remained inactive. Escape works while the parent page has focus, not while focus is inside ONLYOFFICE's iframe.

- [x] One accessible control expands and restores the iframe on every editor or read-only viewer surface.
- [x] Expansion preserves editing/view state, permissions, and ability to leave the mode from keyboard or pointer.
- [x] Browser smoke covers User fill, Admin design, read-only Result, and Correction without invoking browser fullscreen.
