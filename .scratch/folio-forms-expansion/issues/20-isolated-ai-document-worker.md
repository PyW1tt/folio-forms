# 20: Isolated AI Document Worker

**What to build:** When the AI chooses generated Python for document work, it runs only in an ephemeral isolated worker; an Admin still receives a validated DOCX usable in the normal Template Draft flow.

**Blocked by:** 17: Text Prompt to Valid DOCX.

**Status:** ready-for-agent

- [x] Worker limits file access to current AI Authoring Session inputs/outputs, blocks shell/network, and enforces CPU, time, and memory limits outside the API process.
- [x] Worker output is validated as DOCX with supported tagged controls before preview, download, or Form upload.
- [x] Failed or hostile generated code cannot modify other Users, Sessions, canonical Forms, or unrelated files; Folio removes worker files with session cleanup.
- [x] Exercise real benign and adversarial worker tasks and upload at least one generated DOCX into a Template Draft.

Verification (attempt 2): The real disposable worker generated a 36,689-byte DOCX; native `execve`, oversized-output, and cancellation probes were contained, and no job containers or workspaces remained. The isolated HTTP authoring/upload scenario passed (1 test, 86 assertions), including host-file, outside-workspace, network, shell, CPU, and memory attacks; other Users/Sessions/Forms remained intact and a generated DOCX uploaded into a Template Draft. Review found unbounded Docker logs and abandoned state after sidecar interruption; both were fixed. A real stderr-flood job retained no Docker log file and still generated a valid DOCX; restarting the sidecar removed an orphan container and workspace. Server typecheck and targeted worker-file Ultracite passed. The final full source suite ran with the real worker enabled: 53 pass, 3 pre-existing out-of-ticket failures (`http.test.ts:7877` native picture ZIP-byte equality, `:15478` ticket-17 cleanup-error propagation, `:15673` ticket-17 shutdown fixtures missing `toolState.pending`); the suite also surfaced the existing pending cleanup error.

## Standards

The required Standards reviewer completed with no remaining actionable documented-standard violations or baseline smells.

## Spec

The required Spec reviewer confirmed both worker findings were resolved; no remaining actionable spec findings.

Review summary: Standards 0 remaining findings; Spec 0 remaining findings.

Remaining before ready-for-human: Commit the shared AI-authoring integration and HTTP regression hunks after the interleaved ticket-17/18/19 prerequisites are committed. Dedicated worker implementation, worker-only Compose configuration, documentation, and this ticket are safely separable; staging all shared files would include unrelated prerequisite changes. Preserve the existing partial integration. The three known full-suite failures belong to other tickets and remain outside this ticket's scope.
