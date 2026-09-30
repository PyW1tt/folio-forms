# 20: Isolated AI Document Worker

**What to build:** When the AI chooses generated Python for document work, it runs only in an ephemeral isolated worker; an Admin still receives a validated DOCX usable in the normal Template Draft flow.

**Blocked by:** 17: Text Prompt to Valid DOCX.

**Status:** ready-for-human

- [x] Worker limits file access to current AI Authoring Session inputs/outputs, blocks shell/network, and enforces CPU, time, and memory limits outside the API process.
- [x] Worker output is validated as DOCX with supported tagged controls before preview, download, or Form upload.
- [x] Failed or hostile generated code cannot modify other Users, Sessions, canonical Forms, or unrelated files; Folio removes worker files with session cleanup.
- [x] Exercise real benign and adversarial worker tasks and upload at least one generated DOCX into a Template Draft.

Verification (reset run, attempt 1): The real-worker HTTP scenario passed (1 test, 259 assertions): generation, deliberate revisions using exact OLD edit declarations, malformed/unsupported/mismatched output rejection, host-file/outside-workspace/network/shell/CPU/memory containment, other Users/Sessions/Forms isolation, and Template Draft upload. A metadata-bypass regression failed before the fix (expected 502, received 200) and passed afterward. A real TCP HTTP smoke using the corrected committed application generated a 36,695-byte DOCX, rejected source-only paragraph loss with 502, preserved the prior document, and uploaded the generated DOCX into a Template Draft. Direct worker probes denied native `execve`, rejected oversized output, contained stderr flooding, and stopped a sleeping job after 30.32 seconds. Final inspection found no job containers or source workspaces. Server typecheck and worker-only Ultracite passed. The final available source suite passed 53 tests with 1,793 assertions and 3 known out-of-ticket scenarios filtered out; the existing ticket-17 pending-cleanup diagnostic still appeared. This is not a green full-suite or full AI-module lint claim.

## Standards

The required Standards reviewer re-reviewed committed integration and validation fixes through `1699394`, including the dedicated worker in `a0c229f`: no remaining actionable documented-standard violations or baseline smells.

## Spec

The required Spec reviewer re-reviewed committed integration and validation fixes through `1699394`, including the dedicated worker in `a0c229f`: no remaining actionable requirement gaps, incorrect behavior, or scope creep.

Review summary: Standards 0 remaining findings; Spec 0 remaining findings.

Commit scope: `ed81216` integrates the existing worker and its HTTP scenario; `d434ee7` validates actual Python DOCX text/field metadata; `1699394` corrects the selective projection of parser collectors. The minimum coupled dependencies are the existing template parser's optional text collection, generated-document validator, and two AI route callbacks. Independent native-field and ticket-17/18 changes remain unstaged. Dedicated worker/configuration/documentation remain in `a0c229f`. No ticket-20 blockers remain. The excluded native-picture ZIP-byte comparison and ticket-17 cleanup/shutdown scenarios, plus inherited full AI-module Ultracite diagnostics, remain outside this ticket. Temporary smoke source, isolated verification database, and private verification bucket were removed.
