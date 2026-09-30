# Issue tracker: Local Markdown

Issues and specs live as Markdown files under `.scratch/`.

## Layout

- One directory per feature: `.scratch/<feature-slug>/`.
- Feature spec: `.scratch/<feature-slug>/spec.md`.
- Implementation ticket: `.scratch/<feature-slug>/issues/<NN>-<slug>.md`, starting at `01`.
- Put `Status: <triage label>` near the top of each ticket.
- Append discussion under `## Comments`.

## Skill operations

- Publish a spec or issue by writing its local Markdown file, not by calling GitHub Issues.
- Fetch a ticket by reading its path; use the supplied path or issue number to locate it.
- For wayfinding, keep `.scratch/<effort>/map.md` and numbered child tickets under `issues/`.
- Wayfinding tickets use `Type: research|prototype|grilling|task` and `Status: claimed|resolved`; `Blocked by: NN, NN` records dependencies.
