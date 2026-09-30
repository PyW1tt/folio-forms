# Domain glossary

## Admin

An authenticated person who creates, edits, publishes, and shares Forms and can inspect every Draft and Submission. Admin includes the form-designer responsibility; there is no separate Designer role.

## User

An authenticated person who fills Forms and can inspect only their own Drafts and Submissions. A User may sign in locally or through a linked Legacy Identity; their Responses remain owned by the same Folio account.

## Legacy Identity

A person identified by a stable subject in the old system and an explicit provider. A Legacy Identity links to at most one Folio User; a matching email alone never establishes the link.

## SSO Bridge

The login handoff in which the old system verifies a Legacy Identity and Folio creates its own session for the linked User. The systems do not share passwords or login tokens.

## Pending Account Link

A verified Legacy Identity awaiting an Admin's explicit approval to associate it with an existing Folio User. It does not grant access before the link is approved.

## Form

A shareable questionnaire-like document owned by the application. A Form has a title, description, Template Draft, optional Published Template, and one stable share link.

## Fill Method

An Admin-selected Form setting that routes all current and new Responses through either the native web form or ONLYOFFICE. Changing the method does not change the shared Fields or saved DOCX.

## AI Authoring Session

An Admin's temporary workspace for creating or editing a Form from a PDF or chat description. It contains an iterative conversation and a generated DOCX that the Admin can preview and download.

## OmniRoute
The external model gateway used by AI Authoring Sessions. Folio calls it server-side; the upstream provider and model are configured in OmniRoute.

## Template Draft

The editable document an Admin is currently preparing. Saving it does not change the document available through the Form's share link.

## Published Template

The current document used to start new Responses. Publishing replaces the previous Published Template and invalidates every unsubmitted Draft for that Form. Published Templates have no user-visible history.

## Field
A tagged content control in a Form. Its Title/Alias is the user-facing label; its Tag is the stable identity in prefill and extracted data; its Placeholder guides entry.
If the Title/Alias is blank, use the Tag as the label. Field tags must be present and unique within a Form.

## Response

A User's single attempt to fill a Form. A User can have at most one Response for a Form. A Response may hold a Draft and may produce one Submission.

## Draft

A manually saved, resumable Response state containing both the current document and its extracted field data. A Draft is not complete unless both representations were saved successfully.

## Submission

The immutable completed result of a Response. A Submission consists of extracted field data, a filled DOCX, and a PDF. It exists only when all three artifacts were persisted successfully.

## Correction
An Admin-authored revision to a submitted Response. It records changed Field values, a reason, and a revised document while preserving the original Submission.


## Operation

A tracked asynchronous request to save a Template Draft, publish a Form, save a Draft, or submit a Response. An Operation progresses through pending, processing, completed, or failed and correlates ONLYOFFICE callbacks with the initiating action.

## Prefill

A snapshot of application-provided Field values and editability policies applied when a Response starts. Resume does not refresh Prefill; starting again after draft invalidation does.
