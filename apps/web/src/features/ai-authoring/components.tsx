import { Download, FileCheck2, WandSparkles } from "lucide-react";
import type { ChangeEvent, FormEvent, RefObject } from "react";

import {
  Button,
  Card,
  Input,
  Notice,
  Spinner,
  Textarea,
} from "@/components/ui";

import type { AiAuthoringStatus, AuthoringPreview, BusyAction } from "./types";

interface RevisionAuthoringFormProps {
  busyAction: BusyAction;
  consent: boolean;
  preview: AuthoringPreview;
  revisionPrompt: string;
  reviseDocument: (event: FormEvent<HTMLFormElement>) => Promise<void>;
  setConsent: (value: boolean) => void;
  setRevisionPrompt: (value: string) => void;
  status: AiAuthoringStatus | null;
}

export const RevisionAuthoringForm = ({
  busyAction,
  consent,
  preview,
  revisionPrompt,
  reviseDocument,
  setConsent,
  setRevisionPrompt,
  status,
}: RevisionAuthoringFormProps) => {
  const busy = busyAction !== null;
  return (
    <Card className="space-y-5 p-5 sm:p-7">
      <h2 className="text-xl font-bold">Refine this document</h2>
      <form
        aria-busy={busyAction === "revision"}
        className="space-y-5"
        onSubmit={reviseDocument}
      >
        <label className="block space-y-2">
          <span className="font-semibold">Next instruction</span>
          <Textarea
            disabled={!status?.enabled || busy}
            maxLength={16_000}
            onChange={(event) => setRevisionPrompt(event.target.value)}
            placeholder="For example: Add an approval date and make the equipment list clearer."
            required
            rows={4}
            value={revisionPrompt}
          />
        </label>
        <Notice>
          {status?.disclosure ??
            "Your prompt and document content go to OmniRoute and its configured provider. Folio Forms cannot promise upstream deletion."}
          {preview.hasSourcePdf
            ? " The original source PDF is sent again with each revision."
            : null}
        </Notice>
        <label className="flex items-start gap-3 text-sm">
          <input
            checked={consent}
            className="mt-1 size-4 accent-[var(--ink)]"
            disabled={!status?.enabled || busy}
            onChange={(event) => setConsent(event.target.checked)}
            required
            type="checkbox"
          />
          <span>
            I consent to sending this instruction, document content
            {preview.hasSourcePdf ? ", and the original PDF" : ""} to OmniRoute
            and its configured provider.
          </span>
        </label>
        <Button
          disabled={
            !status?.enabled || busy || !consent || !revisionPrompt.trim()
          }
          type="submit"
        >
          {busyAction === "revision" ? <Spinner /> : <WandSparkles size={16} />}
          {busyAction === "revision" ? "Updating DOCX…" : "Update DOCX"}
        </Button>
      </form>
    </Card>
  );
};

interface ExistingAuthoringSessionProps {
  busyAction: BusyAction;
  consent: boolean;
  download: () => Promise<void>;
  endSession: () => Promise<void>;
  preview: AuthoringPreview;
  revisionPrompt: string;
  reviseDocument: (event: FormEvent<HTMLFormElement>) => Promise<void>;
  setConsent: (value: boolean) => void;
  setRevisionPrompt: (value: string) => void;
  status: AiAuthoringStatus | null;
  uploadTemplateDraft: () => Promise<void>;
}

export const ExistingAuthoringSession = ({
  busyAction,
  consent,
  download,
  endSession,
  preview,
  revisionPrompt,
  reviseDocument,
  setConsent,
  setRevisionPrompt,
  status,
  uploadTemplateDraft,
}: ExistingAuthoringSessionProps) => {
  const busy = busyAction !== null;
  return (
    <div className="space-y-5">
      <section aria-label="Authoring conversation" className="space-y-4">
        <h2 className="text-xl font-bold">Conversation</h2>
        {preview.hasSourcePdf ? (
          <p className="text-sm text-[var(--ink-soft)]">
            Source PDF retained for this session. Each revision re-inspects the
            original PDF; no re-upload needed.
          </p>
        ) : null}
        <ol className="space-y-4">
          {preview.turns.map((turn, index) => (
            <li className="space-y-3" key={`${index}-${turn.prompt}`}>
              <div>
                <h3 className="text-sm font-semibold">Your instruction</h3>
                <p className="whitespace-pre-wrap break-words">{turn.prompt}</p>
              </div>
              <div>
                <h3 className="text-sm font-semibold">AI response</h3>
                <p className="whitespace-pre-wrap break-words text-[var(--ink-soft)]">
                  {turn.assistantMessage}
                </p>
              </div>
            </li>
          ))}
        </ol>
      </section>
      <RevisionAuthoringForm
        busyAction={busyAction}
        consent={consent}
        preview={preview}
        revisionPrompt={revisionPrompt}
        reviseDocument={reviseDocument}
        setConsent={setConsent}
        setRevisionPrompt={setRevisionPrompt}
        status={status}
      />
      <Card
        className="space-y-5 p-5 sm:p-7"
        aria-label="Read-only DOCX preview"
      >
        <div>
          <h2 className="text-2xl font-bold">{preview.title}</h2>
          {preview.description ? (
            <p className="mt-2 text-[var(--ink-soft)]">{preview.description}</p>
          ) : null}
        </div>
        {preview.paragraphs.map((paragraph, index) => (
          <p
            key={`${index}-${paragraph}`}
            className="whitespace-pre-wrap text-[var(--ink-soft)]"
          >
            {paragraph}
          </p>
        ))}
        <div className="space-y-4">
          {preview.fields.map((field) => (
            <label key={field.tag} className="block space-y-1.5">
              <span className="text-sm font-semibold">{field.label}</span>
              <Input
                aria-label={`${field.label} (read-only preview)`}
                placeholder={field.placeholder}
                readOnly
                value=""
              />
              <span className="block text-xs text-[var(--ink-soft)]">
                Tag: {field.tag}
              </span>
            </label>
          ))}
        </div>
      </Card>
      <div className="flex flex-wrap gap-3">
        <Button
          disabled={busy}
          onClick={download}
          type="button"
          variant="secondary"
        >
          {busyAction === "download" ? <Spinner /> : <Download size={16} />}
          {busyAction === "download" ? "Downloading DOCX…" : "Download DOCX"}
        </Button>
        <Button disabled={busy} onClick={uploadTemplateDraft} type="button">
          {busyAction === "upload" ? <Spinner /> : <FileCheck2 size={16} />}
          {busyAction === "upload"
            ? "Uploading Template Draft…"
            : "Upload as Template Draft"}
        </Button>
        <Button
          disabled={busy}
          onClick={endSession}
          type="button"
          variant="ghost"
        >
          {busyAction === "end"
            ? "Ending session…"
            : "End session and delete Folio copy"}
        </Button>
      </div>
    </div>
  );
};

interface CreateAuthoringFormProps {
  busyAction: BusyAction;
  consent: boolean;
  createDocument: (event: FormEvent<HTMLFormElement>) => Promise<void>;
  prompt: string;
  selectSourcePdf: (event: ChangeEvent<HTMLInputElement>) => void;
  setConsent: (value: boolean) => void;
  setPrompt: (value: string) => void;
  sourcePdf: File | null;
  sourcePdfInput: RefObject<HTMLInputElement | null>;
  status: AiAuthoringStatus | null;
}

export const CreateAuthoringForm = ({
  busyAction,
  consent,
  createDocument,
  prompt,
  selectSourcePdf,
  setConsent,
  setPrompt,
  sourcePdf,
  sourcePdfInput,
  status,
}: CreateAuthoringFormProps) => {
  const busy = busyAction !== null;
  return (
    <Card className="space-y-5 p-5 sm:p-7">
      {status?.enabled === false ? (
        <Notice>
          AI Authoring is disabled. The server needs an OmniRoute endpoint,
          permitted service credential, and configured model alias.
        </Notice>
      ) : null}
      <form
        aria-busy={busyAction === "create"}
        className="space-y-5"
        onSubmit={createDocument}
      >
        <label className="block space-y-2">
          <span className="font-semibold">Describe your form</span>
          <Textarea
            disabled={!status?.enabled || busy}
            maxLength={16_000}
            onChange={(event) => setPrompt(event.target.value)}
            placeholder="For example: Create an employee equipment request form with name, department, requested items, and approval notes."
            required={!sourcePdf}
            rows={7}
            value={prompt}
          />
          <span className="block text-xs text-[var(--ink-soft)]">
            Optional when attaching a PDF. Without an instruction, we’ll create
            a form based on the PDF.
          </span>
        </label>
        <div>
          <label className="block font-semibold" htmlFor="source-pdf">
            Source PDF (optional)
          </label>
          <input
            accept="application/pdf"
            aria-describedby="source-pdf-hint"
            className="mt-2 min-h-11 w-full rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-[var(--ink)] shadow-sm focus:border-[var(--ink)] focus:outline-none file:mr-3 file:rounded-md file:border-0 file:bg-[var(--accent-soft)] file:px-3 file:py-2 file:font-semibold"
            disabled={!status?.enabled || busy}
            id="source-pdf"
            onChange={selectSourcePdf}
            ref={sourcePdfInput}
            type="file"
          />
          <p
            className="mt-1 text-xs text-[var(--ink-soft)]"
            id="source-pdf-hint"
          >
            PDF up to 10 MiB. The original PDF stays in this session for later
            revisions.
          </p>
        </div>
        <Notice>
          {status?.disclosure ??
            "Your prompt and document content go to OmniRoute and its configured provider. Folio Forms cannot promise upstream deletion."}
          {sourcePdf
            ? " The original PDF is sent now and again with each revision. Folio Forms retains it until this session ends; upstream deletion cannot be promised."
            : null}
        </Notice>
        <label className="flex items-start gap-3 text-sm">
          <input
            checked={consent}
            className="mt-1 size-4 accent-[var(--ink)]"
            disabled={!status?.enabled || busy}
            onChange={(event) => setConsent(event.target.checked)}
            required
            type="checkbox"
          />
          <span>
            I consent to sending this prompt and generated document content to
            OmniRoute and its configured provider.
            {sourcePdf
              ? " I also consent to sending the original PDF now and on future revisions."
              : null}
          </span>
        </label>
        <Button
          disabled={
            !status?.enabled ||
            busy ||
            !consent ||
            (!prompt.trim() && !sourcePdf)
          }
          type="submit"
        >
          {busyAction === "create" ? <Spinner /> : <WandSparkles size={16} />}
          {busyAction === "create" ? "Generating DOCX…" : "Generate DOCX"}
        </Button>
      </form>
    </Card>
  );
};
