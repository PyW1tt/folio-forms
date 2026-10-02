import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { Download, FileCheck2, WandSparkles } from "lucide-react";
import type { ChangeEvent, FormEvent, RefObject } from "react";
import { useCallback, useEffect, useRef, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import {
  Button,
  Card,
  Input,
  Notice,
  Spinner,
  Textarea,
} from "@/components/ui";
import {
  ApiError,
  apiDelete,
  apiGet,
  apiGetBlob,
  apiPost,
  apiPostFormData,
  downloadArtifact,
  getToken,
} from "@/lib/api";
import type { FormSummary } from "@/lib/api";
import { useAuth } from "@/lib/auth";

const docxContentType =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const maximumSourcePdfBytes = 10 * 1024 * 1024;
const pdfOnlyInstruction = "Create a form based on the attached PDF.";
interface AiAuthoringStatus {
  disclosure: string;
  enabled: boolean;
}
interface GeneratedField {
  label: string;
  placeholder: string;
  tag: string;
}
interface AuthoringPreview {
  assistantMessage: string;
  description: string;
  downloadUrl: string;
  fields: GeneratedField[];
  hasSourcePdf: boolean;
  paragraphs: string[];
  sessionId: string;
  title: string;
  turns: { prompt: string; assistantMessage: string }[];
}
interface AuthoringStartResponse {
  session: AuthoringPreview;
}
interface CurrentAuthoringResponse {
  session: AuthoringPreview | null;
}
interface CreateTemplateDraftResponse {
  form: FormSummary;
}
type BusyAction = "create" | "revision" | "download" | "upload" | "end" | null;

const authoringErrorMessage = (error: unknown): string => {
  if (!(error instanceof ApiError)) {
    return "AI Authoring is unavailable. Try again later.";
  }
  if (error.status === 413 || error.code === "payload_too_large") {
    return "PDF must be 10 MiB or smaller.";
  }
  if (error.code === "invalid_file_type") {
    return "Choose a valid PDF file.";
  }
  if (error.code === "forbidden") {
    return "Only Admins can use AI Authoring.";
  }
  if (error.code === "consent_required") {
    return "Consent is required before sending content to OmniRoute.";
  }
  if (error.code === "ai_authoring_unavailable") {
    return "AI Authoring is disabled. Ask an administrator to configure OmniRoute.";
  }
  return "AI Authoring could not complete this request.";
};

const docxFilename = (title: string): string =>
  `${title.replaceAll(/[\\/:*?"<>|]/gu, "-").slice(0, 120) || "template"}.docx`;

const RevisionAuthoringForm = ({
  busyAction,
  consent,
  preview,
  revisionPrompt,
  reviseDocument,
  setConsent,
  setRevisionPrompt,
  status,
}: {
  busyAction: BusyAction;
  consent: boolean;
  preview: AuthoringPreview;
  revisionPrompt: string;
  reviseDocument: (event: FormEvent<HTMLFormElement>) => Promise<void>;
  setConsent: (value: boolean) => void;
  setRevisionPrompt: (value: string) => void;
  status: AiAuthoringStatus | null;
}) => {
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

const ExistingAuthoringSession = ({
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
}: {
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
}) => {
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

const CreateAuthoringForm = ({
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
}: {
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
}) => {
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

const AiAuthoringWorkspace = ({ sessionToken }: { sessionToken: string }) => {
  const navigate = useNavigate();
  const [status, setStatus] = useState<AiAuthoringStatus | null>(null);
  const [prompt, setPrompt] = useState("");
  const [sourcePdf, setSourcePdf] = useState<File | null>(null);
  const [revisionPrompt, setRevisionPrompt] = useState("");
  const [consent, setConsent] = useState(false);
  const [preview, setPreview] = useState<AuthoringPreview | null>(null);
  const [loading, setLoading] = useState(true);
  const [busyAction, setBusyAction] = useState<BusyAction>(null);
  const busy = busyAction !== null;
  const [error, setError] = useState<string | null>(null);
  const mounted = useRef(true);
  const sourcePdfInput = useRef<HTMLInputElement>(null);
  const isCurrent = useCallback(
    () => mounted.current && getToken() === sessionToken,
    [sessionToken]
  );

  useEffect(() => {
    mounted.current = true;
    let active = true;
    const load = async () => {
      try {
        const [authoringStatus, current] = await Promise.all([
          apiGet<AiAuthoringStatus>("/api/admin/ai-authoring"),
          apiGet<CurrentAuthoringResponse>(
            "/api/admin/ai-authoring/sessions/current"
          ),
        ]);
        if (active && isCurrent()) {
          setStatus(authoringStatus);
          setPreview(current.session);
        }
      } catch (caughtError) {
        if (active && isCurrent()) {
          setError(authoringErrorMessage(caughtError));
        }
      } finally {
        if (active && isCurrent()) {
          setLoading(false);
        }
      }
    };
    void load();
    return () => {
      active = false;
      mounted.current = false;
    };
  }, [isCurrent]);

  const clearSourcePdf = () => {
    setSourcePdf(null);
    if (sourcePdfInput.current) {
      sourcePdfInput.current.value = "";
    }
  };

  const selectSourcePdf = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0] ?? null;
    setConsent(false);
    if (file && file.size > maximumSourcePdfBytes) {
      clearSourcePdf();
      setError("PDF must be 10 MiB or smaller.");
      return;
    }
    setSourcePdf(file);
    setError(null);
  };

  const createDocument = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const instruction = prompt.trim() || (sourcePdf ? pdfOnlyInstruction : "");
    if (busy || !status?.enabled || !consent || !instruction || !isCurrent()) {
      return;
    }
    setBusyAction("create");
    setError(null);
    try {
      let result: AuthoringStartResponse;
      if (sourcePdf) {
        const formData = new FormData();
        formData.set("prompt", instruction);
        formData.set("consent", "true");
        formData.set("pdf", sourcePdf);
        result = await apiPostFormData<AuthoringStartResponse>(
          "/api/admin/ai-authoring/sessions",
          formData
        );
      } else {
        result = await apiPost<AuthoringStartResponse>(
          "/api/admin/ai-authoring/sessions",
          { consent: true, prompt }
        );
      }
      if (isCurrent()) {
        setPreview(result.session);
        setPrompt("");
        clearSourcePdf();
      }
    } catch (caughtError) {
      if (isCurrent()) {
        setError(authoringErrorMessage(caughtError));
      }
    } finally {
      if (isCurrent()) {
        setBusyAction(null);
      }
    }
  };

  const reportSessionError = (
    caughtError: unknown,
    requestTargetsSession: boolean
  ) => {
    if (!isCurrent()) {
      return;
    }
    if (
      requestTargetsSession &&
      caughtError instanceof ApiError &&
      caughtError.status === 404
    ) {
      setPreview(null);
      setError("This authoring session expired. Start a new document.");
      return;
    }
    setError(authoringErrorMessage(caughtError));
  };

  const reviseDocument = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (
      !preview ||
      busy ||
      !status?.enabled ||
      !consent ||
      !revisionPrompt.trim() ||
      !isCurrent()
    ) {
      return;
    }
    setBusyAction("revision");
    setError(null);
    try {
      const result = await apiPost<AuthoringStartResponse>(
        `/api/admin/ai-authoring/sessions/${preview.sessionId}/revisions`,
        { consent: true, prompt: revisionPrompt }
      );
      if (isCurrent()) {
        setPreview(result.session);
        setRevisionPrompt("");
      }
    } catch (caughtError) {
      reportSessionError(caughtError, true);
    } finally {
      if (isCurrent()) {
        setBusyAction(null);
      }
    }
  };
  const download = async () => {
    if (!preview || busy || !isCurrent()) {
      return;
    }
    setBusyAction("download");
    setError(null);
    try {
      await downloadArtifact(preview.downloadUrl, docxFilename(preview.title));
    } catch (caughtError) {
      reportSessionError(caughtError, true);
    } finally {
      if (isCurrent()) {
        setBusyAction(null);
      }
    }
  };

  const uploadTemplateDraft = async () => {
    if (!preview || busy || !isCurrent()) {
      return;
    }
    setBusyAction("upload");
    setError(null);
    let downloadedDocx = false;
    try {
      const docx = await apiGetBlob(preview.downloadUrl);
      downloadedDocx = true;
      if (!isCurrent()) {
        return;
      }
      const formData = new FormData();
      formData.set("title", preview.title);
      formData.set("description", preview.description);
      formData.set("source", "upload");
      formData.set(
        "template",
        new File([docx], docxFilename(preview.title), { type: docxContentType })
      );
      const result = await apiPostFormData<CreateTemplateDraftResponse>(
        "/api/admin/forms",
        formData
      );
      if (isCurrent()) {
        await navigate({
          params: { formId: result.form.publicId },
          to: "/admin/forms/$formId",
        });
      }
    } catch (caughtError) {
      reportSessionError(caughtError, !downloadedDocx);
    } finally {
      if (isCurrent()) {
        setBusyAction(null);
      }
    }
  };

  const endSession = async () => {
    if (!preview || busy || !isCurrent()) {
      return;
    }
    setBusyAction("end");
    setError(null);
    try {
      await apiDelete(`/api/admin/ai-authoring/sessions/${preview.sessionId}`);
      if (isCurrent()) {
        setPreview(null);
        clearSourcePdf();
        setPrompt("");
        setRevisionPrompt("");
        setConsent(false);
      }
    } catch (caughtError) {
      if (isCurrent()) {
        if (caughtError instanceof ApiError && caughtError.status === 404) {
          setPreview(null);
          setPrompt("");
          clearSourcePdf();
          setRevisionPrompt("");
          setConsent(false);
        } else {
          setError(authoringErrorMessage(caughtError));
        }
      }
    } finally {
      if (isCurrent()) {
        setBusyAction(null);
      }
    }
  };

  if (loading) {
    return (
      <div
        aria-busy="true"
        className="grid min-h-[40vh] place-items-center gap-3 text-sm text-[var(--ink-soft)]"
      >
        <Spinner />
        <span>Loading your authoring session…</span>
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-3xl">
      <PageHeader
        title="AI Authoring"
        description="Describe a form, refine its DOCX with follow-up instructions, then upload it into a Template Draft."
      />
      {error ? <Notice tone="danger">{error}</Notice> : null}
      {preview ? (
        <ExistingAuthoringSession
          busyAction={busyAction}
          consent={consent}
          download={download}
          endSession={endSession}
          preview={preview}
          revisionPrompt={revisionPrompt}
          reviseDocument={reviseDocument}
          setConsent={setConsent}
          setRevisionPrompt={setRevisionPrompt}
          status={status}
          uploadTemplateDraft={uploadTemplateDraft}
        />
      ) : (
        <CreateAuthoringForm
          busyAction={busyAction}
          consent={consent}
          createDocument={createDocument}
          prompt={prompt}
          selectSourcePdf={selectSourcePdf}
          setConsent={setConsent}
          setPrompt={setPrompt}
          sourcePdf={sourcePdf}
          sourcePdfInput={sourcePdfInput}
          status={status}
        />
      )}
    </div>
  );
};

const AiAuthoringRoute = () => {
  const { user } = useAuth();
  const sessionToken = getToken();
  if (!user || !sessionToken) {
    return null;
  }
  return (
    <AiAuthoringWorkspace
      key={`${user.id}:${sessionToken}`}
      sessionToken={sessionToken}
    />
  );
};

export const Route = createFileRoute("/admin/ai-authoring")({
  component: AiAuthoringRoute,
});
