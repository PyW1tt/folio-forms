import { createFileRoute, useNavigate } from "@tanstack/react-router";
import type { ChangeEvent, FormEvent } from "react";
import { useCallback, useEffect, useRef, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import { Notice, Spinner } from "@/components/ui";
import {
  CreateAuthoringForm,
  ExistingAuthoringSession,
} from "@/features/ai-authoring/components";
import type {
  AiAuthoringStatus,
  AuthoringPreview,
  BusyAction,
} from "@/features/ai-authoring/types";
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
interface AuthoringStartResponse {
  session: AuthoringPreview;
}
interface CurrentAuthoringResponse {
  session: AuthoringPreview | null;
}
interface CreateTemplateDraftResponse {
  form: FormSummary;
}

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
