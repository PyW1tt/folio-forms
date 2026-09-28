import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { Download, FileCheck2, WandSparkles } from "lucide-react";
import type { FormEvent } from "react";
import { useEffect, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import { Button, Card, Input, Notice, Spinner, Textarea } from "@/components/ui";
import {
  ApiError,
  apiDelete,
  apiGet,
  apiGetBlob,
  apiPost,
  apiPostFormData,
  downloadArtifact,
} from "@/lib/api";
import type { FormSummary } from "@/lib/api";

const docxContentType =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
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
  paragraphs: string[];
  sessionId: string;
  title: string;
}
interface AuthoringStartResponse {
  session: AuthoringPreview;
}
interface CreateTemplateDraftResponse {
  form: FormSummary;
}

const authoringErrorMessage = (error: unknown): string => {
  if (!(error instanceof ApiError)) {
    return "AI Authoring is unavailable. Try again later.";
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

const AiAuthoringRoute = () => {
  const navigate = useNavigate();
  const [status, setStatus] = useState<AiAuthoringStatus | null>(null);
  const [prompt, setPrompt] = useState("");
  const [consent, setConsent] = useState(false);
  const [preview, setPreview] = useState<AuthoringPreview | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let active = true;
    apiGet<AiAuthoringStatus>("/api/admin/ai-authoring")
      .then((result) => {
        if (active) {
          setStatus(result);
        }
      })
      .catch((caughtError: unknown) => {
        if (active) {
          setError(authoringErrorMessage(caughtError));
        }
      })
      .finally(() => {
        if (active) {
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, []);

  const createDocument = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (busy || !status?.enabled || !consent || !prompt.trim()) {
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const result = await apiPost<AuthoringStartResponse>(
        "/api/admin/ai-authoring/sessions",
        { consent: true, prompt }
      );
      setPreview(result.session);
    } catch (caughtError) {
      setError(authoringErrorMessage(caughtError));
    } finally {
      setBusy(false);
    }
  };

  const download = async () => {
    if (!preview || busy) {
      return;
    }
    setBusy(true);
    setError(null);
    try {
      await downloadArtifact(preview.downloadUrl, docxFilename(preview.title));
    } catch (caughtError) {
      setError(authoringErrorMessage(caughtError));
    } finally {
      setBusy(false);
    }
  };

  const uploadTemplateDraft = async () => {
    if (!preview || busy) {
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const docx = await apiGetBlob(preview.downloadUrl);
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
      await navigate({
        params: { formId: result.form.publicId },
        to: "/admin/forms/$formId",
      });
    } catch (caughtError) {
      setError(authoringErrorMessage(caughtError));
      setBusy(false);
    }
  };

  const endSession = async () => {
    if (!preview || busy) {
      return;
    }
    setBusy(true);
    setError(null);
    try {
      await apiDelete(
        `/api/admin/ai-authoring/sessions/${preview.sessionId}`
      );
      setPreview(null);
      setPrompt("");
      setConsent(false);
    } catch (caughtError) {
      setError(authoringErrorMessage(caughtError));
    } finally {
      setBusy(false);
    }
  };

  if (loading) {
    return (
      <div className="grid min-h-[40vh] place-items-center">
        <Spinner />
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-3xl">
      <PageHeader
        title="AI Authoring"
        description="Describe a form in text. Preview the generated DOCX, then upload it into a normal Template Draft."
      />
      {error ? <Notice tone="danger">{error}</Notice> : null}
      {preview ? (
        <div className="space-y-5">
          <Notice tone="success">
            {preview.assistantMessage}
          </Notice>
          <Card className="space-y-5 p-5 sm:p-7" aria-label="Read-only DOCX preview">
            <div>
              <h2 className="text-2xl font-bold">{preview.title}</h2>
              {preview.description ? (
                <p className="mt-2 text-[var(--ink-soft)]">
                  {preview.description}
                </p>
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
            <Button disabled={busy} onClick={download} type="button" variant="secondary">
              <Download size={16} />
              Download DOCX
            </Button>
            <Button disabled={busy} onClick={uploadTemplateDraft} type="button">
              {busy ? <Spinner /> : <FileCheck2 size={16} />}
              Upload as Template Draft
            </Button>
            <Button
              disabled={busy}
              onClick={endSession}
              type="button"
              variant="ghost"
            >
              End session and delete Folio copy
            </Button>
          </div>
        </div>
      ) : (
        <Card className="space-y-5 p-5 sm:p-7">
          {!status?.enabled ? (
            <Notice>
              AI Authoring is disabled. The server needs an OmniRoute endpoint,
              permitted service credential, and configured model alias.
            </Notice>
          ) : null}
          <form className="space-y-5" onSubmit={createDocument}>
            <label className="block space-y-2">
              <span className="font-semibold">Describe your form</span>
              <Textarea
                disabled={!status?.enabled || busy}
                maxLength={16_000}
                onChange={(event) => setPrompt(event.target.value)}
                placeholder="For example: Create an employee equipment request form with name, department, requested items, and approval notes."
                required
                rows={7}
                value={prompt}
              />
            </label>
            <Notice>
              {status?.disclosure ??
                "Your prompt and document content go to OmniRoute and its configured provider. Folio Forms cannot promise upstream deletion."}
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
                I consent to sending this prompt and generated document content
                to OmniRoute and its configured provider.
              </span>
            </label>
            <Button
              disabled={!status?.enabled || busy || !consent || !prompt.trim()}
              type="submit"
            >
              {busy ? <Spinner /> : <WandSparkles size={16} />}
              Generate DOCX
            </Button>
          </form>
        </Card>
      )}
    </div>
  );
};

export const Route = createFileRoute("/admin/ai-authoring")({
  component: AiAuthoringRoute,
});
