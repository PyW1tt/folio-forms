import type { KeyboardEvent } from "react";
import { useState } from "react";

import { OnlyOfficeEditor } from "@/components/onlyoffice-editor";
import { Notice } from "@/components/ui";
import type { ReceiptField } from "@/lib/api";

const fieldValue = (
  field: ReceiptField,
  value: unknown,
  picturePresent?: boolean
): string => {
  if (field.type === "picture") {
    if (picturePresent === undefined) {
      return "รูปภาพตรวจสอบไม่ได้";
    }
    return picturePresent ? "มีรูปภาพแนบ" : "ไม่มีรูปภาพ";
  }
  if (value === null || value === undefined || value === "") {
    return "—";
  }
  if (field.type === "checkbox" && typeof value === "boolean") {
    return value ? "ใช่" : "ไม่ใช่";
  }
  if (typeof value === "string") {
    return (
      field.options.find((option) => option.value === value)?.displayText ??
      value
    );
  }
  if (typeof value === "number") {
    return String(value);
  }
  return "มีข้อมูลที่บันทึกแล้ว";
};

export const SubmissionFields = ({
  data,
  fields,
  pictures,
  title = "ข้อมูลคำตอบ",
}: {
  data: Record<string, unknown>;
  fields: ReceiptField[];
  pictures: Record<string, boolean> | null;
  title?: string;
}) => (
  <>
    <h2 className="mb-4 font-semibold">{title}</h2>
    <dl className="space-y-3">
      {fields.map((field) => (
        <div
          className="rounded-[10px] border border-[var(--line)] p-4"
          key={field.tag}
        >
          <dt className="font-semibold">{field.label}</dt>
          <dd className="mt-2 whitespace-pre-wrap break-words text-[var(--ink-soft)]">
            {fieldValue(field, data[field.tag], pictures?.[field.tag])}
          </dd>
        </div>
      ))}
    </dl>
  </>
);
export const SubmissionResultViewer = ({
  configUrl,
  data,
  documentAvailable,
  fields,
  pictures,
}: {
  configUrl?: string;
  data: Record<string, unknown>;
  documentAvailable: boolean;
  fields: ReceiptField[];
  pictures: Record<string, boolean> | null;
}) => {
  const [activeTab, setActiveTab] = useState<"document" | "fields">(
    "document"
  );
  const activateAdjacentTab = (event: KeyboardEvent<HTMLButtonElement>) => {
    if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") {
      return;
    }
    event.preventDefault();
    const nextTab = activeTab === "document" ? "fields" : "document";
    setActiveTab(nextTab);
    document.getElementById(`submission-${nextTab}-tab`)?.focus();
  };

  return (
    <div>
      <div
        aria-label="Result view"
        className="flex flex-wrap gap-2 border-b border-[var(--line)] pb-4"
        role="tablist"
      >
        <button
          aria-controls="submission-document-panel"
          aria-selected={activeTab === "document"}
          className={`min-h-10 rounded-[10px] border border-[var(--line-strong)] px-3 text-sm font-semibold ${
            activeTab === "document"
              ? "bg-[var(--ink)] text-[var(--paper)]"
              : "bg-[var(--paper)] text-[var(--ink-soft)]"
          }`}
          id="submission-document-tab"
          onClick={() => setActiveTab("document")}
          onKeyDown={activateAdjacentTab}
          role="tab"
          type="button"
        >
          Document
        </button>
        <button
          aria-controls="submission-fields-panel"
          aria-selected={activeTab === "fields"}
          className={`min-h-10 rounded-[10px] border border-[var(--line-strong)] px-3 text-sm font-semibold ${
            activeTab === "fields"
              ? "bg-[var(--ink)] text-[var(--paper)]"
              : "bg-[var(--paper)] text-[var(--ink-soft)]"
          }`}
          id="submission-fields-tab"
          onClick={() => setActiveTab("fields")}
          onKeyDown={activateAdjacentTab}
          role="tab"
          type="button"
        >
          Fields
        </button>
      </div>
      <div
        aria-labelledby="submission-document-tab"
        className="pt-4"
        hidden={activeTab !== "document"}
        id="submission-document-panel"
        role="tabpanel"
        tabIndex={0}
      >
        {activeTab === "document" && documentAvailable && configUrl ? (
          <OnlyOfficeEditor
            configUrl={configUrl}
            readOnly
            title="Read-only submission document"
          />
        ) : null}
        {activeTab === "document" && (!documentAvailable || !configUrl) ? (
          <Notice>เอกสารสำหรับฉบับนี้ไม่พร้อมใช้งาน</Notice>
        ) : null}
      </div>
      <div
        aria-labelledby="submission-fields-tab"
        className="pt-4"
        hidden={activeTab !== "fields"}
        id="submission-fields-panel"
        role="tabpanel"
        tabIndex={0}
      >
        {activeTab === "fields" ? (
          <SubmissionFields data={data} fields={fields} pictures={pictures} />
        ) : null}
      </div>
    </div>
  );
};
