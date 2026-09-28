import type { FormEvent } from "react";

import { Button, Card, Textarea } from "@/components/ui";

export interface NativeTextField {
  tag: string;
  label: string;
  placeholder: string | null;
  position: number;
  required: boolean;
  type: "text";
}

export interface NativeTextFormProps {
  fields: NativeTextField[];
  lockedFields: Record<string, boolean>;
  operationBusy: boolean;
  values: Record<string, string>;
  onChange: (tag: string, value: string) => void;
  onExportDocx: () => void;
  onExportPdf: () => void;
  onSave: () => void;
  onSubmit: () => void;
}

export const NativeTextForm = ({
  fields,
  lockedFields,
  operationBusy,
  values,
  onChange,
  onExportDocx,
  onExportPdf,
  onSave,
  onSubmit,
}: NativeTextFormProps) => {
  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    onSubmit();
  };

  return (
    <Card className="space-y-5 p-5 sm:p-6">
      <div>
        <h2 className="text-lg font-semibold text-[var(--ink)]">
          กรอกข้อมูลในแบบฟอร์ม
        </h2>
        <p className="mt-1 text-sm text-[var(--ink-soft)]">
          ข้อมูลที่มีเครื่องหมาย * จำเป็นต้องกรอก
        </p>
      </div>
      <form className="space-y-5" onSubmit={handleSubmit}>
        {fields.map((field, index) => {
          const fieldId = `native-field-${index}`;
          const helpId = `${fieldId}-help`;
          const locked = lockedFields[field.tag] === true;
          const helpText = [
            field.placeholder ? `คำแนะนำ: ${field.placeholder}` : "",
            locked ? "ค่านี้มาจาก Prefill และแก้ไขไม่ได้" : "",
          ]
            .filter(Boolean)
            .join(" ");
          return (
            <div className="space-y-2" key={field.tag}>
              <label
                className="block text-sm font-semibold text-[var(--ink)]"
                htmlFor={fieldId}
              >
                {field.label}
                {field.required ? (
                  <span aria-hidden="true" className="text-[var(--danger)]">
                    {" "}*
                  </span>
                ) : null}
              </label>
              <Textarea
                aria-describedby={field.placeholder || locked ? helpId : undefined}
                className={locked ? "bg-[var(--muted)]" : undefined}
                id={fieldId}
                maxLength={10_000}
                onChange={(event) => onChange(field.tag, event.target.value)}
                placeholder={field.placeholder ?? undefined}
                readOnly={locked}
                required={field.required}
                rows={3}
                value={values[field.tag] ?? ""}
              />
              {helpText ? (
                <p className="text-xs text-[var(--ink-soft)]" id={helpId}>
                  {helpText}
                </p>
              ) : null}
            </div>
          );
        })}
        <div className="flex flex-wrap gap-2 border-t border-[var(--line)] pt-4">
          <Button disabled={operationBusy} onClick={onSave} type="button" variant="secondary">
            {operationBusy ? "กำลังดำเนินการ..." : "บันทึกฉบับร่าง"}
          </Button>
          <Button disabled={operationBusy} type="submit">
            ส่งแบบฟอร์ม
          </Button>
          <Button
            disabled={operationBusy}
            onClick={onExportDocx}
            type="button"
            variant="secondary"
          >
            ดาวน์โหลด DOCX
          </Button>
          <Button
            disabled={operationBusy}
            onClick={onExportPdf}
            type="button"
            variant="secondary"
          >
            ดาวน์โหลด PDF
          </Button>
        </div>
      </form>
    </Card>
  );
};
