import type { FormEvent } from "react";

import { Button, Card, Input, Textarea } from "@/components/ui";

export interface NativeFormField {
  tag: string;
  label: string;
  placeholder: string | null;
  options: { displayText: string; value: string }[];
  position: number;
  required: boolean;
  type: "text" | "checkbox" | "date" | "dropdown" | "combo" | "picture";
  pictureMaxBytes: number | null;
  pictureMaxWidth: number | null;
  pictureMaxHeight: number | null;
}

export interface NativeFormProps {
  fields: NativeFormField[];
  lockedFields: Record<string, boolean>;
  operationBusy: boolean;
  values: Record<string, unknown>;
  pictures: Record<string, boolean>;
  pictureFiles: Record<string, File>;
  pictureInputKey: number;
  onChange: (tag: string, value: string | boolean | null) => void;
  onPictureChange: (tag: string, file: File | null) => void;
  onExportDocx: () => void;
  onExportPdf: () => void;
  onSave: () => void;
  onSubmit: () => void;
}

export const comboDisplayValue = (
  field: NativeFormField,
  value: unknown
): string => {
  if (typeof value !== "string") {
    return "";
  }
  return (
    field.options.find((option) => option.value === value)?.displayText ?? value
  );
};

export const comboStoredValue = (
  field: NativeFormField,
  input: string
): string | null => {
  if (input === "") {
    return null;
  }
  return (
    field.options.find((option) => option.displayText === input)?.value ?? input
  );
};

interface NativeFieldControlProps {
  describedBy?: string;
  field: NativeFormField;
  fieldId: string;
  locked: boolean;
  onChange: NativeFormProps["onChange"];
  onPictureChange: NativeFormProps["onPictureChange"];
  operationBusy: boolean;
  pictureFiles: Record<string, File>;
  pictureInputKey: number;
  pictures: Record<string, boolean>;
  value: unknown;
}
const NativeDropdownControl = ({
  describedBy,
  field,
  fieldId,
  locked,
  onChange,
  operationBusy,
  value,
}: Pick<
  NativeFieldControlProps,
  | "describedBy"
  | "field"
  | "fieldId"
  | "locked"
  | "onChange"
  | "operationBusy"
  | "value"
>) => {
  const selectedOptionIndex =
    typeof value === "string"
      ? field.options.findIndex((option) => option.value === value)
      : -1;
  return (
    <select
      aria-describedby={describedBy}
      className="min-h-10 w-full rounded-md border border-[var(--line)] bg-[var(--paper)] px-3 text-sm"
      disabled={operationBusy || locked}
      id={fieldId}
      onChange={(event) => {
        const optionIndex = Number(event.target.value) - 1;
        onChange(
          field.tag,
          optionIndex < 0 ? null : (field.options[optionIndex]?.value ?? null)
        );
      }}
      required={field.required}
      value={selectedOptionIndex < 0 ? "" : String(selectedOptionIndex + 1)}
    >
      <option value="">{field.placeholder || "เลือกตัวเลือก"}</option>
      {field.options.map((option, index) => (
        <option key={`${option.value}-${index}`} value={String(index + 1)}>
          {option.displayText}
        </option>
      ))}
    </select>
  );
};

const NativePictureControl = ({
  describedBy,
  field,
  fieldId,
  onPictureChange,
  operationBusy,
  pictureFiles,
  pictureInputKey,
  pictures,
}: Pick<
  NativeFieldControlProps,
  | "describedBy"
  | "field"
  | "fieldId"
  | "onPictureChange"
  | "operationBusy"
  | "pictureFiles"
  | "pictureInputKey"
  | "pictures"
>) => (
  <input
    accept="image/jpeg,image/png"
    aria-describedby={describedBy}
    disabled={operationBusy}
    id={fieldId}
    key={`${fieldId}-${pictureInputKey}`}
    onChange={(event) =>
      onPictureChange(field.tag, event.target.files?.[0] ?? null)
    }
    required={
      field.required && !pictures[field.tag] && !pictureFiles[field.tag]
    }
    type="file"
  />
);

const NativeFieldControl = ({
  describedBy,
  field,
  fieldId,
  locked,
  onChange,
  onPictureChange,
  operationBusy,
  pictureFiles,
  pictureInputKey,
  pictures,
  value,
}: NativeFieldControlProps) => {
  switch (field.type) {
    case "checkbox": {
      return (
        <input
          aria-describedby={describedBy}
          checked={value === true}
          disabled={operationBusy || locked}
          id={fieldId}
          onChange={(event) => onChange(field.tag, event.target.checked)}
          required={field.required}
          type="checkbox"
        />
      );
    }
    case "combo": {
      return (
        <>
          <Input
            aria-describedby={describedBy}
            className={locked ? "bg-[var(--muted)]" : undefined}
            disabled={operationBusy}
            id={fieldId}
            list={`${fieldId}-options`}
            maxLength={10_000}
            onChange={(event) =>
              onChange(field.tag, comboStoredValue(field, event.target.value))
            }
            placeholder={field.placeholder ?? undefined}
            readOnly={locked}
            required={field.required}
            value={comboDisplayValue(field, value)}
          />
          <datalist id={`${fieldId}-options`}>
            {field.options.map((option) => (
              <option key={option.value} value={option.displayText}>
                {option.displayText}
              </option>
            ))}
          </datalist>
        </>
      );
    }
    case "date": {
      return (
        <Input
          aria-describedby={describedBy}
          className={locked ? "bg-[var(--muted)]" : undefined}
          disabled={operationBusy || locked}
          id={fieldId}
          onChange={(event) => onChange(field.tag, event.target.value)}
          required={field.required}
          type="date"
          value={typeof value === "string" ? value : ""}
        />
      );
    }
    case "dropdown": {
      return (
        <NativeDropdownControl
          describedBy={describedBy}
          field={field}
          fieldId={fieldId}
          locked={locked}
          onChange={onChange}
          operationBusy={operationBusy}
          value={value}
        />
      );
    }
    case "picture": {
      return (
        <NativePictureControl
          describedBy={describedBy}
          field={field}
          fieldId={fieldId}
          onPictureChange={onPictureChange}
          operationBusy={operationBusy}
          pictureFiles={pictureFiles}
          pictureInputKey={pictureInputKey}
          pictures={pictures}
        />
      );
    }
    case "text": {
      return (
        <Textarea
          aria-describedby={describedBy}
          className={locked ? "bg-[var(--muted)]" : undefined}
          disabled={operationBusy}
          id={fieldId}
          maxLength={10_000}
          onChange={(event) => onChange(field.tag, event.target.value)}
          placeholder={field.placeholder ?? undefined}
          readOnly={locked}
          required={field.required}
          rows={3}
          value={typeof value === "string" ? value : ""}
        />
      );
    }
    default: {
      throw new TypeError("Unsupported native field type");
    }
  }
};

interface NativeFieldProps extends Omit<
  NativeFieldControlProps,
  "describedBy"
> {
  label: string;
  placeholder: string | null;
  pictureMaxBytes: number | null;
  pictureMaxHeight: number | null;
  pictureMaxWidth: number | null;
  required: boolean;
}

const NativeField = ({
  field,
  fieldId,
  label,
  locked,
  onChange,
  onPictureChange,
  operationBusy,
  pictureFiles,
  pictureInputKey,
  pictures,
  placeholder,
  pictureMaxBytes,
  pictureMaxHeight,
  pictureMaxWidth,
  required,
  value,
}: NativeFieldProps) => {
  const helpId = `${fieldId}-help`;
  const pictureHelpId = `${fieldId}-picture-help`;
  const helpText = [
    placeholder ? `คำแนะนำ: ${placeholder}` : "",
    locked ? "ค่านี้มาจาก Prefill และแก้ไขไม่ได้" : "",
  ]
    .filter(Boolean)
    .join(" ");
  let inputDescription: string | undefined;
  if (field.type === "picture") {
    inputDescription = `${pictureHelpId}${helpText ? ` ${helpId}` : ""}`;
  } else if (helpText) {
    inputDescription = helpId;
  }
  return (
    <div className="space-y-2">
      {field.type === "checkbox" ? (
        <label
          className="flex items-center gap-2 text-sm font-semibold text-[var(--ink)]"
          htmlFor={fieldId}
        >
          <NativeFieldControl
            describedBy={inputDescription}
            field={field}
            fieldId={fieldId}
            locked={locked}
            onChange={onChange}
            onPictureChange={onPictureChange}
            operationBusy={operationBusy}
            pictureFiles={pictureFiles}
            pictureInputKey={pictureInputKey}
            pictures={pictures}
            value={value}
          />
          {label}
          {required ? (
            <span aria-hidden="true" className="text-[var(--danger)]">
              {" "}
              *
            </span>
          ) : null}
        </label>
      ) : (
        <>
          <label
            className="block text-sm font-semibold text-[var(--ink)]"
            htmlFor={fieldId}
          >
            {label}
            {required ? (
              <span aria-hidden="true" className="text-[var(--danger)]">
                {" "}
                *
              </span>
            ) : null}
          </label>
          <NativeFieldControl
            describedBy={inputDescription}
            field={field}
            fieldId={fieldId}
            locked={locked}
            onChange={onChange}
            onPictureChange={onPictureChange}
            operationBusy={operationBusy}
            pictureFiles={pictureFiles}
            pictureInputKey={pictureInputKey}
            pictures={pictures}
            value={value}
          />
        </>
      )}
      {field.type === "picture" ? (
        <p className="text-xs text-[var(--ink-soft)]" id={pictureHelpId}>
          {pictures[field.tag] ? "มีรูปภาพที่บันทึกไว้แล้ว" : "ยังไม่มีรูปภาพที่บันทึกไว้"}
          {pictureMaxBytes === null ? "" : ` ขนาดไม่เกิน ${pictureMaxBytes} ไบต์`}
          {pictureMaxWidth === null || pictureMaxHeight === null
            ? ""
            : ` ขนาดภาพไม่เกิน ${pictureMaxWidth} × ${pictureMaxHeight} พิกเซล`}
        </p>
      ) : null}
      {helpText ? (
        <p className="text-xs text-[var(--ink-soft)]" id={helpId}>
          {helpText}
        </p>
      ) : null}
    </div>
  );
};

export const NativeForm = ({
  fields,
  lockedFields,
  operationBusy,
  values,
  pictures,
  pictureFiles,
  pictureInputKey,
  onChange,
  onPictureChange,
  onExportDocx,
  onExportPdf,
  onSave,
  onSubmit,
}: NativeFormProps) => {
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
        {fields.map((field) => (
          <NativeField
            field={field}
            fieldId={`native-field-${field.position}`}
            key={field.tag}
            label={field.label}
            locked={lockedFields[field.tag] === true}
            onChange={onChange}
            onPictureChange={onPictureChange}
            operationBusy={operationBusy}
            pictureFiles={pictureFiles}
            pictureInputKey={pictureInputKey}
            pictureMaxBytes={field.pictureMaxBytes}
            pictureMaxHeight={field.pictureMaxHeight}
            pictureMaxWidth={field.pictureMaxWidth}
            pictures={pictures}
            placeholder={field.placeholder}
            required={field.required}
            value={values[field.tag]}
          />
        ))}
        <div className="flex flex-wrap gap-2 border-t border-[var(--line)] pt-4">
          <Button
            disabled={operationBusy}
            onClick={onSave}
            type="button"
            variant="secondary"
          >
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
