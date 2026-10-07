import { Save } from "lucide-react";
import type { FormEvent } from "react";

import { Button, Input, Spinner, Textarea } from "@/components/ui";
import type { FillMethod, FormSummary } from "@/lib/api";

interface FormMetadataPanelProps {
  title: string;
  description: string;
  metadataBusy: "archive" | "fill-method" | "save" | "duplicate" | null;
  metadataCanAct: boolean;
  fillMethod: FillMethod;
  status: FormSummary["status"];
  nativeFillAvailable: boolean;
  fillMethodHelp: string;
  setTitle: (value: string) => void;
  setDescription: (value: string) => void;
  updateMetadata: (event: FormEvent<HTMLFormElement>) => Promise<void>;
  updateFillMethod: (value: FillMethod) => Promise<void>;
}

export const FormMetadataPanel = ({
  title,
  description,
  metadataBusy,
  metadataCanAct,
  fillMethod,
  status,
  nativeFillAvailable,
  fillMethodHelp,
  setTitle,
  setDescription,
  updateMetadata,
  updateFillMethod,
}: FormMetadataPanelProps) => (
  <section
    className="mb-4 rounded-[var(--radius)] border border-[var(--line)] bg-[var(--paper)] p-5"
    aria-labelledby="metadata-title"
  >
    <div className="max-w-2xl">
      <h2 id="metadata-title" className="text-lg font-semibold">
        ข้อมูลแบบฟอร์ม
      </h2>
      <p className="mt-1 text-sm text-[var(--ink-soft)]">
        แก้ไขชื่อและคำอธิบายได้ทั้งแบบร่างและแบบฟอร์มที่เผยแพร่แล้ว
      </p>
      <form
        className="mt-4 space-y-4"
        onSubmit={updateMetadata}
        aria-busy={metadataBusy === "save"}
      >
        <div className="space-y-2">
          <label className="text-sm font-semibold" htmlFor="form-title">
            ชื่อแบบฟอร์ม
          </label>
          <Input
            id="form-title"
            name="title"
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            maxLength={200}
            required
            aria-describedby="form-title-help"
            disabled={!metadataCanAct}
          />
          <p id="form-title-help" className="text-xs text-[var(--ink-soft)]">
            ต้องระบุชื่อ ความยาวไม่เกิน 200 ตัวอักษร
          </p>
        </div>
        <div className="space-y-2">
          <label className="text-sm font-semibold" htmlFor="form-description">
            คำอธิบาย
          </label>
          <Textarea
            id="form-description"
            name="description"
            value={description}
            onChange={(event) => setDescription(event.target.value)}
            maxLength={2000}
            rows={4}
            aria-describedby="form-description-help"
            disabled={!metadataCanAct}
          />
          <p
            id="form-description-help"
            className="text-xs text-[var(--ink-soft)]"
          >
            ใส่คำอธิบายเพิ่มเติมได้ไม่เกิน 2,000 ตัวอักษร
          </p>
        </div>
        <div className="flex justify-end">
          <Button type="submit" disabled={!metadataCanAct}>
            {metadataBusy === "save" ? <Spinner /> : <Save size={15} />}
            {metadataBusy === "save" ? "กำลังบันทึกข้อมูล…" : "บันทึกข้อมูลแบบฟอร์ม"}
          </Button>
        </div>
      </form>
      <div
        className="mt-5 space-y-2"
        aria-busy={metadataBusy === "fill-method"}
      >
        <label className="text-sm font-semibold" htmlFor="form-fill-method">
          Fill Method
        </label>
        <select
          className="min-h-11 w-full rounded-[10px] border border-[var(--line-strong)] bg-[var(--paper)] px-3 text-[var(--ink)] shadow-sm focus:border-[var(--ink)] focus:outline-none"
          disabled={!metadataCanAct || status === "draft"}
          id="form-fill-method"
          onChange={(event) => {
            const { value } = event.target;
            if (value === "native" || value === "onlyoffice") {
              void updateFillMethod(value);
            }
          }}
          value={fillMethod}
        >
          <option value="onlyoffice">ONLYOFFICE</option>
          <option disabled={!nativeFillAvailable} value="native">
            Native form
          </option>
        </select>
        <p
          className="text-xs text-[var(--ink-soft)]"
          id="form-fill-method-help"
        >
          {fillMethodHelp}
        </p>
        {metadataBusy === "fill-method" ? (
          <p className="inline-flex items-center gap-2 text-xs" role="status">
            <Spinner />
            Updating Fill Method…
          </p>
        ) : null}
      </div>
    </div>
  </section>
);
