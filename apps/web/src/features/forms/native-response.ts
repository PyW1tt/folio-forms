import type { NativeFormField } from "@/components/native-form";
import { apiPostFormData, downloadArtifact, waitForOperation } from "@/lib/api";
import type { Operation } from "@/lib/api";

export interface NativeEditorConfig {
  capabilities: Record<"save-draft" | "submit", string>;
  data: Record<string, unknown>;
  documentKey: string;
  fields: NativeFormField[];
  fillMethod: "native";
  lockedFields: Record<string, boolean>;
  pictures: Record<string, boolean>;
  responseId: string;
}

export type ExportFormat = "docx" | "pdf";

export const nativeValuesFromConfig = (
  config: NativeEditorConfig
): Record<string, unknown> => {
  const values: Record<string, unknown> = {};
  for (const field of config.fields) {
    const value = config.data[field.tag];
    switch (field.type) {
      case "checkbox": {
        values[field.tag] = value === true;
        break;
      }
      case "combo":
      case "dropdown": {
        values[field.tag] = typeof value === "string" ? value : null;
        break;
      }
      case "picture": {
        values[field.tag] = "";
        break;
      }
      default: {
        values[field.tag] = typeof value === "string" ? value : "";
      }
    }
  }
  return values;
};

export const draftSaveSuccessMessage = (
  exportFormat: ExportFormat | null | undefined
): string => {
  if (exportFormat === "docx") {
    return "บันทึกและดาวน์โหลด DOCX แล้ว";
  }
  if (exportFormat === "pdf") {
    return "บันทึกและดาวน์โหลด PDF แล้ว";
  }
  return "บันทึกฉบับร่างคำตอบแล้ว";
};

export const nativeSaveErrorMessage = (
  action: "save-draft" | "submit",
  saved: boolean,
  exportFormat: ExportFormat | null | undefined
): string => {
  if (action === "submit") {
    return "ส่งแบบฟอร์มไม่สำเร็จ กรุณาลองใหม่";
  }
  if (!saved) {
    return "บันทึกฉบับร่างไม่สำเร็จ กรุณาลองใหม่";
  }
  if (exportFormat) {
    return "บันทึกแล้ว แต่ดาวน์โหลดไฟล์ไม่สำเร็จ กรุณาลองใหม่";
  }
  return "บันทึกแล้ว แต่โหลดคำตอบล่าสุดไม่สำเร็จ กรุณาโหลดแบบฟอร์มใหม่";
};

export const downloadDraftArtifact = (
  responseId: string,
  format: ExportFormat
): Promise<void> =>
  downloadArtifact(
    `/api/responses/${responseId}/draft/${format}`,
    `response-${responseId}.${format}`
  );

export const submitNativeResponse = async ({
  publicId,
  action,
  config,
  values,
  pictureFiles,
  onOperation,
}: {
  publicId: string;
  action: "save-draft" | "submit";
  config: NativeEditorConfig;
  values: Record<string, unknown>;
  pictureFiles: Record<string, File>;
  onOperation: (operation: Operation) => void;
}): Promise<Operation> => {
  const payload = {
    data: Object.fromEntries(
      config.fields
        .filter((field) => field.type !== "picture")
        .map((field): [string, unknown] => {
          const value = values[field.tag];
          return [
            field.tag,
            field.type === "date" && value === "" ? null : value,
          ];
        })
    ),
    documentKey: config.documentKey,
    fillMethod: "native",
    responseId: config.responseId,
  };
  const formData = new FormData();
  formData.set("payload", JSON.stringify(payload));
  for (const [tag, file] of Object.entries(pictureFiles)) {
    formData.set(`picture:${tag}`, file);
  }
  const endpoint = action === "save-draft" ? "draft" : "submit";
  const { operationId } = await apiPostFormData<{ operationId: string }>(
    `/api/forms/${publicId}/${endpoint}`,
    formData,
    config.capabilities[action]
  );
  onOperation({ id: operationId, status: "pending" });
  const completed = await waitForOperation(operationId, onOperation);
  onOperation(completed);
  return completed;
};
