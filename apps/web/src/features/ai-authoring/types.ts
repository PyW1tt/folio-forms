export interface AiAuthoringStatus {
  disclosure: string;
  enabled: boolean;
}
export interface GeneratedField {
  label: string;
  placeholder: string;
  tag: string;
}
export interface AuthoringPreview {
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

export type BusyAction =
  | "create"
  | "revision"
  | "download"
  | "upload"
  | "end"
  | null;
