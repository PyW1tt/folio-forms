import type { AgentSession } from "@earendil-works/pi-coding-agent";

import type { GeneratedField, GeneratedTemplate } from "./document";

export interface AuthoringOwner {
  authSessionId: string;
  userId: string;
}

export interface AuthoringRequestReservation {
  readonly ownerSessionId: string;
}

export interface AuthoringToolState {
  current?: GeneratedTemplate;
  generated?: GeneratedTemplate;
  pending: PendingGeneration;
  validateDocument: (
    document: Uint8Array,
    expected?: Omit<GeneratedTemplate, "document">
  ) => string[];
}

export interface AuthoringSession {
  document: GeneratedTemplate;
  sourcePdf?: Uint8Array;
  lastActivity: Date;
  expiryTimer?: ReturnType<typeof setTimeout>;
  ownerSessionId: string;
  piSession: AgentSession;
  tempDirectory: string;
  pythonAvailable: boolean;
  turns: { prompt: string; assistantMessage: string }[];
  toolState: AuthoringToolState;
  userId: string;
}

export interface PendingGeneration {
  cancelled: boolean;
  readonly inspectionAbort: AbortController;
  done: Promise<void>;
  disposeError?: { cause: unknown };
  finish: () => void;
  piSession?: AgentSession;
  reservation: AuthoringRequestReservation;
  started: boolean;
}

export interface AuthoringPreview {
  assistantMessage: string;
  description: string;
  downloadUrl: string;
  fields: GeneratedField[];
  hasSourcePdf: boolean;
  paragraphs: string[];
  sessionId: string;
  turns: { prompt: string; assistantMessage: string }[];
  title: string;
}
