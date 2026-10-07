import type { AiAuthoringSessions } from "../ai-authoring";
import type { LegacySsoConfig } from "../legacy-sso/config";
import type { OnlyOfficeClient } from "../onlyoffice";
import type { putObject } from "../storage";

export interface RouteDependencies {
  onlyOffice: OnlyOfficeClient;
  removeObject: (key: string) => Promise<void>;
  storeObject: typeof putObject;
  allowedCallbackOrigins: ReadonlySet<string>;
  callbackMaximumBytes: number;
  prefillHandoffSecret: string;
  prefillReturnUrl: string;
  handoffClock: () => Date;
  legacySso: LegacySsoConfig | null;
  aiAuthoring: AiAuthoringSessions;
  requestIp?: (request: Request) => string | null | undefined;
}
