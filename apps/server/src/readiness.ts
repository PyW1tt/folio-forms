import { prisma } from "@onlyoffice/db";
import { env } from "@onlyoffice/env/server";

import { findTemplateSource } from "./template-source";

export const readinessStatus = async (): Promise<boolean> => {
  try {
    const rustfsReady = await fetch(
      new URL("/health/ready", env.RUSTFS_ENDPOINT),
      { signal: AbortSignal.timeout(2000) }
    );
    const [templateSource] = await Promise.all([
      findTemplateSource(),
      prisma.$queryRaw`SELECT 1`,
    ]);
    return rustfsReady.ok && templateSource !== null;
  } catch {
    return false;
  }
};
