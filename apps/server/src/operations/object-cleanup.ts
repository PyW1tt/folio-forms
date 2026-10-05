import { prisma, OperationStatus } from "@onlyoffice/db";

import type { Operation } from "../model-types";
// oxlint-disable func-style prefer-destructuring no-await-in-loop no-use-before-define no-nested-ternary complexity no-shadow -- Route modules keep declaration order and sequential persistence invariants.
import { deleteObject, objectExists } from "../storage";
import { operationMetadata } from "./model";

export const objectCleanupIntentGraceMs = 15 * 60_000;

export function uniqueObjectKeys(
  keys: readonly (string | null | undefined)[]
): string[] {
  return [
    ...new Set(
      keys.filter(
        (key): key is string => typeof key === "string" && key.length > 0
      )
    ),
  ];
}

export async function deleteObjects(
  keys: readonly (string | null | undefined)[]
): Promise<void> {
  const uniqueKeys = uniqueObjectKeys(keys);
  await Promise.all(
    uniqueKeys.map(async (key) => {
      try {
        await deleteObject(key);
      } catch (error) {
        console.error(`Could not remove object ${key}`, error);
      }
    })
  );
}

export async function drainObjectCleanupIntents(
  objectKeys?: readonly string[],
  deletionResponseLookupDigest?: string,
  removeObject: (key: string) => Promise<void> = deleteObject
): Promise<void> {
  const cleanupAfter = new Date();
  const intents = await prisma.objectCleanupIntent.findMany({
    orderBy: { createdAt: "asc" },
    where: {
      cleanupAfter: { lte: cleanupAfter },
      ...(deletionResponseLookupDigest ? { deletionResponseLookupDigest } : {}),
      ...(objectKeys ? { objectKey: { in: [...objectKeys] } } : {}),
    },
  });
  for (const intent of intents) {
    if (!(await deleteObjectUnlessCanonical(intent.objectKey, removeObject))) {
      continue;
    }
    await prisma.objectCleanupIntent.deleteMany({
      where: { id: intent.id, objectKey: intent.objectKey },
    });
  }
}

export async function deleteObjectUnlessCanonical(
  key: string,
  removeObject: (key: string) => Promise<void> = deleteObject
): Promise<boolean> {
  try {
    const references = await Promise.all([
      prisma.templateDraft.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.publishedTemplate.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.response.findFirst({
        select: { id: true },
        where: { draftObjectKey: key },
      }),
      prisma.submission.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.correction.findFirst({
        select: { id: true },
        where: { objectKey: key },
      }),
      prisma.editorLease.findFirst({
        select: { id: true },
        where: { workspaceObjectKey: key },
      }),
    ]);
    if (references.some((reference) => reference !== null)) {
      return false;
    }
    await removeObject(key);
    return !(await objectExists(key));
  } catch (error) {
    console.error(`Could not verify whether object ${key} is canonical`, error);
    return false;
  }
}

export async function cleanupTerminalOperationObjects(
  operation: Operation
): Promise<void> {
  if (
    operation.status !== OperationStatus.completed &&
    operation.status !== OperationStatus.failed
  ) {
    return;
  }
  const metadata = operationMetadata(operation.metadata);
  if (operation.status === OperationStatus.completed) {
    await deleteObjects(
      metadata.cleanupObjectKeys ?? [metadata.stagedObjectKey]
    );
    return;
  }
  await deleteObjects([metadata.stagedObjectKey]);
  await deleteObjectUnlessCanonical(metadata.finalObjectKey);
  if (metadata.action === "save-correction" && metadata.workspaceDocumentKey) {
    const workspaceLease = await prisma.editorLease.findUnique({
      select: { workspaceObjectKey: true },
      where: { workspaceDocumentKey: metadata.workspaceDocumentKey },
    });
    if (!workspaceLease && metadata.workspaceObjectKey) {
      await deleteObjectUnlessCanonical(metadata.workspaceObjectKey);
    }
  }
}
