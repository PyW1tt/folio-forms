const maxSourceBytes = 128 * 1024;
const maxOutputBytes = 25 * 1024 * 1024;
export const documentWorkerAvailable = async (): Promise<boolean> => {
  const url = process.env.DOCUMENT_WORKER_URL;
  if (!url) {
    return false;
  }
  try {
    const response = await fetch(new URL("/health", url), {
      signal: AbortSignal.timeout(1000),
    });
    return response.ok;
  } catch {
    return false;
  }
};

export const runDocumentWorker = async (
  source: string,
  signal?: AbortSignal
): Promise<Uint8Array> => {
  const url = process.env.DOCUMENT_WORKER_URL;
  if (!url) {
    throw new Error("Document worker is not configured");
  }
  if (Buffer.byteLength(source) > maxSourceBytes) {
    throw new Error("Document source exceeds 128 KiB");
  }
  const response = await fetch(new URL("/run", url), {
    body: source,
    method: "POST",
    signal: AbortSignal.any(
      signal
        ? [signal, AbortSignal.timeout(45_000)]
        : [AbortSignal.timeout(45_000)]
    ),
  });
  if (!response.ok) {
    throw new Error(`Document worker failed (${response.status})`);
  }
  if (Number(response.headers.get("content-length")) > maxOutputBytes) {
    await response.body?.cancel();
    throw new Error("Document output exceeds 25 MiB");
  }
  if (!response.body) {
    throw new Error("Document worker returned no output");
  }
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const value of response.body) {
    total += value.byteLength;
    if (total > maxOutputBytes) {
      throw new Error("Document output exceeds 25 MiB");
    }
    chunks.push(value);
  }
  const output = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
};
