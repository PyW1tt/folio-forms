import { chmod, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";

const maxSourceBytes = 128 * 1024;
const maxOutputBytes = 25 * 1024 * 1024;
const workerPort = 3010;
const image = "folio-document-worker:latest";
const workspaceRoot = process.env.WORKER_WORKSPACE_ROOT ?? "/workspaces";

const docker = (args: string[]): Promise<number> =>
  Bun.spawn(["docker", ...args], {
    stderr: "inherit",
    stdin: "ignore",
    stdout: "inherit",
  }).exited;

if (import.meta.main) {
  // The private daemon receives its image before the HTTP endpoint becomes ready.
  // oxlint-disable no-await-in-loop -- Docker readiness checks must run sequentially.
  while (await docker(["info", "--format", "{{.ServerVersion}}"])) {
    await Bun.sleep(1000);
  }
  // oxlint-enable no-await-in-loop
  const context =
    process.env.WORKER_IMAGE_CONTEXT ??
    path.join(import.meta.dir, "../worker-image");
  if ((await docker(["build", "-t", image, context])) !== 0) {
    throw new Error("Unable to build document worker image");
  }

  // A sidecar interruption skips job finally blocks; reap only this worker's state.
  const abandonedJobs = Bun.spawn(
    ["docker", "ps", "-aq", "--filter", "name=^/folio-doc-"],
    { stderr: "inherit", stdin: "ignore", stdout: "pipe" }
  );
  const abandonedOutput = await new Response(abandonedJobs.stdout).text();
  const abandonedIds = abandonedOutput.trim().split(/\s+/u).filter(Boolean);
  if (
    (await abandonedJobs.exited) !== 0 ||
    (abandonedIds.length > 0 &&
      (await docker(["rm", "-f", ...abandonedIds])) !== 0)
  ) {
    throw new Error("Unable to remove abandoned document containers");
  }
  const workspaces = await readdir(workspaceRoot, { withFileTypes: true });
  await Promise.all(
    workspaces
      .filter(
        (entry) => entry.isDirectory() && entry.name.startsWith("session-")
      )
      .map((entry) =>
        rm(path.join(workspaceRoot, entry.name), {
          force: true,
          recursive: true,
        })
      )
  );

  const execute = async (
    source: string,
    signal: AbortSignal
  ): Promise<Uint8Array> => {
    if (Buffer.byteLength(source) > maxSourceBytes) {
      throw new Error("Document source exceeds 128 KiB");
    }
    if (signal.aborted) {
      throw new Error("Document request cancelled");
    }
    const workspace = await mkdtemp(path.join(workspaceRoot, "session-"));
    const name = `folio-doc-${crypto.randomUUID()}`;
    let created = false;
    try {
      await chmod(workspace, 0o755);
      await writeFile(path.join(workspace, "source.py"), source, {
        mode: 0o644,
      });
      // Python source is data mounted read-only; only a size-limited tmpfs can hold output.
      if (
        (await docker([
          "create",
          "--name",
          name,
          "--network",
          "none",
          "--read-only",
          "--cap-drop=ALL",
          "--security-opt",
          "no-new-privileges",
          "--pids-limit",
          "32",
          "--memory",
          "256m",
          "--memory-swap",
          "256m",
          "--cpus",
          "1",
          "--ulimit",
          "cpu=10:10",
          "--shm-size",
          "1m",
          "--user",
          "65532:65532",
          "--workdir",
          "/work",
          "--log-driver",
          "none",
          "--mount",
          `type=bind,source=${workspace}/source.py,target=/input/source.py,readonly`,
          "--mount",
          "type=tmpfs,destination=/work,tmpfs-size=67108864,tmpfs-mode=1777",
          image,
        ])) !== 0
      ) {
        throw new Error("Unable to create document container");
      }
      created = true;
      if (signal.aborted) {
        throw new Error("Document request cancelled");
      }
      const child = Bun.spawn(["docker", "start", "--attach", name], {
        stderr: "ignore",
        stdin: "ignore",
        stdout: "pipe",
      });
      let timedOut = false;
      const stop = () => {
        timedOut = true;
        // Killing Docker CLI alone does not stop the container.
        void docker(["rm", "-f", name]);
        child.kill();
      };
      signal.addEventListener("abort", stop, { once: true });
      const timeout = setTimeout(stop, 30_000);
      try {
        const chunks: Uint8Array[] = [];
        let total = 0;
        for await (const chunk of child.stdout) {
          total += chunk.byteLength;
          if (total > maxOutputBytes) {
            stop();
            throw new Error("Document output exceeds 25 MiB");
          }
          chunks.push(chunk);
        }
        const exit = await child.exited;
        if (signal.aborted || timedOut || exit !== 0 || total === 0) {
          throw new Error("Document execution was rejected or failed");
        }
        const output = new Uint8Array(total);
        let offset = 0;
        for (const chunk of chunks) {
          output.set(chunk, offset);
          offset += chunk.byteLength;
        }
        return output;
      } finally {
        clearTimeout(timeout);
        signal.removeEventListener("abort", stop);
      }
    } finally {
      if (created) {
        await docker(["rm", "-f", name]);
      }
      await rm(workspace, { force: true, recursive: true });
    }
  };

  let active = 0;

  Bun.serve({
    async fetch(request) {
      const { pathname } = new URL(request.url);
      if (pathname === "/health" && request.method === "GET") {
        const check = Bun.spawn(["docker", "image", "inspect", image], {
          stderr: "ignore",
          stdin: "ignore",
          stdout: "ignore",
        });
        return (await check.exited) === 0
          ? new Response("ready")
          : new Response("Document worker unavailable", { status: 503 });
      }
      if (pathname !== "/run" || request.method !== "POST") {
        return new Response("Not found", { status: 404 });
      }
      if (active >= 2) {
        return new Response("Document worker busy", { status: 429 });
      }
      active += 1;
      try {
        const source = await request.text();
        const output = await execute(source, request.signal);
        return new Response(output, {
          headers: {
            "content-type":
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          },
        });
      } catch {
        return new Response("Document execution failed", { status: 422 });
      } finally {
        active -= 1;
      }
    },
    hostname: "0.0.0.0",
    maxRequestBodySize: maxSourceBytes,
    port: workerPort,
  });
}
