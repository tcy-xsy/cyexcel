import { encodeTsv } from "./tsv";

/**
 * Narrow interface a real `Worker` satisfies — same "core logic depends on
 * a small interface, not the concrete browser API" pattern as
 * RenderScheduler's `frameScheduler` and EditorManager's `editorFactory`.
 * Lets tests supply a fake worker instead of needing a real Worker global
 * (Node's vitest environment has none).
 */
export interface SerializeWorkerLike {
  postMessage(message: unknown): void;
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  onerror: ((event: unknown) => void) | null;
  terminate(): void;
}

export interface ClipboardWorkerClientOptions {
  /**
   * Not required — `serializeTsv()`/`serializeTsvBlob()` degrade to a
   * synchronous main-thread join when this is absent or throws on
   * construction. The Worker only optimizes the join/escape step for
   * large ranges; it is never load-bearing for correctness.
   */
  readonly createWorker?: () => SerializeWorkerLike;
}

interface SerializeRequestMessage {
  readonly taskId: string;
  readonly storeVersion: number;
  readonly type: "serialize-tsv" | "serialize-tsv-blob";
  readonly payload: { readonly grid: readonly (readonly string[])[] };
}

interface SerializeResponseMessage {
  readonly taskId: string;
  readonly storeVersion: number;
  readonly result: string | Blob;
}

function isSerializeResponse(data: unknown): data is SerializeResponseMessage {
  return typeof data === "object" && data !== null && "taskId" in data && "storeVersion" in data;
}

export interface SerializeTaskOptions {
  readonly storeVersion: number;
  /** Read at response time, not request time — lets a stale result be detected and discarded. */
  readonly currentVersion: () => number;
  readonly signal?: AbortSignal;
}

let nextTaskId = 1;

/**
 * Runs one serialize-and-join task per call, each on a fresh dedicated
 * Worker that gets terminated as soon as it responds (or the task is
 * aborted/superseded) — simpler than a shared long-lived worker pool for
 * this project's scale, and matches architecture.md §10's per-task
 * `taskId`/`storeVersion` protocol example.
 */
export class ClipboardWorkerClient {
  readonly #createWorker: (() => SerializeWorkerLike) | undefined;
  #activeWorker: SerializeWorkerLike | null = null;

  constructor(options: ClipboardWorkerClientOptions = {}) {
    this.#createWorker = options.createWorker;
  }

  serializeTsv(
    grid: readonly (readonly string[])[],
    options: SerializeTaskOptions
  ): Promise<string | undefined> {
    return this.#run("serialize-tsv", grid, options, () => encodeTsv(grid));
  }

  serializeTsvBlob(
    grid: readonly (readonly string[])[],
    options: SerializeTaskOptions
  ): Promise<Blob | undefined> {
    return this.#run(
      "serialize-tsv-blob",
      grid,
      options,
      () => new Blob([encodeTsv(grid)], { type: "text/tab-separated-values" })
    );
  }

  /** Terminates any in-flight task's Worker without resolving its promise — only meant for CyExcelCore.dispose(). */
  dispose(): void {
    this.#activeWorker?.terminate();
    this.#activeWorker = null;
  }

  #run<T extends string | Blob>(
    type: SerializeRequestMessage["type"],
    grid: readonly (readonly string[])[],
    options: SerializeTaskOptions,
    syncFallback: () => T
  ): Promise<T | undefined> {
    if (options.signal?.aborted) {
      return Promise.resolve(undefined);
    }

    let worker: SerializeWorkerLike | null = null;
    try {
      worker = this.#createWorker?.() ?? null;
    } catch {
      worker = null;
    }

    if (!worker) {
      return Promise.resolve(syncFallback());
    }

    this.#activeWorker = worker;
    const taskId = `clip-${nextTaskId}`;
    nextTaskId += 1;

    return new Promise<T | undefined>((resolve) => {
      let settled = false;

      const finish = (value: T | undefined) => {
        if (settled) {
          return;
        }
        settled = true;
        options.signal?.removeEventListener("abort", onAbort);
        worker.terminate();
        if (this.#activeWorker === worker) {
          this.#activeWorker = null;
        }
        resolve(value);
      };

      const onAbort = () => finish(undefined);
      options.signal?.addEventListener("abort", onAbort);

      worker.onerror = () => finish(syncFallback());

      worker.onmessage = (event) => {
        if (!isSerializeResponse(event.data) || event.data.taskId !== taskId) {
          return;
        }
        if (event.data.storeVersion !== options.currentVersion()) {
          finish(undefined);
          return;
        }
        finish(event.data.result as T);
      };

      const request: SerializeRequestMessage = {
        taskId,
        storeVersion: options.storeVersion,
        type,
        payload: { grid }
      };
      worker.postMessage(request);
    });
  }
}
