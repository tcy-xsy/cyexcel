import { encodeTsv } from "../tsv";

/**
 * Runs in a Worker's global scope, not a browsing context — `self` there is
 * a `DedicatedWorkerGlobalScope`, but this project's tsconfig only includes
 * the DOM lib (needed everywhere else), and DOM's `Window.postMessage`
 * declares an incompatible signature from the Worker one. Going through
 * `globalThis` with this narrow local type sidesteps that lib conflict
 * instead of mixing in the separate (and DOM-incompatible) WebWorker lib
 * project-wide — same "depend on a small interface, not the concrete
 * global" pattern as `Canvas2DLike`/`SerializeWorkerLike`.
 */
interface SerializeWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: unknown): void;
}

interface SerializeRequest {
  readonly taskId: string;
  readonly storeVersion: number;
  readonly type: "serialize-tsv" | "serialize-tsv-blob";
  readonly payload: { readonly grid: readonly (readonly string[])[] };
}

const scope = globalThis as unknown as SerializeWorkerScope;

scope.onmessage = (event) => {
  const request = event.data as SerializeRequest;
  const text = encodeTsv(request.payload.grid);
  // The download branch builds the Blob here, in the Worker, and posts the
  // Blob itself (structured-clonable) back — the main thread never has to
  // re-stringify the already-huge text just to wrap it in a Blob.
  const result: string | Blob =
    request.type === "serialize-tsv-blob" ? new Blob([text], { type: "text/tab-separated-values" }) : text;

  scope.postMessage({
    taskId: request.taskId,
    storeVersion: request.storeVersion,
    result
  });
};
