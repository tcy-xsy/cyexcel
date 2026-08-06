import { describe, expect, it } from "vitest";

import { ClipboardWorkerClient, encodeTsv } from "../../../../src/core";
import type { SerializeWorkerLike } from "../../../../src/core";

interface FakeWorker extends SerializeWorkerLike {
  readonly sent: unknown[];
  terminated: boolean;
  respond(data: unknown): void;
  triggerError(): void;
}

function createFakeWorker(): FakeWorker {
  const worker: FakeWorker = {
    sent: [],
    terminated: false,
    onmessage: null,
    onerror: null,
    postMessage(message) {
      worker.sent.push(message);
    },
    terminate() {
      worker.terminated = true;
    },
    respond(data) {
      worker.onmessage?.({ data });
    },
    triggerError() {
      worker.onerror?.(new Error("worker crashed"));
    }
  };
  return worker;
}

function lastRequest(worker: FakeWorker): { taskId: string; storeVersion: number } {
  const request = worker.sent[worker.sent.length - 1] as { taskId: string; storeVersion: number } | undefined;
  if (!request) {
    throw new Error("no request sent to the fake worker");
  }
  return request;
}

const GRID = [["a", "b"], ["c", "d"]];

describe("ClipboardWorkerClient — worker path", () => {
  it("resolves with the worker's result when the response's storeVersion matches currentVersion()", async () => {
    const worker = createFakeWorker();
    const client = new ClipboardWorkerClient({ createWorker: () => worker });

    const resultPromise = client.serializeTsv(GRID, { storeVersion: 5, currentVersion: () => 5 });
    const request = lastRequest(worker);
    worker.respond({ taskId: request.taskId, storeVersion: 5, result: "custom-result" });

    expect(await resultPromise).toBe("custom-result");
    expect(worker.terminated).toBe(true);
  });

  it("discards the result when the store has moved to a newer version by the time the worker responds", async () => {
    const worker = createFakeWorker();
    const client = new ClipboardWorkerClient({ createWorker: () => worker });

    let currentVersion = 1;
    const resultPromise = client.serializeTsv(GRID, { storeVersion: 1, currentVersion: () => currentVersion });
    const request = lastRequest(worker);
    currentVersion = 2; // a write landed while the worker was busy
    worker.respond({ taskId: request.taskId, storeVersion: request.storeVersion, result: "stale-result" });

    expect(await resultPromise).toBeUndefined();
  });

  it("ignores a message whose taskId doesn't match this task's request", async () => {
    const worker = createFakeWorker();
    const client = new ClipboardWorkerClient({ createWorker: () => worker });

    const resultPromise = client.serializeTsv(GRID, { storeVersion: 1, currentVersion: () => 1 });
    worker.respond({ taskId: "some-other-task", storeVersion: 1, result: "wrong" });
    const request = lastRequest(worker);
    worker.respond({ taskId: request.taskId, storeVersion: 1, result: "right" });

    expect(await resultPromise).toBe("right");
  });

  it("falls back to the synchronous join when the worker reports a runtime error", async () => {
    const worker = createFakeWorker();
    const client = new ClipboardWorkerClient({ createWorker: () => worker });

    const resultPromise = client.serializeTsv(GRID, { storeVersion: 1, currentVersion: () => 1 });
    worker.triggerError();

    expect(await resultPromise).toBe(encodeTsv(GRID));
    expect(worker.terminated).toBe(true);
  });

  it("resolves undefined and terminates the worker when the AbortSignal fires before a response arrives", async () => {
    const worker = createFakeWorker();
    const client = new ClipboardWorkerClient({ createWorker: () => worker });
    const controller = new AbortController();

    const resultPromise = client.serializeTsv(GRID, {
      storeVersion: 1,
      currentVersion: () => 1,
      signal: controller.signal
    });
    controller.abort();

    expect(await resultPromise).toBeUndefined();
    expect(worker.terminated).toBe(true);
  });

  it("never constructs a worker when the signal is already aborted", async () => {
    const controller = new AbortController();
    controller.abort();
    let createCount = 0;
    const client = new ClipboardWorkerClient({
      createWorker: () => {
        createCount += 1;
        return createFakeWorker();
      }
    });

    const result = await client.serializeTsv(GRID, {
      storeVersion: 1,
      currentVersion: () => 1,
      signal: controller.signal
    });

    expect(result).toBeUndefined();
    expect(createCount).toBe(0);
  });

  it("dispose() terminates an in-flight task's worker", () => {
    const worker = createFakeWorker();
    const client = new ClipboardWorkerClient({ createWorker: () => worker });

    void client.serializeTsv(GRID, { storeVersion: 1, currentVersion: () => 1 });
    client.dispose();

    expect(worker.terminated).toBe(true);
  });
});

describe("ClipboardWorkerClient — synchronous fallback", () => {
  it("serializes on the main thread when no createWorker is configured", async () => {
    const client = new ClipboardWorkerClient();

    const result = await client.serializeTsv(GRID, { storeVersion: 1, currentVersion: () => 1 });

    expect(result).toBe(encodeTsv(GRID));
  });

  it("serializes on the main thread when createWorker throws during construction", async () => {
    const client = new ClipboardWorkerClient({
      createWorker: () => {
        throw new Error("Worker is not supported in this environment");
      }
    });

    const result = await client.serializeTsv(GRID, { storeVersion: 1, currentVersion: () => 1 });

    expect(result).toBe(encodeTsv(GRID));
  });

  it("produces a Blob with the same TSV text via serializeTsvBlob's sync fallback", async () => {
    const client = new ClipboardWorkerClient();

    const blob = await client.serializeTsvBlob(GRID, { storeVersion: 1, currentVersion: () => 1 });

    expect(blob).toBeInstanceOf(Blob);
    expect(await blob?.text()).toBe(encodeTsv(GRID));
  });
});
