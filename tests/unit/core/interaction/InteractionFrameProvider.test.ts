import { describe, expect, it, vi } from "vitest";

import {
  InteractionFrameProvider,
  type InteractionFrame
} from "../../../../src/core/interaction";
import { DeferredMutationQueue, RevisionGate } from "../../../../src/core/model";
import { SurfaceHost, clientPoint, type SurfaceHostElement } from "../../../../src/core/scene";

interface TestScene {
  readonly modelVersion: number;
  readonly id: string;
}

interface TestSurfaces {
  readonly version: number;
}

interface MutableSurfaceRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

function surfaceElement(
  rect: MutableSurfaceRect,
  beforeRead?: () => void
): SurfaceHostElement {
  return {
    getBoundingClientRect: () => {
      beforeRead?.();
      return { ...rect };
    }
  };
}

function createIntentQueue(gate: RevisionGate) {
  return new DeferredMutationQueue({
    revisionGate: gate,
    isGuarded: () => false,
    onError: (error) => {
      throw error;
    }
  });
}

describe("InteractionFrameProvider", () => {
  it("refreshes surfaces once per dispatch under the active lease and shares the frame identity", () => {
    const gate = new RevisionGate(7);
    const host = new SurfaceHost({ getTransform: () => "none" });
    host.attach("body", surfaceElement({ left: 10, top: 20, width: 300, height: 200 }));
    const getSurfaces = vi.fn(() => {
      expect(gate.activeReadLeases).toBe(1);
      return host.refresh();
    });
    const provider = new InteractionFrameProvider({
      revisionGate: gate,
      intentQueue: createIntentQueue(gate),
      getScene: () => Object.freeze({ modelVersion: 7, id: "scene" }),
      getSurfaces
    });
    const received: InteractionFrame<TestScene, SurfaceHost["snapshot"]>[][] = [];

    for (let dispatch = 0; dispatch < 2; dispatch += 1) {
      provider.dispatch((frame) => {
        const consumers: typeof frame[] = [];
        const consume = (candidate: typeof frame) => consumers.push(candidate);
        consume(frame);
        consume(frame);
        received.push(consumers);
      });
    }

    expect(getSurfaces).toHaveBeenCalledTimes(2);
    expect(received).toHaveLength(2);
    expect(received[0]?.[0]).toBe(received[0]?.[1]);
    expect(received[1]?.[0]).toBe(received[1]?.[1]);
    expect(received[0]?.[0]).not.toBe(received[1]?.[0]);
    expect(received[0]?.[0]?.surfaces.version).toBe(1);
    expect(received[1]?.[0]?.surfaces.version).toBe(2);
    expect(gate.activeReadLeases).toBe(0);
  });

  it("captures new surface geometry on the next dispatch without mutating the old frame", () => {
    const gate = new RevisionGate(3);
    const rect = { left: 10.25, top: 20.5, width: 300, height: 200 };
    const host = new SurfaceHost({ getTransform: () => "none" });
    host.attach("body", surfaceElement(rect));
    const provider = new InteractionFrameProvider({
      revisionGate: gate,
      intentQueue: createIntentQueue(gate),
      getScene: () => Object.freeze({ modelVersion: 3, id: "scene" }),
      getSurfaces: () => host.refresh()
    });
    const frames: InteractionFrame<TestScene, SurfaceHost["snapshot"]>[] = [];

    provider.dispatch((frame) => frames.push(frame));
    rect.left = 110.75;
    rect.top = 220.125;
    provider.dispatch((frame) => frames.push(frame));

    const first = frames[0];
    const second = frames[1];
    expect(first?.surfaces.version).toBe(1);
    expect(second?.surfaces.version).toBe(2);
    expect(first?.surfaces.get("body")?.clientOrigin).toEqual({ x: 10.25, y: 20.5 });
    expect(second?.surfaces.get("body")?.clientOrigin).toEqual({ x: 110.75, y: 220.125 });
    expect(first?.surfaces.clientToSurface(clientPoint(120.75, 240.125), "body")).toMatchObject({
      x: 110.5,
      y: 219.625
    });
    expect(second?.surfaces.clientToSurface(clientPoint(120.75, 240.125), "body")).toMatchObject({
      x: 10,
      y: 20
    });
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.surfaces)).toBe(true);
    expect(Object.isFrozen(first?.surfaces.get("body")?.clientOrigin)).toBe(true);
  });

  it("keeps the last complete surface snapshot and releases the lease when refresh throws", () => {
    const gate = new RevisionGate(9);
    const bodyRect = { left: 10, top: 20, width: 300, height: 200 };
    const headerRect = { left: 10, top: 0, width: 300, height: 20 };
    let failHeaderMeasurement = false;
    const host = new SurfaceHost({ getTransform: () => "none" });
    host.attach("body", surfaceElement(bodyRect));
    host.attach("column-header", surfaceElement(headerRect, () => {
      if (failHeaderMeasurement) {
        throw new Error("header measurement failed");
      }
    }));
    const getScene = vi.fn(() => Object.freeze({ modelVersion: 9, id: "scene" }));
    const getSurfaces = vi.fn(() => host.refresh());
    const getSelection = vi.fn(() => Object.freeze({ id: "selection" }));
    const getStructureVersion = vi.fn(() => 4);
    const provider = new InteractionFrameProvider({
      revisionGate: gate,
      intentQueue: createIntentQueue(gate),
      getScene,
      getSurfaces,
      getSelection,
      getStructureVersion
    });
    let completeSnapshot: SurfaceHost["snapshot"] | undefined;

    provider.dispatch((frame) => {
      completeSnapshot = frame.surfaces;
    });
    expect(completeSnapshot?.surfaces).toHaveLength(2);
    bodyRect.left = 999;
    failHeaderMeasurement = true;
    const handler = vi.fn();

    expect(() => provider.dispatch(handler)).toThrow("header measurement failed");

    expect(handler).not.toHaveBeenCalled();
    expect(getScene).toHaveBeenCalledTimes(2);
    expect(getSurfaces).toHaveBeenCalledTimes(2);
    expect(getSelection).toHaveBeenCalledTimes(1);
    expect(getStructureVersion).toHaveBeenCalledTimes(1);
    expect(gate.activeReadLeases).toBe(0);
    expect(host.snapshot).toBe(completeSnapshot);
    expect(host.snapshot.version).toBe(1);
    expect(host.snapshot.surfaces).toHaveLength(2);
    expect(host.snapshot.get("body")?.clientOrigin).toEqual({ x: 10, y: 20 });
    expect(host.snapshot.get("column-header")?.clientOrigin).toEqual({ x: 10, y: 0 });
  });

  it("reads every source once and shares one immutable frame identity across consumers", () => {
    const gate = new RevisionGate(7);
    const scene = Object.freeze({ modelVersion: 7, id: "scene" });
    const surfaces = Object.freeze({ version: 11 });
    const selection = Object.freeze({ id: "selection" });
    const getScene = vi.fn(() => scene);
    const getSurfaces = vi.fn(() => surfaces);
    const getSelection = vi.fn(() => selection);
    const getStructureVersion = vi.fn(() => 13);
    const provider = new InteractionFrameProvider({
      revisionGate: gate,
      intentQueue: createIntentQueue(gate),
      getScene,
      getSurfaces,
      getSelection,
      getStructureVersion
    });
    const received: InteractionFrame<TestScene, TestSurfaces, typeof selection>[] = [];

    const result = provider.dispatch((frame) => {
      const firstConsumer = (candidate: typeof frame) => received.push(candidate);
      const secondConsumer = (candidate: typeof frame) => received.push(candidate);
      firstConsumer(frame);
      secondConsumer(frame);
      expect(gate.activeReadLeases).toBe(1);
      expect("release" in frame).toBe(false);
      return "handled";
    });

    expect(result).toBe("handled");
    expect(received).toHaveLength(2);
    expect(received[0]).toBe(received[1]);
    expect(received[0]).toMatchObject({
      revision: 7,
      scene,
      surfaces,
      selection,
      structureVersion: 13
    });
    expect(Object.isFrozen(received[0])).toBe(true);
    expect(getScene).toHaveBeenCalledTimes(1);
    expect(getSurfaces).toHaveBeenCalledTimes(1);
    expect(getSelection).toHaveBeenCalledTimes(1);
    expect(getStructureVersion).toHaveBeenCalledTimes(1);
    expect(gate.activeReadLeases).toBe(0);
  });

  it("does not invoke the handler or read dependent snapshots when the scene revision mismatches", () => {
    const gate = new RevisionGate(2);
    const getScene = vi.fn(() => Object.freeze({ modelVersion: 3, id: "stale" }));
    const getSurfaces = vi.fn(() => Object.freeze({ version: 1 }));
    const getSelection = vi.fn(() => Object.freeze({ id: "selection" }));
    const getStructureVersion = vi.fn(() => 4);
    const handler = vi.fn();
    const provider = new InteractionFrameProvider({
      revisionGate: gate,
      intentQueue: createIntentQueue(gate),
      getScene,
      getSurfaces,
      getSelection,
      getStructureVersion
    });

    expect(() => provider.dispatch(handler)).toThrow(/does not match scene modelVersion/);

    expect(handler).not.toHaveBeenCalled();
    expect(getScene).toHaveBeenCalledTimes(1);
    expect(getSurfaces).not.toHaveBeenCalled();
    expect(getSelection).not.toHaveBeenCalled();
    expect(getStructureVersion).not.toHaveBeenCalled();
    expect(gate.activeReadLeases).toBe(0);
  });

  it("releases the short lease when a handler throws", () => {
    const gate = new RevisionGate(5);
    const provider = new InteractionFrameProvider({
      revisionGate: gate,
      intentQueue: createIntentQueue(gate),
      getScene: () => Object.freeze({ modelVersion: 5 }),
      getSurfaces: () => Object.freeze({ version: 1 })
    });

    expect(() => provider.dispatch(() => {
      expect(gate.activeReadLeases).toBe(1);
      throw new Error("handler failed");
    })).toThrow("handler failed");
    expect(gate.activeReadLeases).toBe(0);
  });

  it("runs enqueued intents after lease release and preserves FIFO order", () => {
    const gate = new RevisionGate(10);
    const events: string[] = [];
    const provider = new InteractionFrameProvider({
      revisionGate: gate,
      intentQueue: createIntentQueue(gate),
      getScene: () => Object.freeze({ modelVersion: 10 }),
      getSurfaces: () => Object.freeze({ version: 1 })
    });

    provider.dispatch(() => {
      events.push(`handler:${gate.activeReadLeases}`);
      provider.enqueueIntent(() => {
        events.push(`first:${gate.activeReadLeases}`);
        gate.commit();
      });
      provider.enqueueIntent(() => {
        events.push(`second:${gate.activeReadLeases}`);
        gate.commit();
      });
      events.push(`handler-end:${gate.activeReadLeases}`);
      expect(gate.revision).toBe(10);
    });

    expect(events).toEqual([
      "handler:1",
      "handler-end:1",
      "first:0",
      "second:0"
    ]);
    expect(gate.revision).toBe(12);
  });
});
