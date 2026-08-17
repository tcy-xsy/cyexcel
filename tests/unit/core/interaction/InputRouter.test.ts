import { describe, expect, it, vi } from "vitest";

import {
  InputRouter,
  createEventPathSurfacePolicy,
  declineInput,
  handleInput,
  nativeInputTarget,
  surfaceInputTarget,
  type InteractionFrame
} from "../../../../src/core/interaction";
import { RevisionGate } from "../../../../src/core/model";
import {
  SurfaceHost,
  clientPoint,
  type SceneSurfaceId,
  type SurfaceHostElement
} from "../../../../src/core/scene";

interface TestScene {
  readonly modelVersion: number;
}

interface TestEvent {
  readonly type: string;
}

type TestFrame = InteractionFrame<TestScene, SurfaceHost["snapshot"], { readonly row: number }>;

function element(left: number, top: number, width: number, height: number): SurfaceHostElement {
  return {
    parentElement: null,
    getBoundingClientRect: () => ({ left, top, width, height })
  };
}

function createFrame(): TestFrame {
  const gate = new RevisionGate(4);
  const host = new SurfaceHost({ getTransform: () => "none" });
  host.attach("body", element(100, 50, 300, 200), { zIndex: 1 });
  host.attach("menu-overlay", element(100, 50, 300, 200), { zIndex: 100 });
  return Object.freeze({
    revision: gate.revision,
    scene: Object.freeze({ modelVersion: 4 }),
    surfaces: host.refresh(),
    selection: Object.freeze({ row: 2 }),
    structureVersion: 8
  });
}

describe("InputRouter", () => {
  it("uses the nearest event-path owner instead of topmost surface geometry", () => {
    const bodyCanvas = {};
    const sharedScroller = {};
    const owners = new Map<object, SceneSurfaceId>([
      [bodyCanvas, "body"],
      [sharedScroller, "menu-overlay"]
    ]);
    const router = new InputRouter<TestFrame, TestEvent, string, object>({
      surfacePolicy: createEventPathSurfacePolicy((entry) => owners.get(entry))
    });
    router.register({
      id: "body-selection",
      surface: "body",
      route: ({ point }) => handleInput(`${point.surface}:${point.x},${point.y}`)
    });
    router.register({
      id: "menu",
      surface: "menu-overlay",
      route: () => handleInput("menu")
    });

    const result = router.dispatch(createFrame(), {
      event: { type: "pointerdown" },
      clientPoint: clientPoint(125, 80),
      path: [bodyCanvas, sharedScroller]
    });

    expect(result).toMatchObject({
      handled: true,
      candidateId: "body-selection",
      surface: "body",
      point: { surface: "body", x: 25, y: 30 },
      value: "body:25,30"
    });
  });

  it("tries one surface's candidates by priority and stops at the first acceptance", () => {
    const calls: string[] = [];
    const router = new InputRouter<TestFrame, TestEvent, string>();
    router.register({
      id: "low",
      surface: "body",
      priority: 0,
      route: () => {
        calls.push("low:route");
        return handleInput("low");
      }
    });
    router.register({
      id: "high-reject",
      surface: "body",
      priority: 100,
      route: () => {
        calls.push("high-reject:route");
        return declineInput();
      }
    });
    router.register({
      id: "high-accept",
      surface: "body",
      priority: 100,
      route: () => {
        calls.push("high-accept:route");
        return handleInput("winner");
      }
    });

    const result = router.dispatch(createFrame(), {
      event: { type: "pointerdown" },
      clientPoint: clientPoint(110, 60),
      path: [],
      surface: "body"
    });

    expect(calls).toEqual([
      "high-reject:route",
      "high-accept:route"
    ]);
    expect(result).toMatchObject({ handled: true, candidateId: "high-accept", value: "winner" });
  });

  it("passes the caller's exact frame and one context identity to policy, accepts, and handle", () => {
    const frame = createFrame();
    const policyFrames: TestFrame[] = [];
    const candidateFrames: TestFrame[] = [];
    const contexts: object[] = [];
    const router = new InputRouter<TestFrame, TestEvent, void>({
      surfacePolicy: (context) => {
        policyFrames.push(context.frame);
        return surfaceInputTarget("body");
      }
    });
    router.register({
      id: "identity",
      surface: "body",
      route: (context) => {
        candidateFrames.push(context.frame);
        contexts.push(context);
        return handleInput(undefined);
      }
    });

    router.dispatch(frame, {
      event: { type: "dblclick" },
      clientPoint: clientPoint(130, 90),
      path: []
    });

    expect(policyFrames).toEqual([frame]);
    expect(candidateFrames).toEqual([frame]);
    expect(contexts).toHaveLength(1);
    expect(Object.isFrozen(contexts[0])).toBe(true);
  });

  it("lets an explicit surface override policy and reports non-routable inputs without falling back", () => {
    const policy = vi.fn(() => surfaceInputTarget("menu-overlay"));
    const router = new InputRouter<TestFrame, TestEvent, void>({ surfacePolicy: policy });

    expect(router.dispatch(createFrame(), {
      event: { type: "contextmenu" },
      clientPoint: clientPoint(101, 51),
      path: [],
      surface: "body"
    })).toMatchObject({ handled: false, reason: "not-accepted", surface: "body" });
    expect(policy).not.toHaveBeenCalled();

    const noPolicyRouter = new InputRouter<TestFrame, TestEvent, void>();
    expect(noPolicyRouter.dispatch(createFrame(), {
      event: { type: "pointerdown" },
      clientPoint: clientPoint(101, 51),
      path: []
    })).toEqual({ handled: false, reason: "no-surface", surface: null, point: null });

    expect(noPolicyRouter.dispatch(createFrame(), {
      event: { type: "pointerdown" },
      clientPoint: clientPoint(101, 51),
      path: [],
      surface: "column-header"
    })).toEqual({
      handled: false,
      reason: "surface-unavailable",
      surface: "column-header",
      point: null
    });
  });

  it("treats native ownership as terminal without invoking grid candidates", () => {
    const route = vi.fn(() => handleInput("unexpected"));
    const router = new InputRouter<TestFrame, TestEvent, string>({
      surfacePolicy: () => nativeInputTarget()
    });
    router.register({ id: "selection", surface: "body", route });

    expect(router.dispatch(createFrame(), {
      event: { type: "pointerdown" },
      clientPoint: clientPoint(110, 60),
      path: [{}]
    })).toEqual({
      handled: false,
      reason: "native-owned",
      surface: null,
      point: null
    });
    expect(route).not.toHaveBeenCalled();
  });

  it("unregisters candidates by identity and rejects use after disposal", () => {
    const router = new InputRouter<TestFrame, TestEvent, string>();
    const disposeCandidate = router.register({
      id: "selection",
      surface: "body",
      route: () => handleInput("handled")
    });
    disposeCandidate();
    disposeCandidate();

    expect(router.dispatch(createFrame(), {
      event: { type: "pointerdown" },
      clientPoint: clientPoint(101, 51),
      path: [],
      surface: "body"
    })).toMatchObject({ handled: false, reason: "not-accepted" });

    router.dispose();
    router.dispose();
    expect(() => router.dispatch(createFrame(), {
      event: { type: "pointerdown" },
      clientPoint: clientPoint(101, 51),
      path: [],
      surface: "body"
    })).toThrow(/disposed/i);
    expect(() => router.register({
      id: "late",
      surface: "body",
      route: () => handleInput("late")
    })).toThrow(/disposed/i);
  });
});
