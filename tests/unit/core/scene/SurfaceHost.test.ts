import { describe, expect, it } from "vitest";

import {
  SurfaceHost,
  clientPoint,
  surfacePoint,
  type SurfaceHostElement
} from "../../../../src/core/scene";

interface MutableRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface FakeElement extends SurfaceHostElement {
  readonly parentElement: FakeElement | null;
  readonly transform: string;
}

function element(
  rect: MutableRect,
  options: { readonly parent?: FakeElement; readonly transform?: string } = {}
): FakeElement {
  return {
    parentElement: options.parent ?? null,
    transform: options.transform ?? "none",
    getBoundingClientRect: () => ({ ...rect })
  };
}

describe("SurfaceHost", () => {
  it("captures fractional geometry and converts typed client/surface points through a local region", () => {
    const rect = { left: 10.25, top: 20.5, width: 200.75, height: 100.5 };
    const host = new SurfaceHost({ getTransform: (candidate) => (candidate as FakeElement).transform });
    host.attach("body", element(rect), {
      region: { x: 12.5, y: 7.25, width: 80.5, height: 40.25 },
      zIndex: 3
    });

    const snapshot = host.refresh();
    const body = snapshot.get("body");

    expect(snapshot.version).toBe(1);
    expect(body).toMatchObject({
      id: "body",
      clientOrigin: { x: 22.75, y: 27.75 },
      cssSize: { width: 80.5, height: 40.25 },
      clipRect: { surface: "body", x: 0, y: 0, width: 80.5, height: 40.25 },
      zIndex: 3,
      attachmentVersion: 1
    });
    expect(snapshot.clientToSurface(clientPoint(28.25, 30.5), "body")).toEqual(
      surfacePoint("body", 5.5, 2.75)
    );
    expect(snapshot.surfaceToClient(surfacePoint("body", 5.5, 2.75))).toEqual(
      clientPoint(28.25, 30.5)
    );
    expect(host.clientToSurface(clientPoint(28.25, 30.5), "body")).toEqual(
      surfacePoint("body", 5.5, 2.75)
    );
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(Object.isFrozen(snapshot.surfaces)).toBe(true);
    expect(Object.isFrozen(body)).toBe(true);
  });

  it("picks the highest surface using half-open client bounds", () => {
    const host = new SurfaceHost({ getTransform: () => "none" });
    host.attach("body", element({ left: 10, top: 20, width: 100, height: 50 }), { zIndex: 1 });
    host.attach("menu-overlay", element({ left: 10, top: 20, width: 30, height: 20 }), { zIndex: 10 });
    host.refresh();

    expect(host.pick(clientPoint(10, 20))).toMatchObject({
      surface: "menu-overlay",
      point: { surface: "menu-overlay", x: 0, y: 0 }
    });
    expect(host.pick(clientPoint(39.999, 39.999))?.surface).toBe("menu-overlay");
    expect(host.pick(clientPoint(40, 40))?.surface).toBe("body");
    expect(host.pick(clientPoint(109.999, 69.999))?.surface).toBe("body");
    expect(host.pick(clientPoint(110, 69.999))).toBeNull();
    expect(host.pick(clientPoint(109.999, 70))).toBeNull();
  });

  it("publishes monotonically versioned immutable snapshots", () => {
    const rect = { left: 1.5, top: 2.5, width: 30, height: 40 };
    const host = new SurfaceHost({ getTransform: () => "none" });
    host.attach("column-header", element(rect));

    const first = host.refresh();
    rect.left = 101.5;
    rect.width = 50;
    const second = host.refresh();

    expect(first.version).toBe(1);
    expect(second.version).toBe(2);
    expect(first.get("column-header")?.clientOrigin.x).toBe(1.5);
    expect(first.get("column-header")?.cssSize.width).toBe(30);
    expect(second.get("column-header")?.clientOrigin.x).toBe(101.5);
    expect(second.get("column-header")?.cssSize.width).toBe(50);
  });

  it("detaches by attachment identity and disposes to an empty final snapshot", () => {
    const host = new SurfaceHost({ getTransform: () => "none" });
    const detach = host.attach("row-header", element({ left: 0, top: 0, width: 30, height: 100 }));
    const attached = host.refresh();

    detach();
    expect(host.snapshot.version).toBe(2);
    expect(host.snapshot.get("row-header")).toBeUndefined();
    expect(attached.get("row-header")).toBeDefined();
    detach();
    expect(host.snapshot.version).toBe(2);

    host.attach("row-header", element({ left: 5, top: 6, width: 35, height: 90 }));
    expect(host.refresh().version).toBe(3);
    host.dispose();
    expect(host.snapshot.version).toBe(4);
    expect(host.snapshot.surfaces).toEqual([]);
    expect(host.pick(clientPoint(5, 6))).toBeNull();
    expect(() => host.refresh()).toThrow(/disposed/i);
    expect(() => host.attach("body", element({ left: 0, top: 0, width: 1, height: 1 }))).toThrow(/disposed/i);
  });

  it("rejects transforms on an attachment or ancestor without publishing a partial snapshot", () => {
    const ancestor = element(
      { left: 0, top: 0, width: 200, height: 200 },
      { transform: "matrix(2, 0, 0, 2, 0, 0)" }
    );
    const child = element(
      { left: 10, top: 10, width: 100, height: 50 },
      { parent: ancestor }
    );
    const host = new SurfaceHost({ getTransform: (candidate) => (candidate as FakeElement).transform });
    host.attach("body", child);

    expect(() => host.refresh()).toThrow(/CSS transform/i);
    expect(host.snapshot.version).toBe(0);
    expect(host.snapshot.surfaces).toEqual([]);
  });

  it("validates all local regions before atomically publishing", () => {
    const host = new SurfaceHost({ getTransform: () => "none" });
    host.attach("body", element({ left: 0, top: 0, width: 100, height: 50 }));
    host.attach("corner", element({ left: 0, top: 0, width: 30, height: 20 }), {
      region: { x: 25, y: 0, width: 10, height: 20 }
    });

    expect(() => host.refresh()).toThrow(/region/i);
    expect(host.snapshot.version).toBe(0);
    expect(host.snapshot.surfaces).toEqual([]);
  });
});
