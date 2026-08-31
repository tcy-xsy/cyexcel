import { describe, expect, it } from "vitest";

import { ColumnAxisSnapshot, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import {
  SceneSnapshot,
  type SceneColumnHeaderBandSpec
} from "../../../../src/core/scene";

function createScene(
  columnHeaderBands: readonly SceneColumnHeaderBandSpec[],
  options: {
    readonly columnWidths?: readonly number[];
    readonly frozenColumnCount?: number;
    readonly scrollLeft?: number;
    readonly width?: number;
    readonly version?: number;
  } = {}
): SceneSnapshot {
  const columnWidths = options.columnWidths ?? [50, 50, 50, 50];
  return new SceneSnapshot({
    version: options.version ?? 1,
    skeleton: new SkeletonSnapshot({
      modelVersion: 1,
      layoutVersion: 1,
      rowHeights: [20],
      columnAxis: new ColumnAxisSnapshot({
        columnWidths,
        frozenColumnCount: options.frozenColumnCount ?? 0
      }),
      frozenRowCount: 0
    }),
    viewport: new ViewportSnapshot({
      version: options.version ?? 1,
      scrollTop: 0,
      scrollLeft: options.scrollLeft ?? 0,
      width: options.width ?? 150,
      height: 20
    }),
    surfaceMetrics: {
      columnHeaderHeight: 60,
      columnHeaderBands,
      rowHeaderWidth: 0
    }
  });
}

const LEAF_BAND: readonly SceneColumnHeaderBandSpec[] = [{ role: "leaf", y: 0, height: 60 }];

describe("SceneSnapshot column-header bands", () => {
  it("publishes an immutable typed descriptor through the column-header layer", () => {
    const scene = createScene(LEAF_BAND);
    const bands = scene.getColumnHeaderBands();
    const layer = scene.getSurface("column-header").layers.find(({ id }) => id === "column-header");

    expect(bands).toMatchObject([
      { id: "leaf", role: "leaf", level: null, rect: { x: 0, y: 0, width: 150, height: 60 } }
    ]);
    expect(layer?.nodes).toBe(bands);
    expect(scene.getColumnHeaderLeafBand()).toBe(bands[0]);
    expect(Object.isFrozen(bands)).toBe(true);
    expect(bands.every((band) => Object.isFrozen(band) && Object.isFrozen(band.rect))).toBe(true);
  });

  it("splits the leaf band exactly once at the frozen boundary", () => {
    const scene = createScene(LEAF_BAND, {
      frozenColumnCount: 1,
      scrollLeft: 20,
      width: 150
    });
    const band = scene.getColumnHeaderLeafBand();

    const fragments = scene.getColumnHeaderBandFragments(band, 0, 3);

    expect(fragments).toMatchObject([
      {
        pane: { id: "frozen" },
        columns: { start: 0, count: 1 },
        rect: { x: 0, y: 0, width: 50, height: 60 }
      },
      {
        pane: { id: "scroll" },
        columns: { start: 1, count: 3 },
        rect: { x: 30, y: 0, width: 150, height: 60 }
      }
    ]);
    expect(fragments.every((fragment) => fragment.band === band)).toBe(true);
  });

  it("rejects a band descriptor from another Scene generation", () => {
    const first = createScene(LEAF_BAND, { version: 1 });
    const second = createScene(LEAF_BAND, { version: 2 });

    expect(() => second.getColumnHeaderBandFragments(first.getColumnHeaderLeafBand(), 0, 0))
      .toThrow("does not belong to this SceneSnapshot");
    expect(() => second.getVisibleColumnHeaderFragments(first.getColumnHeaderLeafBand()))
      .toThrow("does not belong to this SceneSnapshot");
  });

  it.each<{
    readonly name: string;
    readonly bands: readonly SceneColumnHeaderBandSpec[];
  }>([
    { name: "no bands", bands: [] },
    {
      name: "multiple leaves",
      bands: [
        { role: "leaf", y: 0, height: 20 },
        { role: "leaf", y: 20, height: 20 }
      ]
    }
  ])("rejects $name", ({ bands }) => {
    expect(() => createScene(bands)).toThrow(RangeError);
  });

  it("bounds visible frozen traversal independently of a million-column count", () => {
    const scene = createScene(
      [{ role: "leaf", y: 0, height: 60 }],
      {
        columnWidths: new Array<number>(1_000_000).fill(1),
        frozenColumnCount: 500_000,
        width: 20
      }
    );

    const fragments = scene.getVisibleColumnHeaderFragments(scene.getColumnHeaderLeafBand());

    expect(fragments).toHaveLength(1);
    expect(fragments[0]?.pane.id).toBe("frozen");
    expect(fragments[0]?.columns).toEqual({ start: 0, count: 21 });
  });
});
