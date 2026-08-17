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

const BANDS: readonly SceneColumnHeaderBandSpec[] = [
  { role: "group", level: 1, y: 0, height: 20 },
  { role: "group", level: 0, y: 20, height: 20 },
  { role: "leaf", y: 40, height: 20 }
];

describe("SceneSnapshot column-header bands", () => {
  it("publishes immutable typed descriptors through the column-header layer", () => {
    const scene = createScene(BANDS);
    const bands = scene.getColumnHeaderBands();
    const layer = scene.getSurface("column-header").layers.find(({ id }) => id === "column-header");

    expect(bands).toMatchObject([
      { id: "group:1", role: "group", level: 1, rect: { x: 0, y: 0, width: 150, height: 20 } },
      { id: "group:0", role: "group", level: 0, rect: { x: 0, y: 20, width: 150, height: 20 } },
      { id: "leaf", role: "leaf", level: null, rect: { x: 0, y: 40, width: 150, height: 20 } }
    ]);
    expect(layer?.nodes).toBe(bands);
    expect(scene.getColumnHeaderLeafBand()).toBe(bands[2]);
    expect(scene.getColumnHeaderGroupBand(0)).toBe(bands[1]);
    expect(scene.getColumnHeaderGroupBand(2)).toBeNull();
    expect(Object.isFrozen(bands)).toBe(true);
    expect(bands.every((band) => Object.isFrozen(band) && Object.isFrozen(band.rect))).toBe(true);
  });

  it("splits one group span exactly once at the frozen boundary", () => {
    const scene = createScene(BANDS, {
      frozenColumnCount: 1,
      scrollLeft: 20,
      width: 150
    });
    const band = scene.getColumnHeaderGroupBand(0);
    expect(band).not.toBeNull();

    const fragments = scene.getColumnHeaderBandFragments(band!, 0, 3);

    expect(fragments).toMatchObject([
      {
        pane: { id: "frozen" },
        columns: { start: 0, count: 1 },
        rect: { x: 0, y: 20, width: 50, height: 20 }
      },
      {
        pane: { id: "scroll" },
        columns: { start: 1, count: 3 },
        rect: { x: 30, y: 20, width: 150, height: 20 }
      }
    ]);
    expect(fragments.every((fragment) => fragment.band === band)).toBe(true);
  });

  it("rejects a band descriptor from another Scene generation", () => {
    const first = createScene(BANDS, { version: 1 });
    const second = createScene(BANDS, { version: 2 });

    expect(() => second.getColumnHeaderBandFragments(first.getColumnHeaderLeafBand(), 0, 0))
      .toThrow("does not belong to this SceneSnapshot");
    expect(() => second.getVisibleColumnHeaderFragments(first.getColumnHeaderLeafBand()))
      .toThrow("does not belong to this SceneSnapshot");
  });

  it.each<{
    readonly name: string;
    readonly bands: readonly SceneColumnHeaderBandSpec[];
  }>([
    { name: "missing leaf", bands: [{ role: "group", level: 0, y: 0, height: 20 }] },
    {
      name: "multiple leaves",
      bands: [
        { role: "leaf", y: 0, height: 20 },
        { role: "leaf", y: 20, height: 20 }
      ]
    },
    {
      name: "duplicate group levels",
      bands: [
        { role: "group", level: 0, y: 0, height: 20 },
        { role: "group", level: 0, y: 20, height: 20 },
        { role: "leaf", y: 40, height: 20 }
      ]
    },
    {
      name: "non-contiguous group levels",
      bands: [
        { role: "group", level: 1, y: 0, height: 20 },
        { role: "leaf", y: 40, height: 20 }
      ]
    },
    {
      name: "group levels ordered toward rather than away from the leaf",
      bands: [
        { role: "group", level: 0, y: 0, height: 20 },
        { role: "group", level: 1, y: 20, height: 20 },
        { role: "leaf", y: 40, height: 20 }
      ]
    },
    {
      name: "overlapping bands",
      bands: [
        { role: "group", level: 0, y: 0, height: 30 },
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
