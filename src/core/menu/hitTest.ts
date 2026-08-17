import type { SelectionSnapshot } from "../interaction";
import type { SceneSnapshot } from "../scene";
import type { ContextMenuContext } from "./types";

function containsPoint(
  rect: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  point: { readonly x: number; readonly y: number }
): boolean {
  return point.x >= rect.x && point.y >= rect.y &&
    point.x < rect.x + rect.width && point.y < rect.y + rect.height;
}

/**
 * Classifies a point (already relative to the scroller/content surface origin,
 * same convention PointerController uses) into the five
 * doc/customization-and-hooks.md §3.6 targets. Reuses SceneSnapshot's
 * "past the data extent returns null" hit-test semantics for "blank".
 */
export function classifyContextMenuTarget(
  point: { readonly x: number; readonly y: number },
  scene: SceneSnapshot,
  selection: SelectionSnapshot
): ContextMenuContext {
  const corner = scene.getSurface("corner").rootRect;
  const rowHeader = scene.getSurface("row-header").rootRect;
  const columnHeader = scene.getSurface("column-header").rootRect;
  const body = scene.getSurface("body").rootRect;

  if (containsPoint(corner, point)) {
    return { target: "corner", selection, point };
  }

  if (containsPoint(rowHeader, point)) {
    const hit = scene.hitTestRowHeader({
      x: point.x - rowHeader.x,
      y: point.y - rowHeader.y
    });
    return { target: "row-header", selection, point, ...(hit ? { row: hit.row } : {}) };
  }

  if (containsPoint(columnHeader, point)) {
    const hit = scene.hitTestColumnHeader({
      x: point.x - columnHeader.x,
      y: point.y - columnHeader.y
    });
    return { target: "column-header", selection, point, ...(hit ? { column: hit.column } : {}) };
  }

  if (!containsPoint(body, point)) {
    return { target: "blank", selection, point };
  }
  const gridPoint = { x: point.x - body.x, y: point.y - body.y };
  const hit = scene.hitTestCellGeometry(gridPoint);
  if (!hit) {
    return { target: "blank", selection, point };
  }
  return {
    target: "cell",
    selection,
    point,
    row: hit.cell.row,
    column: hit.cell.column
  };
}
