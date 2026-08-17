import type { SceneSnapshot } from "../scene";
import type { HeaderManager } from "./HeaderManager";
import type { RuntimeHeaderGroup } from "./types";

/**
 * Resolves a point (in the same header-canvas-local coordinate space
 * `HeaderCanvasRenderer` paints in — x is scroll/frozen-aware, y never
 * moves with scroll) to the specific group it landed on, if any. Pure
 * geometry, no DOM — mirrors `classifyContextMenuTarget()`/
 * `splitByFrozenBoundary()`'s own "independently testable function" style.
 * Not wired into any real UI yet (no context-menu/click integration) — see
 * doc/progress.md's Step 19 entry for why that's out of scope this round.
 */
export function hitTestHeaderGroup(
  point: { readonly x: number; readonly y: number },
  scene: SceneSnapshot,
  headerManager: HeaderManager
): { readonly level: number; readonly group: RuntimeHeaderGroup } | null {
  const hit = scene.hitTestColumnHeader(point);
  if (!hit || hit.band.role !== "group") {
    return null;
  }
  const group = headerManager
    .getGroups(hit.band.level)
    .find((candidate) => hit.column >= candidate.startColumn && hit.column <= candidate.endColumn);

  return group ? { level: hit.band.level, group } : null;
}
