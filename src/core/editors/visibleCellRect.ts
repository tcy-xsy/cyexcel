import type { CellRect } from "../layout";
import type { SceneSnapshot } from "../scene";
import type { CellRange } from "../types";

export function intersectRect(rect: CellRect, clipRect: CellRect): CellRect | null {
  const left = Math.max(rect.x, clipRect.x);
  const top = Math.max(rect.y, clipRect.y);
  const right = Math.min(rect.x + rect.width, clipRect.x + clipRect.width);
  const bottom = Math.min(rect.y + rect.height, clipRect.y + clipRect.height);
  if (right <= left || bottom <= top) {
    return null;
  }
  return { x: left, y: top, width: right - left, height: bottom - top };
}

/**
 * `range` 在当前 Scene 里的可见矩形（编辑层坐标系）：取第一个片段，再夹到它所在面板
 * 的裁剪框里。完全滚出视口返回 `null`。编辑器和校验提示都靠它把 DOM 摆到单元格上。
 */
export function visibleRectForRange(scene: SceneSnapshot, range: CellRange): CellRect | null {
  const fragment = scene.getFragmentsForRange(range)[0];
  return fragment ? intersectRect(fragment.rect, fragment.pane.clipRect) : null;
}
