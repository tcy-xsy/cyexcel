import type { SceneSnapshot, SurfacePoint } from "../scene";
import type { SelectionModel, SelectionSnapshot } from "./SelectionModel";
import { DRAG_THRESHOLD_PX } from "./PointerController";

export type HeaderSelectionSurface = "column-header" | "row-header" | "corner";

export type HeaderSelectionPoint = {
  readonly [S in HeaderSelectionSurface]: SurfacePoint<S>;
}[HeaderSelectionSurface];

export interface HeaderSelectionModifiers {
  readonly shift?: boolean;
  /** 0 for the primary button and 2 for the secondary button. */
  readonly button?: number;
  readonly isPrimary?: boolean;
}

/** A selection mutation already applied by the controller. */
export interface HeaderSelectionAction {
  readonly surface: HeaderSelectionSurface;
  readonly snapshot: SelectionSnapshot;
}

interface DragState {
  readonly surface: HeaderSelectionSurface;
  readonly startAxisPoint: number;
  readonly thresholdExceeded: boolean;
  readonly structureVersion: number;
}

function action(
  surface: HeaderSelectionSurface,
  snapshot: SelectionSnapshot
): HeaderSelectionAction {
  return Object.freeze({ surface, snapshot });
}

function pointInSurface(
  point: HeaderSelectionPoint,
  scene: SceneSnapshot
): boolean {
  return scene.getSurface(point.surface).panes.some(({ clipRect }) => (
    Number.isFinite(point.x) &&
    Number.isFinite(point.y) &&
    point.x >= clipRect.x &&
    point.y >= clipRect.y &&
    point.x < clipRect.x + clipRect.width &&
    point.y < clipRect.y + clipRect.height
  ));
}

/**
 * Header-only selection gesture state machine. The input owner supplies the
 * surface-local point and one immutable Scene per event; this class owns no
 * DOM listeners, coordinate conversion, resize, or reorder behavior.
 */
export class HeaderSelectionController {
  #drag: DragState | null = null;

  get dragging(): boolean {
    return this.#drag !== null;
  }

  down(
    point: HeaderSelectionPoint,
    scene: SceneSnapshot,
    selection: SelectionModel,
    structureVersion: number,
    modifiers: HeaderSelectionModifiers = {}
  ): HeaderSelectionAction | null {
    this.cancel();
    const button = modifiers.button ?? 0;
    if ((button !== 0 && button !== 2) || modifiers.isPrimary === false) {
      return null;
    }

    if (point.surface === "corner") {
      if (!pointInSurface(point, scene)) {
        return null;
      }
      return action(point.surface, selection.selectAll());
    }

    if (point.surface === "column-header") {
      const hit = scene.hitTestColumnHeader(point);
      if (!hit || hit.band.role !== "leaf") {
        return null;
      }
      if (
        button === 2 &&
        hit.column >= selection.snapshot.range.startColumn &&
        hit.column <= selection.snapshot.range.endColumn
      ) {
        return null;
      }
      const snapshot = modifiers.shift
        ? selection.extendTo(0, hit.column)
        : selection.selectColumn(hit.column);
      if (button === 0) {
        this.#drag = {
          surface: point.surface,
          startAxisPoint: point.x,
          thresholdExceeded: false,
          structureVersion
        };
      }
      return action(point.surface, snapshot);
    }

    const hit = scene.hitTestRowHeader(point);
    if (!hit) {
      return null;
    }
    if (
      button === 2 &&
      hit.row >= selection.snapshot.range.startRow &&
      hit.row <= selection.snapshot.range.endRow
    ) {
      return null;
    }
    const snapshot = modifiers.shift
      ? selection.extendTo(hit.row, 0)
      : selection.selectRow(hit.row);
    if (button === 0) {
      this.#drag = {
        surface: point.surface,
        startAxisPoint: point.y,
        thresholdExceeded: false,
        structureVersion
      };
    }
    return action(point.surface, snapshot);
  }

  move(
    point: HeaderSelectionPoint,
    scene: SceneSnapshot,
    selection: SelectionModel,
    structureVersion: number
  ): HeaderSelectionAction | null {
    const drag = this.#drag;
    if (!drag) {
      return null;
    }
    if (structureVersion !== drag.structureVersion || point.surface !== drag.surface) {
      this.cancel();
      return null;
    }
    if (drag.surface === "corner") {
      return null;
    }

    const axisPoint = drag.surface === "column-header" ? point.x : point.y;
    if (!drag.thresholdExceeded) {
      if (Math.abs(axisPoint - drag.startAxisPoint) < DRAG_THRESHOLD_PX) {
        return null;
      }
      this.#drag = { ...drag, thresholdExceeded: true };
    }

    const snapshot = drag.surface === "column-header"
      ? selection.extendTo(0, scene.resolveColumn(point.x))
      : selection.extendTo(scene.resolveRow(point.y), 0);
    return action(drag.surface, snapshot);
  }

  up(structureVersion: number): void {
    if (this.#drag?.structureVersion !== structureVersion) {
      this.cancel();
      return;
    }
    this.cancel();
  }

  cancel(): void {
    this.#drag = null;
  }
}
