import type { CanvasHeaderRenderer, DomHeaderRendererFactory } from "../cells";

export type {
  HeaderState,
  HeaderRenderContext,
  HeaderHitTestContext,
  HeaderAction,
  CanvasHeaderRenderer,
  DomHeaderContext,
  DomHeaderRenderer,
  DomHeaderRendererFactory,
  HeaderDefinition
} from "../cells";

/**
 * A group is either a leaf (declares `columns` directly — the leaf-adjacent
 * row, closest to the column headers) or a branch (declares `children` —
 * its own column span is the union of its children's spans, one row further
 * out). Never both, never neither. Step 9's flat single-level shape
 * (`{id, title, columns}`) is exactly a depth-1 tree under this — every
 * existing config still works unchanged. See `HeaderManager` for how depth
 * is resolved (level 0 = leaf-adjacent, counting outward) and
 * `doc/progress.md`'s Step 19 entry for why branches are allowed to reach
 * different depths.
 */
export interface HeaderGroupDefinition {
  readonly id: string;
  readonly title: string;
  readonly columns?: readonly number[];
  readonly children?: readonly HeaderGroupDefinition[];
}

/** Resolved once at HeaderManager construction and cached; never re-merged when painting. */
export interface RuntimeHeader {
  readonly column: number;
  readonly title: string;
  readonly renderer?: CanvasHeaderRenderer | DomHeaderRendererFactory;
}

export interface RuntimeHeaderGroup {
  readonly id: string;
  readonly title: string;
  readonly startColumn: number;
  readonly endColumn: number;
}
