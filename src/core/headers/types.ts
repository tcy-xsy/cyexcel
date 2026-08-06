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

/** Leaf column indices a group spans; must be contiguous — Step 9 only supports two levels (leaves + one group layer), not nested groups. */
export interface HeaderGroupDefinition {
  readonly id: string;
  readonly title: string;
  readonly columns: readonly number[];
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

/**
 * One paintable slice of a group — a group whose column span straddles the
 * frozen/scrollable boundary produces two of these, both referencing the
 * same `group` (same `.id`), one with `frozenColumn: true`.
 */
export interface HeaderGroupFragment {
  readonly group: RuntimeHeaderGroup;
  readonly startColumn: number;
  readonly endColumn: number;
  readonly frozenColumn: boolean;
}
