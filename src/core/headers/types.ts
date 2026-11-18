import type { CanvasHeaderRenderer, DomHeaderRendererFactory } from "../cells/index.js";

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
} from "../cells/index.js";

/** Resolved once at HeaderManager construction and cached; never re-merged when painting. */
export interface RuntimeHeader {
  readonly column: number;
  readonly title: string;
  readonly renderer?: CanvasHeaderRenderer | DomHeaderRendererFactory;
}
