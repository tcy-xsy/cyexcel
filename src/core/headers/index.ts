export {
  HeaderManager,
  type HeaderManagerOptions,
  type HeaderManagerRoot
} from "./HeaderManager.js";
export {
  HeaderCanvasRenderer,
  type HeaderCanvasRendererOptions,
  DEFAULT_HEADER_SORT_ZONE_WIDTH,
  DEFAULT_HEADER_SORT_ZONE_MIN_WIDTH,
  DEFAULT_HEADER_FILTER_ZONE_WIDTH,
  DEFAULT_HEADER_FILTER_ZONE_MIN_WIDTH,
  defaultHeaderHitTest
} from "./HeaderCanvasRenderer.js";
export { RowHeaderCanvasRenderer, type RowHeaderCanvasRendererOptions } from "./RowHeaderCanvasRenderer.js";
export { DomHeaderManager, type DomHeaderManagerOptions } from "./DomHeaderManager.js";
export { columnIndexToLetter } from "./columnLetter.js";
export type {
  HeaderState,
  HeaderRenderContext,
  HeaderHitTestContext,
  HeaderAction,
  CanvasHeaderRenderer,
  DomHeaderContext,
  DomHeaderRenderer,
  DomHeaderRendererFactory,
  HeaderDefinition,
  RuntimeHeader,
} from "./types.js";
