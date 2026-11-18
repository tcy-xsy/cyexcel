export { type Canvas2DLike } from "./Canvas2DLike.js";
export {
  RenderScheduler,
  type FrameScheduler,
  type DirtyState,
  type PresentationDirty,
  type RenderFrame,
  type RenderTarget,
  type RenderSchedulerOptions
} from "./RenderScheduler.js";
export {
  ContentLayerRenderer,
  type CellPaintContext,
  type CellPainter,
  type ContentLayerRendererOptions
} from "./ContentLayerRenderer.js";
export {
  computeCapsule,
  isInDropdownArrowZone,
  isInDropdownHitArea,
  type CapsuleGeometry,
  type CellDropdown,
  type DropdownMode
} from "./dropdownGeometry.js";
export {
  SelectionLayerRenderer,
  type SelectionLayerRendererOptions
} from "./SelectionLayerRenderer.js";
export {
  configureCanvasForDpr,
  getEffectiveDpr,
  readDevicePixelRatio,
  type ResizableCanvas
} from "./dprCanvas.js";
export {
  CanvasSurface,
  type CanvasSurfaceSnapshot,
  type CanvasTransformContext
} from "./CanvasSurface.js";
export { wrapText, computeWrappedCellHeight } from "./textWrap.js";
export { ImageCache, sharedImageCache, type ImageCacheEntry } from "./ImageCache.js";
