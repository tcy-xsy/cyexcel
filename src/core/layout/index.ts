export {
  LayoutEngine,
  type LayoutOptions,
  type ViewportState,
  type IndexRange,
  type CellRect
} from "./LayoutEngine";
export { toPhysicalPixelSize } from "./dpr";
export { toViewportRect } from "./viewportRect";
export { splitByFrozenBoundary, type AxisSplit } from "./frozenRegions";
export { remapRowIndex, type RowStructureChange } from "./rowStructure";
export { getSelectionBottomRightRect } from "./selectionGeometry";
