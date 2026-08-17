export {
  ArrayDataSource,
  type ArrayDataSourceColumn,
  type ArrayDataSourceOptions
} from "./ArrayDataSource";
export { TableStore, type TableStoreCommitResult } from "./TableStore";
export { mergePatches } from "./patch";
export {
  createChangeSet,
  compileDirtySet,
  summarizeStructuralChanges,
  type ChangeSet,
  type SemanticChange,
  type DirtySet,
  type StructuralChangeSummary,
  type DirtyRegion
} from "./changes";
