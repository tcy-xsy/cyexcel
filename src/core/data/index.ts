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
  type ChangeSet,
  type SemanticChange,
  type DirtySet,
  type DirtyRegion
} from "./changes";
