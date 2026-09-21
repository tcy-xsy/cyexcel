export { TableStore, type TableStoreCommitResult, type TableStoreOptions } from "./TableStore";
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
