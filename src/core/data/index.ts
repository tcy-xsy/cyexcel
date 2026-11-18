export { TableStore, type TableStoreCommitResult, type TableStoreOptions } from "./TableStore.js";
export { mergePatches } from "./patch.js";
export {
  createChangeSet,
  compileDirtySet,
  summarizeStructuralChanges,
  type ChangeSet,
  type SemanticChange,
  type DirtySet,
  type StructuralChangeSummary,
  type DirtyRegion
} from "./changes.js";
