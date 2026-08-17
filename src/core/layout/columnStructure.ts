import { remapIndex, type StructuralChange } from "./structuralChange";

export type ColumnStructureChange = StructuralChange;

/** Column-axis wrapper around the shared `remapIndex()` — see `structuralChange.ts` for the actual arithmetic (identical for rows and columns). */
export const remapColumnIndex: (column: number, change: ColumnStructureChange) => number = remapIndex;
