import { remapIndex, type StructuralChange } from "./structuralChange";

export type RowStructureChange = StructuralChange;

/** Row-axis wrapper around the shared `remapIndex()` — see `structuralChange.ts` for the actual arithmetic (identical for rows and columns). */
export const remapRowIndex: (row: number, change: RowStructureChange) => number = remapIndex;
