import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { Command } from "../types";

export type RowHeightCommandProvenance = "auto" | "manual";

export interface SetRowHeightUpdate {
  readonly row: number;
  readonly height: number;
}

export interface SetRowHeightsCommandParams {
  readonly updates: readonly SetRowHeightUpdate[];
  readonly provenance: RowHeightCommandProvenance;
}

function invalidUpdate(message: string): CyExcelError {
  return new CyExcelError(CyExcelErrorCode.InvalidOptions, message);
}

/** Immutable user intent for one atomic row-height batch. */
export class SetRowHeightsCommand implements Command<"set-row-heights", SetRowHeightsCommandParams> {
  readonly type = "set-row-heights";
  readonly params: SetRowHeightsCommandParams;

  constructor(
    updates: readonly SetRowHeightUpdate[],
    provenance: RowHeightCommandProvenance
  ) {
    if (updates.length === 0) {
      throw invalidUpdate("SetRowHeightsCommand requires at least one update");
    }

    const rows = new Set<number>();
    const snapshot = updates.map((update, index) => {
      if (!Number.isSafeInteger(update.row) || update.row < 0) {
        throw invalidUpdate(`Invalid row-height update row at index ${index}: ${update.row}`);
      }
      if (!Number.isFinite(update.height) || update.height <= 0) {
        throw invalidUpdate(`Invalid row-height update height at index ${index}: ${update.height}`);
      }
      if (rows.has(update.row)) {
        throw invalidUpdate(`Duplicate row-height update for row ${update.row}`);
      }
      rows.add(update.row);
      return Object.freeze({ row: update.row, height: update.height });
    });

    if (provenance !== "auto" && provenance !== "manual") {
      throw invalidUpdate(`Invalid row-height provenance: ${String(provenance)}`);
    }

    this.params = Object.freeze({
      updates: Object.freeze(snapshot),
      provenance
    });
    Object.freeze(this);
  }
}
