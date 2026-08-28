import type { ColumnManager, RuntimeColumnRoot } from "../cells/ColumnManager";
import { snapshotColumnDefinitions } from "../cells/columnSnapshot";
import type { ColumnDefinition } from "../cells/types";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { HeaderManager, HeaderManagerRoot } from "../headers/HeaderManager";
import type { HeaderGroupDefinition } from "../headers/types";
import type {
  LayoutRoot,
  LayoutStructurePayload,
  RowHeightProvenance
} from "../layout/LayoutEngine";
import type { LayoutEngine } from "../layout/LayoutEngine";
import type { RowMetricsFragment } from "../layout/RowMetrics";
import type { AxisChangeDescriptor, StructuralChange } from "../layout/structuralChange";
import type { SelectionModel, SelectionRoot } from "../interaction/SelectionModel";
import type { MutationPrepareContext } from "../mutations/types";
import type {
  PreparedStructureIntegration,
  StructureIntegrationOwner
} from "../mutations/StructureIntegrationMutation";

export interface GridStructureRoot {
  readonly columns: readonly ColumnDefinition[];
  readonly runtimeColumns: RuntimeColumnRoot;
  readonly headers: HeaderManagerRoot;
  readonly layout: LayoutRoot;
  readonly selection: SelectionRoot;
}

export interface GridStructurePayload {
  readonly columns?: readonly ColumnDefinition[];
  readonly rowHeights?: ReadonlyMap<number, number>;
  readonly rowMetricsFragment?: RowMetricsFragment;
  readonly rowHeightProvenance?: ReadonlyMap<
    number,
    Exclude<RowHeightProvenance, "default">
  >;
  readonly columnWidths?: readonly number[];
  readonly hiddenColumnOffsets?: ReadonlyMap<number, number>;
}

export interface GridStructureOwnerOptions {
  readonly getColumns: () => readonly ColumnDefinition[];
  readonly replaceColumns: (columns: readonly ColumnDefinition[]) => void;
  readonly headerGroups?: readonly HeaderGroupDefinition[];
  readonly columnManager: ColumnManager;
  readonly headerManager: HeaderManager;
  readonly layout: LayoutEngine;
  readonly selection: SelectionModel;
}

function toStructuralChange(change: AxisChangeDescriptor): StructuralChange {
  return {
    operation: change.operation,
    start: change.start,
    count: change.count,
    ...(change.target !== undefined ? { target: change.target } : {})
  };
}

function invertMove(change: AxisChangeDescriptor): AxisChangeDescriptor {
  const target = change.target as number;
  const adjustedTarget = target > change.start ? target - change.count : target;
  const undoTarget = change.start > adjustedTarget ? change.start + change.count : change.start;
  return {
    axis: change.axis,
    operation: "move",
    start: adjustedTarget,
    count: change.count,
    target: undoTarget,
    targetSpace: "pre-removal"
  };
}

function inverseChange(change: AxisChangeDescriptor): AxisChangeDescriptor {
  if (change.operation === "move") {
    return invertMove(change);
  }
  return {
    axis: change.axis,
    operation: change.operation === "insert" ? "delete" : "insert",
    start: change.start,
    count: change.count,
    targetSpace: "pre-removal"
  };
}

function moveColumns(
  columns: readonly ColumnDefinition[],
  start: number,
  count: number,
  target: number
): readonly ColumnDefinition[] {
  const adjustedTarget = target > start ? target - count : target;
  const remaining = columns.slice(0, start).concat(columns.slice(start + count));
  return Object.freeze(
    remaining.slice(0, adjustedTarget).concat(columns.slice(start, start + count), remaining.slice(adjustedTarget))
  );
}

/** Composite owner for every non-SheetModel root coupled to a structural edit. */
export class GridStructureOwner
  implements StructureIntegrationOwner<GridStructureRoot, GridStructurePayload>
{
  readonly #columnManager: ColumnManager;
  readonly #headerManager: HeaderManager;
  readonly #layout: LayoutEngine;
  readonly #selection: SelectionModel;
  readonly #headerGroups: readonly HeaderGroupDefinition[];
  readonly #getColumns: () => readonly ColumnDefinition[];
  readonly #replaceColumns: (columns: readonly ColumnDefinition[]) => void;

  constructor(options: GridStructureOwnerOptions) {
    this.#getColumns = options.getColumns;
    this.#replaceColumns = options.replaceColumns;
    this.#headerGroups = options.headerGroups ?? [];
    this.#columnManager = options.columnManager;
    this.#headerManager = options.headerManager;
    this.#layout = options.layout;
    this.#selection = options.selection;
  }

  get columns(): readonly ColumnDefinition[] {
    return this.#getColumns();
  }

  getRoot(): GridStructureRoot {
    return Object.freeze({
      columns: this.#getColumns(),
      runtimeColumns: this.#columnManager.getColumns(),
      headers: this.#headerManager.root,
      layout: this.#layout.root,
      selection: this.#selection.root
    });
  }

  isCurrent(root: GridStructureRoot): boolean {
    return this.#getColumns() === root.columns &&
      this.#columnManager.getColumns() === root.runtimeColumns &&
      this.#headerManager.root === root.headers &&
      this.#layout.root === root.layout &&
      this.#selection.root === root.selection;
  }

  getPlanningRoot(root: GridStructureRoot, context: MutationPrepareContext): GridStructureRoot {
    const layout = context.roots.get(this.#layout, () => root.layout);
    if (layout === root.layout) {
      return root;
    }
    return Object.freeze({ ...root, layout });
  }

  prepare(
    root: GridStructureRoot,
    change: AxisChangeDescriptor,
    payload: GridStructurePayload,
    context: MutationPrepareContext
  ): PreparedStructureIntegration<GridStructureRoot, GridStructurePayload> {
    const structuralChange = toStructuralChange(change);
    let inverseRowMetricsFragment: RowMetricsFragment | undefined;
    let inverseRowHeightProvenance: ReadonlyMap<
      number,
      Exclude<RowHeightProvenance, "default">
    > | undefined;
    let inverseColumns: readonly ColumnDefinition[] | undefined;
    let inverseColumnWidths: readonly number[] | undefined;
    let inverseHiddenColumnOffsets: ReadonlyMap<number, number> | undefined;
    let forwardPayload: GridStructurePayload = {};
    let columns = root.columns;
    let runtimeColumns = root.runtimeColumns;
    let headers = root.headers;
    let layoutPayload: LayoutStructurePayload = {};

    if (change.axis === "row") {
      if (change.operation === "insert") {
        layoutPayload = {
          rowHeights: payload.rowHeights,
          rowMetricsFragment: payload.rowMetricsFragment,
          rowHeightProvenance: payload.rowHeightProvenance
        };
        forwardPayload = createGridStructurePayload(
          payload.rowMetricsFragment !== undefined
            ? {
                rowMetricsFragment: payload.rowMetricsFragment,
                ...(payload.rowHeightProvenance !== undefined
                  ? { rowHeightProvenance: payload.rowHeightProvenance }
                  : {})
              }
            : payload.rowHeights !== undefined
              ? {
                  rowHeights: payload.rowHeights,
                  ...(payload.rowHeightProvenance !== undefined
                    ? { rowHeightProvenance: payload.rowHeightProvenance }
                    : {})
                }
              : {}
        );
      } else if (change.operation === "delete") {
        inverseRowMetricsFragment = root.layout.rowMetrics.slice(change.start, change.count);
        inverseRowHeightProvenance = new Map(
          Array.from(root.layout.rowHeightProvenance)
            .filter(([row]) => row >= change.start && row < change.start + change.count)
            .map(([row, provenance]) => [row - change.start, provenance])
        );
      }
    } else {
      if (this.#headerGroups.length > 0) {
        throw new CyExcelError(
          CyExcelErrorCode.HeaderGroupsBlockColumnStructureChange,
          "Column structure changes require header-group remapping"
        );
      }
      if (change.operation === "insert") {
        if (payload.columns?.length !== change.count || payload.columnWidths?.length !== change.count) {
          throw new CyExcelError(CyExcelErrorCode.InvalidOptions, "Column insertion payload does not match count");
        }
        const inserted = snapshotColumnDefinitions(payload.columns);
        columns = Object.freeze(root.columns.slice(0, change.start).concat(inserted, root.columns.slice(change.start)));
        layoutPayload = {
          columnWidths: payload.columnWidths,
          hiddenColumnOffsets: payload.hiddenColumnOffsets
        };
        forwardPayload = createGridStructurePayload({
          columns: inserted,
          columnWidths: payload.columnWidths,
          ...(payload.hiddenColumnOffsets !== undefined
            ? { hiddenColumnOffsets: payload.hiddenColumnOffsets }
            : {})
        });
      } else if (change.operation === "delete") {
        inverseColumns = root.columns.slice(change.start, change.start + change.count);
        const widths: number[] = [];
        const hiddenOffsets = new Map<number, number>();
        for (let offset = 0; offset < change.count; offset += 1) {
          const column = change.start + offset;
          const restoreWidth = root.layout.columnAxis.isColumnHidden(column)
            ? root.layout.columnAxis.getColumnRestoreWidth(column)
            : undefined;
          const width = root.layout.columnAxis.getColumnRestoreWidth(column);
          widths.push(width);
          if (restoreWidth !== undefined) {
            hiddenOffsets.set(offset, restoreWidth);
          }
        }
        inverseColumnWidths = Object.freeze(widths);
        inverseHiddenColumnOffsets = hiddenOffsets;
        columns = Object.freeze(root.columns.slice(0, change.start).concat(root.columns.slice(change.start + change.count)));
      } else {
        columns = moveColumns(root.columns, change.start, change.count, change.target as number);
      }
      runtimeColumns = this.#columnManager.prepareRoot(columns);
      headers = this.#headerManager.prepareRoot(columns, this.#headerGroups);
    }

    const layout = this.#layout.prepareStructureRoot(change, layoutPayload, root.layout);
    const rowCount = change.axis === "row" ? layout.rowCount : root.selection.bounds.rowCount;
    const columnCount = change.axis === "column" ? layout.columnCount : root.selection.bounds.columnCount;
    const selection = this.#selection.prepareBoundsRoot(
      { rowCount, columnCount },
      change.axis === "row" ? structuralChange : undefined,
      change.axis === "column" ? structuralChange : undefined,
      root.selection
    );

    const nextRoot = Object.freeze({ columns, runtimeColumns, headers, layout, selection });
    context.roots.set(this.#layout, layout);
    return {
      nextRoot,
      forwardPayload,
      inverseChange: inverseChange(change),
      inversePayload: createGridStructurePayload({
        ...(inverseRowMetricsFragment !== undefined
          ? { rowMetricsFragment: inverseRowMetricsFragment }
          : {}),
        ...(inverseRowHeightProvenance !== undefined
          ? { rowHeightProvenance: inverseRowHeightProvenance }
          : {}),
        ...(inverseColumns !== undefined ? { columns: inverseColumns } : {}),
        ...(inverseColumnWidths !== undefined ? { columnWidths: inverseColumnWidths } : {}),
        ...(inverseHiddenColumnOffsets !== undefined
          ? { hiddenColumnOffsets: inverseHiddenColumnOffsets }
          : {})
      }),
      estimatedBytes: 192 + change.count * (change.axis === "column" ? 96 : 8)
    };
  }

  replaceRoot(root: GridStructureRoot): void {
    this.#replaceColumns(root.columns);
    this.#columnManager.replaceRoot(root.runtimeColumns);
    this.#headerManager.replaceRoot(root.headers);
    this.#layout.replaceRoot(root.layout);
    this.#selection.replaceRoot(root.selection);
  }
}

export function createGridStructurePayload(payload: GridStructurePayload = {}): GridStructurePayload {
  return Object.freeze({
    ...(payload.columns !== undefined ? { columns: snapshotColumnDefinitions(payload.columns) } : {}),
    ...(payload.rowHeights !== undefined ? { rowHeights: new Map(payload.rowHeights) } : {}),
    ...(payload.rowMetricsFragment !== undefined
      ? { rowMetricsFragment: payload.rowMetricsFragment }
      : {}),
    ...(payload.rowHeightProvenance !== undefined
      ? { rowHeightProvenance: new Map(payload.rowHeightProvenance) }
      : {}),
    ...(payload.columnWidths !== undefined ? { columnWidths: Object.freeze([...payload.columnWidths]) } : {}),
    ...(payload.hiddenColumnOffsets !== undefined
      ? { hiddenColumnOffsets: new Map(payload.hiddenColumnOffsets) }
      : {})
  });
}
