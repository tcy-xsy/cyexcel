import type { ColumnManager, RuntimeColumnRoot } from "../cells/ColumnManager";
import { snapshotColumnDefinitions } from "../cells/columnSnapshot";
import type { ColumnDefinition } from "../cells/types";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { HeaderManager, HeaderManagerRoot } from "../headers/HeaderManager";
import type { HeaderGroupDefinition } from "../headers/types";
import type { LayoutRoot, LayoutStructurePayload } from "../layout/LayoutEngine";
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
  /**
   * 没有 `rowHeightProvenance`：行高来源的权威在 `SheetModel` 的 `rowData[row].hp`，
   * 而 `rowData` 是登记在 `IndexedTableRegistry` 里的侧表——插入/删除/移动会连同
   * `h`/`hp` 一起搬走，删除的逆（`restore-rows`）也会把整条 `RowMeta` 带回来。
   * 这边只需要搬**索引**（`RowMetrics`），来源那半边一个字节都不用它管
   * （ADR 0004 阶段 4B T8，设计稿 §4.1）。
   */
  /**
   * 插入列时每一列的**几何宽度**（隐藏列即 0）。
   *
   * 没有 `hiddenColumnOffsets`：列宽与「这列是不是隐藏的」的权威在 `columnData[c]` 的
   * `w`/`hd`，而 `columnData` 与 `rowData` 一样是登记在册的列侧表——删除的逆
   * （`restore-columns`）会把整条 `ColumnMeta` 带回来，隐藏位自然跟着回来。这边只需要
   * 搬索引（`ColumnAxisSnapshot` 的几何宽度），「取消隐藏后恢复成多宽」那半边一个字节
   * 都不用它管（ADR 0004 阶段 4B T8b，设计稿 §4.4）。
   */
  readonly columnWidths?: readonly number[];
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

  /**
   * 一列没有显式宽度时算多宽——转自布局配置。
   *
   * 插入列时要用它填 `columnWidths`：`ColumnDefinition.width` 缺省的那些列在
   * `columnData` 里不会有 `w`，几何宽度就得由缺省项决定；若这里写死 120 而布局配了
   * 别的数，插进来的列当场就违反 §4.2 那条列轴不变量。
   */
  get defaultColumnWidth(): number {
    return this.#layout.defaultColumnWidth;
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
    let inverseColumns: readonly ColumnDefinition[] | undefined;
    let inverseColumnWidths: readonly number[] | undefined;
    let forwardPayload: GridStructurePayload = {};
    let columns = root.columns;
    let runtimeColumns = root.runtimeColumns;
    let headers = root.headers;
    let layoutPayload: LayoutStructurePayload = {};

    if (change.axis === "row") {
      if (change.operation === "insert") {
        layoutPayload = {
          rowHeights: payload.rowHeights,
          rowMetricsFragment: payload.rowMetricsFragment
        };
        forwardPayload = createGridStructurePayload(
          payload.rowMetricsFragment !== undefined
            ? { rowMetricsFragment: payload.rowMetricsFragment }
            : payload.rowHeights !== undefined
              ? { rowHeights: payload.rowHeights }
              : {}
        );
      } else if (change.operation === "delete") {
        inverseRowMetricsFragment = root.layout.rowMetrics.slice(change.start, change.count);
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
        layoutPayload = { columnWidths: payload.columnWidths };
        forwardPayload = createGridStructurePayload({
          columns: inserted,
          columnWidths: payload.columnWidths
        });
      } else if (change.operation === "delete") {
        inverseColumns = root.columns.slice(change.start, change.start + change.count);
        const widths: number[] = [];
        for (let offset = 0; offset < change.count; offset += 1) {
          // 存**几何宽度**而不是恢复宽度：撤销时权威那半边由 `restore-columns` 把整条
          // `ColumnMeta`（含 `hd`）原样带回来，这里只要索引与它对得上。隐藏列的几何宽度
          // 是 0，恢复后重算出来的也是 0，两边自然一致（设计稿 §4.4）。
          widths.push(root.layout.columnAxis.getColumnWidth(change.start + offset));
        }
        inverseColumnWidths = Object.freeze(widths);
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
        ...(inverseColumns !== undefined ? { columns: inverseColumns } : {}),
        ...(inverseColumnWidths !== undefined ? { columnWidths: inverseColumnWidths } : {})
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
    ...(payload.columnWidths !== undefined ? { columnWidths: Object.freeze([...payload.columnWidths]) } : {})
  });
}
