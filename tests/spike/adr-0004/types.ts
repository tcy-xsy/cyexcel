/**
 * ADR 0004 Spike 的最小单元格记录。字段取 Univer `ICellData` 中与本次验证
 * 相关的三个（值 / 类型标记 / 样式 id），不是完整形态——Spike 只需要"格子是
 * 一条记录而非裸值"这个性质，字段越多内存数字越偏向记录模型，反而不诚实��
 */
export interface CellRecord {
  /** 值；有公式时存计算结果 */
  v: unknown;
  /** 1=string 2=number 3=boolean 4=force text，未设置由引擎推断 */
  t?: number;
  /** 样式 id，指向 workbook 级 styles map（ADR §3.4：存储层只存 id） */
  s?: string;
}

export interface CellMatrix {
  readonly rowCount: number;
  readonly columnCount: number;
  /** 未占用的格子返回 undefined，不抛错 */
  get(row: number, column: number): CellRecord | undefined;
  set(row: number, column: number, record: CellRecord): void;
  delete(row: number, column: number): void;
  /** 已占用格子的总数，用于内��测量时确认填充规模 */
  occupiedCount(): number;
  /** order[newPosition] = oldPosition，仓库既有约定（ADR 0003） */
  reorderRows(order: Int32Array): void;
}
