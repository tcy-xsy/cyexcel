import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { ModelRevision } from "../model";
import type { TransactionReadView } from "../mutations/types";
import { applySheetMutation, estimateSheetMutationBytes } from "./apply";
import { invertSheetMutation } from "./invert";
import type { SheetMutation } from "./mutation";
import type { SheetModel } from "./SheetModel";
import type { CellRecord, ColumnMeta, RowMeta } from "./types";

/**
 * 一条已暂存、尚未写进活模型的变更。
 *
 * 形状与被替代的 `PreparedDataSourceMutation` 一致（`forward` / `inverse` /
 * `estimatedBytes` / `commit` / `rollback` / `dispose`），因为消费它的是同一个
 * `MutationCoordinator`。差别只在载荷类型：`SheetMutation` 而不是 `DataMutation`。
 */
export interface PreparedSheetMutation {
  readonly forward: SheetMutation;
  readonly inverse: SheetMutation;
  readonly estimatedBytes: number;
  commit(): void;
  rollback(): void;
  dispose(): void;
}

/**
 * 一次事务的暂存层：开始时克隆一份**影子模型**，暂存写影子，commit 才写活模型。
 *
 * ## 协议（设计稿 §2.2 方案 A）
 *
 * ```
 * begin:    shadow = live.clone()
 * prepare:  inv = invertSheetMutation(shadow, m)   // 逆在「前面几条已生效」的状态上求
 *           applySheetMutation(shadow, m)          // 预览随之前进
 * commit:   applySheetMutation(live, m)
 * rollback: applySheetMutation(live, inv)
 * ```
 *
 * ## 为什么是克隆，不是覆盖层
 *
 * **决定性理由不是性能，是这条分支已经连续三次栽在同一个坑里：同一套语义两份实现，
 * 两份就会分叉**（阶段二的侧表 extent、阶段三的 `reorderRows` 校验顺序、阶段 4A 的
 * 合法 `order`）。覆盖层方案要求把 7 个结构操作的语义在暂存层里**再实现一遍**，而
 * 那一遍必须与 `SheetModel` 逐位一致——正是上面三个缺陷的生成条件。克隆方案在结构
 * 上排除了它：影子就是一个 `SheetModel`，跑的是与活模型**同一份**
 * `insertRows` / `applyRowPlan` / `IndexedTableRegistry.applyRowPlan` 代码，本文件里
 * 一行新的结构代数都没有。
 *
 * 克隆的代价是 O(占用行数) 而不是 O(格子数)——行桶按引用共享、写时才私有化，见
 * `CellMatrix.clone`。
 *
 * ## 它满足 `MutationCoordinator` 的三条要求
 *
 * 1. **prepare 不得触碰活状态。** 本类的 `prepareMutation` 只写影子；影子与活模型的
 *    容器已经分离（桶级 COW），`invertSheetMutation` 本身是只读的。
 * 2. **预览每暂存一条就前进。** `getCell` / `getRowCount` 读的是影子，而影子已经吃过
 *    前面每一条变更——于是第 k 条的逆是在「前 k-1 条已生效」的状态上求出来的，
 *    组合因此成立。
 * 3. **commit 逐个写、失败按相反顺序回滚，所以逆必须精确。** 参与者按顺序 commit，
 *    第 k 个 commit 时活模型 ≡ 影子在第 k 条变更之前的状态；回滚按相反顺序逐条抵消，
 *    回到基态。
 *
 * ## 别名规则
 *
 * 逆变更里的 `CellRecord` / `RowMeta` 别名的是**影子**里那些对象，而影子与活模型按
 * 引用共享它们（`types.ts` 顶部的浅共享规则）。所以历史栈里的记录在事务 `dispose()`
 * 之后依然有效——`dispose()` 不销毁任何值，它只是关掉这个句柄。
 */
export class SheetTransaction implements TransactionReadView {
  readonly baseRevision: ModelRevision;
  readonly #live: SheetModel;
  readonly #shadow: SheetModel;
  #disposed = false;

  constructor(model: SheetModel, baseRevision: ModelRevision) {
    this.#live = model;
    this.#shadow = model.clone();
    this.baseRevision = baseRevision;
  }

  /**
   * 预览的坐标空间。
   *
   * 与 `getCell` 不同，这两个**不查是否已 dispose**——沿用被替代的
   * `ArrayDataSource.beginTransaction` 的口径。理由是它们答的不是一个会过期的数：
   * 影子在 `dispose()` 之后仍然是「这次事务规划完的那个状态」，而 commit 之后活模型
   * 恰好就是那个状态，两者不会给出互相矛盾的答案。
   */
  getRowCount(): number {
    return this.#shadow.rowCount();
  }

  getColumnCount(): number {
    return this.#shadow.columnCount();
  }

  /** 裸值（ADR 0004 §3.6）——`TransactionReadView` 要的正是这个，不需要适配。 */
  getCell(row: number, column: number): unknown {
    this.#assertActive();
    return this.#shadow.getCell(row, column);
  }

  /**
   * 带样式的完整记录。
   *
   * `TransactionReadView` 只要裸值，但**变更层**需要记录：`set-cell` 载的是
   * `CellRecord`，一条只改值的变更要保留该格已有的 `t`/`s`/`f`/`si`，就得先读到它。
   * 这是在读视图之外**多开**的一个方法，不是把 `getCell` 改掉——那个接口有别的实现方。
   */
  getRecord(row: number, column: number): CellRecord | undefined {
    this.#assertActive();
    return this.#shadow.getRecord(row, column);
  }

  /**
   * 影子上的行 / 列元数据。
   *
   * 与 `getRecord` 同一条理由、同一个位置：这两个也是**变更层**要的读，不是读视图
   * 的一部分。`set-row-meta` 载的是整条 `RowMeta`，所以一次只改行高的变更要保住这一
   * 行已有的 `hd`，就得先读到它；而**必须读影子**——同一个事务里先隐藏再改高的两条
   * 变更，第二条若读活模型就会拿到没被第一条更新过的旧条目，把隐藏悄悄盖掉。
   *
   * 这也是 ADR 0004 阶段 4B T8 之后「行高的权威在模型」能成立的前提：布局层的
   * `prepare*` 只算索引，权威那一半由这里读出来、再作为 `set-row-meta` 写回去，
   * 两半在同一个事务里（设计稿 §4.3）。
   */
  getRowMeta(row: number): RowMeta | undefined {
    this.#assertActive();
    return this.#shadow.getRowMeta(row);
  }

  getColumnMeta(column: number): ColumnMeta | undefined {
    this.#assertActive();
    return this.#shadow.getColumnMeta(column);
  }

  /**
   * 暂存一条变更：在影子上求逆、把它施加到影子上，返回一个可 commit / rollback 的句柄。
   *
   * **求逆在施加之前**——逆要读的正是这条变更将要覆盖掉的那部分状态
   * （`invert.ts` 顶部的契约）。
   *
   * **校验也发生在这里**，因为 `invertSheetMutation` 头一件事就是
   * `assertMutationIsAppliable`：参数、合并区、`restore-*` 的 `tables` 三条不变量，
   * 全在 prepare 期查完。这正是 `MutationCoordinator` 要的「能查出来的都在 prepare
   * 查出来，commit 因此不会失败」。
   *
   * 万一 `applySheetMutation(shadow, …)` 仍然抛了（按上面这条不该发生），影子会停在
   * 半路——**这不影响正确性**：活模型一个字节都没动，而抛出会让 `stage()` 走
   * `discardPlanning()`，整个规划会话连同这个影子一起作废。
   */
  prepareMutation(mutation: SheetMutation): PreparedSheetMutation {
    this.#assertActive();
    const inverse = invertSheetMutation(this.#shadow, mutation);
    applySheetMutation(this.#shadow, mutation);

    const live = this.#live;
    let commitStarted = false;
    let rolledBack = false;
    return {
      forward: mutation,
      inverse,
      // 正逆两条都会进历史栈（`CommitReceipt` 同时带着 `forward` 与 `inverse`），
      // 所以估的是两条之和。删除类变更这个和是不对称的：`delete-rows` 本身只有几个
      // 下标，它的逆 `restore-rows` 却背着整段内容——真正占内存的是后者。
      estimatedBytes: estimateSheetMutationBytes(mutation) + estimateSheetMutationBytes(inverse),
      commit: () => {
        // **先置位再施加**，与 `MutationCoordinator` 那边 `started.push` 在
        // `participant.commit()` 之前是同一条理由：施加到一半抛出的变更也必须被回滚。
        commitStarted = true;
        applySheetMutation(live, mutation);
      },
      rollback: () => {
        // 没 commit 过就没有什么要撤的；撤过一次也不再撤第二次——
        // `PreparedMutationParticipant` 的原话是 rollback 必须幂等。
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        applySheetMutation(live, inverse);
      },
      dispose: () => {}
    };
  }

  /**
   * 关掉这个句柄。
   *
   * 没有要释放的资源——影子随事务对象一起被回收，逆变更里别名的那些值归活模型所有
   * （见类注释的「别名规则」）。它存在是为了满足 `MutationPlanningResources.dispose`
   * 的形状，并让 dispose 之后的读被拦下来而不是悄悄给一个过期答案。
   */
  dispose(): void {
    this.#disposed = true;
  }

  #assertActive(): void {
    if (this.#disposed) {
      throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Sheet transaction is disposed");
    }
  }
}
