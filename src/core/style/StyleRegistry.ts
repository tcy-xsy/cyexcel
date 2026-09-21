import { CyExcelError, CyExcelErrorCode } from "../errors";
import { styleKey } from "./styleKey";
import type { CellStyle } from "./types";

/**
 * 6 位 base36 随机串。用 `Math.random` 而不是 `crypto.randomUUID`：这个值不需要
 * 密码学强度，只需要"两次 intern 不撞"，而 `crypto` 在 Worker/SSR 里还要多考虑
 * 一层可用性。
 */
function randomStyleId(): string {
  return `s${Math.random().toString(36).slice(2, 8)}`;
}

/**
 * workbook 级样式表：存储层只有 id，API 层接受对象（ADR 0004 §3.4）。
 *
 * # 三处与 Univer 的 `Styles` 不同
 *
 * 1. **key 规范化**，见 `styleKey`——它那条"cache miss 就遍历整表深比较"的 O(n)
 *    兜底在这里不需要存在。
 * 2. **不接受内联样式对象**。Univer 的 `getStyleByCell` 要处理 `cell.s` 既可能是
 *    id 又可能是整个样式对象的联合类型，正是 ADR 0004 §3.4/§4.4 否掉的那条。
 * 3. **存盘时剔除无人引用的条目**（`toJSON(usedIds)`）。
 *
 * # 为什么不做引用计数
 *
 * 引用计数要在 7 个结构操作、撤销、剪贴板每条路径上都记得加减，漏一处就是静默泄漏
 * 或者过早回收——后者会让一个还在用的 id 指向空，是不会报错的那种坏。"存盘时扫一遍"
 * 是 O(占用格)、只在存盘时发生、不可能算错。代价是长会话的内存里留着废弃样式，
 * 一条几十字节。
 *
 * 随机 id 的代价也落在同一处：它**不幂等**，同一个样式反复 intern（undo/redo 来回
 * 折腾就会）会生成多个 id，表越滚越大——同样由存盘剔除兜底。
 */
export class StyleRegistry {
  readonly #byId = new Map<string, CellStyle>();
  /** 规范化 key → id。intern 的快路径，也是"同一样式只存一条"的实现。 */
  readonly #idByKey = new Map<string, string>();
  readonly #generateId: () => string;

  /**
   * @param styles 已有的样式表（从快照加载）。
   * @param generateId 注入 id 生成器——测试用它构造确定性 id 与碰撞，生产用默认的随机串。
   */
  constructor(styles: Readonly<Record<string, CellStyle>> = {}, generateId: () => string = randomStyleId) {
    this.#generateId = generateId;
    for (const [id, style] of Object.entries(styles)) {
      this.#byId.set(id, style);
      // 后来的同 key 条目不覆盖先来的：加载一份"同一样式存了两个 id"的快照时
      // （随机 id 不幂等，这种快照是可能存在的），两个 id 都保留可用，但之后
      // intern 这个样式一律得到第一个，慢慢收敛到一条。
      const key = styleKey(style);
      if (!this.#idByKey.has(key)) {
        this.#idByKey.set(key, id);
      }
    }
  }

  /** 查 id 对应的样式。id 不存在时返回 `undefined`——不抛，因为快照可能来自外部。 */
  get(id: string): CellStyle | undefined {
    return this.#byId.get(id);
  }

  /**
   * 登记一个样式，返回它的 id。命中已有的就复用，这是"一万个加粗格子只存一条"
   * 的实现。
   */
  intern(style: CellStyle): string {
    const key = styleKey(style);
    const existing = this.#idByKey.get(key);
    if (existing !== undefined) {
      return existing;
    }
    if (key === "") {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidMutation,
        "Cannot intern an empty style: a cell with no style has no `s` field at all"
      );
    }

    let id = this.#generateId();
    // 随机 6 位 base36 撞上的概率极低，但"极低"不是"不会"——注入的生成器更是可能
    // 故意撞。重算一个，而不是覆盖掉别人的样式。
    while (this.#byId.has(id)) {
      id = this.#generateId();
    }
    this.#byId.set(id, style);
    this.#idByKey.set(key, id);
    return id;
  }

  /**
   * 序列化成快照里的 `styles`。
   *
   * 传 `usedIds` 时只保留被引用的条目——这是本类唯一的回收时机（见类注释）。
   * 不传则整表原样导出，给"只想看看现在表里有什么"的调用方用。
   */
  toJSON(usedIds?: Iterable<string>): Record<string, CellStyle> {
    if (usedIds === undefined) {
      return Object.fromEntries(this.#byId);
    }
    const kept: Record<string, CellStyle> = {};
    for (const id of usedIds) {
      const style = this.#byId.get(id);
      if (style !== undefined) {
        kept[id] = style;
      }
    }
    return kept;
  }

  /** 表里现有多少条。playground 的调试信息与稀疏性测试用。 */
  get size(): number {
    return this.#byId.size;
  }
}
