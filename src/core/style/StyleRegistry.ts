import { CyExcelError, CyExcelErrorCode } from "../errors";
import { styleKey } from "./styleKey";
import type { CellStyle } from "./types";

/** 6 位 base36 随机串。不用 crypto：只需要"两次 intern 不撞"，且它在 Worker/SSR 下要多一层可用性判断。 */
function randomStyleId(): string {
  return `s${Math.random().toString(36).slice(2, 8)}`;
}

/**
 * workbook 级样式表：存储层只有 id，API 层接受对象（ADR 0004 §3.4）。
 *
 * **不做引用计数**，改为存盘时剔除无人引用的条目（`toJSON(usedIds)`）。引用计数要在
 * 7 个结构操作、撤销、剪贴板每条路径上都记得加减，漏一处就是静默泄漏或过早回收；
 * 存盘扫一遍是 O(占用格)、只在存盘时发生、不可能算错。随机 id 不幂等带来的重复条目
 * 也由它兜住。
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
