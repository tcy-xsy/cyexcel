import { CyExcelError, CyExcelErrorCode } from "../errors";
import { assertValidResourceEntry, resourceToTable, tableToResource } from "./resources";
import { assertValidWorkbookSnapshot } from "./serialize";
import { SheetModel, type SheetModelOptions } from "./SheetModel";
import type { IndexedTable } from "./IndexedTable";
import type { ResourceEntry, WorkbookSnapshot } from "./types";
import { SNAPSHOT_VERSION } from "./types";

export interface WorkbookModelOptions {
  readonly sheet?: SheetModelOptions;
}

/**
 * 文档根:`sheets` + `styles` + `resources`(spec §3 顶层结构)。
 *
 * **资源在 workbook 级,重键在 sheet 级。** 规格把 `resources` 与 `styles`、
 * `sheets` 并列放在顶层,所以它的家在这里;但重键是行列结构操作的一部分,
 * 发生在 sheet 上。因此按下标索引的资源在这里**登记进那张 sheet 的登记处**,
 * 之后由 `SheetModel` 的结构操作自动搬运(spec §7.2:核心统一提供,不外包给
 * 插件作者);`indexedBy: "none"` 的留在这里当不透明 blob。
 *
 * 现在 `sheetOrder.length === 1`(不变量 6),"哪张 sheet"没有歧义。M3 放开
 * 多 Sheet 时按下标索引的资源需要带 sheet id——那时再改这里。
 *
 * `styles` 本阶段是**不透明透传**:字段集属于 Step 24(spec §10),这里只保证往返。
 */
export class WorkbookModel {
  readonly #sheet: SheetModel;
  /** 按下标索引的资源:登记在 sheet 上,这里只留引用以便 toSnapshot 收拢 */
  readonly #indexedResources = new Map<string, IndexedTable<unknown>>();
  /** indexedBy: "none",核心不碰 */
  readonly #opaqueResources = new Map<string, ResourceEntry>();
  /**
   * 资源加入的顺序,横跨两个 Map。`toSnapshot()` 的 `resources` 数组按这个顺序
   * 输出,而不是先输出全部 indexed 再输出全部 opaque——否则同一份快照
   * 交叉排列两种资源时,`fromSnapshot(s).toSnapshot().resources` 的顺序会
   * 变成按种类分组,不再等于 `s.resources`,round trip 就不成立了。
   */
  readonly #resourceOrder: string[] = [];
  /** 本阶段不透明透传,没有 setter,故 readonly(字段集属 Step 24,spec §10) */
  readonly #styles: Readonly<Record<string, unknown>>;

  private constructor(sheet: SheetModel, styles: Readonly<Record<string, unknown>>) {
    this.#sheet = sheet;
    this.#styles = styles;
  }

  static createEmpty(options: WorkbookModelOptions = {}): WorkbookModel {
    return new WorkbookModel(SheetModel.createEmpty(options.sheet), {});
  }

  static fromSnapshot(snapshot: WorkbookSnapshot): WorkbookModel {
    assertValidWorkbookSnapshot(snapshot);
    const id = snapshot.sheetOrder[0] as string;
    const sheetSnapshot = snapshot.sheets[id];
    if (!sheetSnapshot) {
      // assertValidWorkbookSnapshot 已经保证了这一条,走到这里说明校验之后被改过
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `sheets is missing "${id}" despite passing snapshot validation`
      );
    }
    const workbook = new WorkbookModel(SheetModel.fromSnapshot(sheetSnapshot), {
      ...snapshot.styles
    });
    for (const entry of snapshot.resources) {
      workbook.addResource(entry);
    }
    return workbook;
  }

  sheet(): SheetModel {
    return this.#sheet;
  }

  sheetIds(): readonly string[] {
    return [this.#sheet.id()];
  }

  /**
   * 加一条插件资源。`indexedBy` 为 `"row"` / `"column"` 的会被登记进 sheet 的
   * 侧表登记处,此后 7 个结构操作自动重键它。
   */
  addResource(entry: ResourceEntry): void {
    assertValidResourceEntry(entry, `resource "${entry.name}"`);
    if (this.#indexedResources.has(entry.name) || this.#opaqueResources.has(entry.name)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `a resource named "${entry.name}" already exists`
      );
    }
    if (entry.indexedBy === "none") {
      this.#opaqueResources.set(entry.name, {
        name: entry.name,
        indexedBy: "none",
        data: entry.data
      });
      this.#resourceOrder.push(entry.name);
      return;
    }
    // 重名(含撞上核心的 rowData / columnData)由登记处拒绝——registerTable 可能抛,
    // 所以 push 必须在它之后,失败的一次不该在顺序表里留下痕迹
    const table = this.#sheet.registerTable(resourceToTable(entry));
    this.#indexedResources.set(entry.name, table);
    this.#resourceOrder.push(entry.name);
  }

  resource(name: string): ResourceEntry | undefined {
    const table = this.#indexedResources.get(name);
    if (table) {
      return tableToResource(table);
    }
    return this.#opaqueResources.get(name);
  }

  resourceNames(): readonly string[] {
    return [...this.#resourceOrder];
  }

  toSnapshot(): WorkbookSnapshot {
    const sheetSnapshot = this.#sheet.toSnapshot();
    const resources: ResourceEntry[] = this.#resourceOrder.map((name) => {
      const entry = this.resource(name);
      if (!entry) {
        // #resourceOrder 只在 addResource 成功记录之后才追加,查不到对应条目
        // 说明两处状态被绕过登记处直接改过,是编程错误,必须炸出来
        throw new CyExcelError(
          CyExcelErrorCode.InvalidOptions,
          `resource "${name}" is in the insertion order but has no entry in either resource map`
        );
      }
      return entry;
    });
    return {
      version: SNAPSHOT_VERSION,
      sheetOrder: [sheetSnapshot.id],
      sheets: { [sheetSnapshot.id]: sheetSnapshot },
      styles: { ...this.#styles },
      resources
    };
  }
}
