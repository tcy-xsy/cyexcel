/**
 * Type-checked (never executed, never imported by a test runner) against
 * the package's own `exports` map, resolving the built declarations just
 * as an npm consumer would. If this fails to compile, the shipped `.d.ts` files don't
 * actually match what `exports`/the built JS provide. Same
 * `@ts-expect-error`-assertion style as `tests/types/core-types.ts`, which
 * checks the source (not built) types instead.
 */
import { CyExcelCore, createSheet, DEFAULT_CYEXCEL_TAG_NAME, rowsToWorkbookSnapshot } from "cyexcel";
import type { CyExcelCoreOptions } from "cyexcel";
import * as CoreApi from "cyexcel/core";
import { defineCyExcel } from "cyexcel/register";
import { createFormulaEngine } from "cyexcel/formula";
import { createToolbar } from "cyexcel/toolbar";

// These must remain typed even when consumers enable skipLibCheck.
// @ts-expect-error A sheet factory is not a symbol.
const sheetFactorySymbol: symbol = createSheet;
// @ts-expect-error A core factory is not a symbol.
const coreFactorySymbol: symbol = CoreApi.createCyExcelCore;
// @ts-expect-error A registration function is not a symbol.
const registerSymbol: symbol = defineCyExcel;
// @ts-expect-error A formula factory is not a symbol.
const formulaFactorySymbol: symbol = createFormulaEngine;
// @ts-expect-error A toolbar factory is not a symbol.
const toolbarFactorySymbol: symbol = createToolbar;
void [sheetFactorySymbol, coreFactorySymbol, registerSymbol, formulaFactorySymbol, toolbarFactorySymbol];

// --- "cyexcel" re-exports the same CyExcelCore as "cyexcel/core" --------

const sameClass: typeof CyExcelCore = CoreApi.CyExcelCore;
void sameClass;

type IsCoreExported<Name extends PropertyKey> = Name extends keyof typeof CoreApi ? true : false;
const internalRuntimeIsNotShipped: [
  IsCoreExported<"TableStore">,
  IsCoreExported<"MutationCoordinator">,
  IsCoreExported<"LayoutEngine">,
  IsCoreExported<"GridStructureOwner">,
  IsCoreExported<"createGridCommandManager">,
  IsCoreExported<"CommandManager">,
  IsCoreExported<"SelectionModel">,
  IsCoreExported<"SurfaceHost">,
  IsCoreExported<"RenderScheduler">,
  IsCoreExported<"CyExcelOptions">,
  IsCoreExported<"assertValidCyExcelOptions">
] = [false, false, false, false, false, false, false, false, false, false, false];
void internalRuntimeIsNotShipped;

// --- CyExcelCoreOptions: fixed shape, columns/snapshot optional ----------
//
// 顺带钉住「根入口自己就能装数据」：`rowsToWorkbookSnapshot` 必须从 "cyexcel"
// 出得来，否则只 import `CyExcelCore` 的消费者只能开一张空表（见 src/index.ts）。

const options: CyExcelCoreOptions = {
  // 列配置里没有 id、也没有 field（ADR 0004 §3.8.2）：字段名只在下面那次导入里
  // 出现一次，列本身就是它的下标。
  columns: [{ width: 120 }],
  snapshot: rowsToWorkbookSnapshot([{ name: "Ada" }], ["name"]),
  rowHeight: 24
};
void options;

// --- CyExcelCore actually constructs from a real HTMLElement -------------

declare const container: HTMLElement;
const sheet = new CyExcelCore(container, options);
sheet.setCell(0, 0, "Grace");
const value: unknown = sheet.getCell(0, 0);
void value;

// --- "cyexcel/register" exposes the registration entry points ------------

const tagName: string = DEFAULT_CYEXCEL_TAG_NAME;
void tagName;
const registered: CustomElementConstructor = defineCyExcel();
void registered;
