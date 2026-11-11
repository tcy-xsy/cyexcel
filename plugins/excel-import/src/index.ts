import type { WorkbookHandle } from "cyexcel";
import type { ToolbarItem } from "cyexcel/toolbar";
import { readExcelWorkbook } from "./readExcel";
import type { ExcelImportInput, ExcelImportOptions, ExcelImportResult } from "./readExcel";

export { readExcelWorkbook };
export type { ExcelImportInput, ExcelImportOptions, ExcelImportResult };

export interface ExcelImportPlugin {
  readonly toolbarItems: readonly ToolbarItem[];
  /** Bind after createWorkbook(). One plugin instance belongs to one workbook. */
  attach(workbook: WorkbookHandle): void;
  /** Parse first, then replace all tabs. Rejecting or cancelling keeps existing data. */
  importExcel(input: ExcelImportInput, options?: ExcelImportOptions): Promise<ExcelImportResult>;
  dispose(): void;
}

/** Optional integration using cyexcel's existing custom toolbar API. */
export function createExcelImportPlugin(defaults: ExcelImportOptions = {}): ExcelImportPlugin {
  let workbook: WorkbookHandle | undefined;
  let generation = 0;
  let busy = false;
  let disposed = false;
  const dialogs = new Set<() => void>();

  const importExcel = async (input: ExcelImportInput, options: ExcelImportOptions = {}): Promise<ExcelImportResult> => {
    if (disposed || !workbook) throw new Error("Attach the Excel import plugin to a workbook before importing");
    if (busy) throw new Error("An Excel import is already in progress");
    const target = workbook;
    const current = generation;
    busy = true;
    try {
      const result = await readExcelWorkbook(input, { ...defaults, ...options });
      options.signal?.throwIfAborted();
      defaults.signal?.throwIfAborted();
      if (disposed || generation !== current) throw new Error("Excel import was cancelled");
      target.replaceSheets(result.sheets);
      return result;
    } finally {
      busy = false;
    }
  };

  const item: ToolbarItem = {
    id: "import-excel", group: "data", order: 49, type: "dialog", label: "导入 Excel",
    tooltip: "导入 .xlsx 文件并替换当前工作簿",
    icon: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 16V3m-4 4 4-4 4 4M5 20h14"/></svg>',
    dialog: (_ctx, body, close) => {
      const controller = new AbortController();
      const form = document.createElement("form");
      form.style.cssText = "display:grid;gap:12px;max-width:360px";
      const description = document.createElement("p");
      description.textContent = "选择 .xlsx 文件。导入会替换全部工作表，清空撤销历史；图片、图表和数据验证规则不导入。";
      description.style.margin = "0";
      const file = document.createElement("input");
      file.type = "file";
      file.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      file.setAttribute("aria-label", "Excel 文件");
      const confirmation = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      confirmation.append(checkbox, document.createTextNode("替换当前工作簿"));
      const status = document.createElement("p");
      status.setAttribute("role", "status");
      status.setAttribute("aria-live", "polite");
      status.style.margin = "0";
      const submit = document.createElement("button");
      submit.type = "submit";
      submit.textContent = "导入并替换";
      submit.disabled = true;
      const refresh = (): void => { submit.disabled = !file.files?.length || !checkbox.checked || busy; };
      file.addEventListener("change", refresh);
      checkbox.addEventListener("change", refresh);
      form.append(description, file, confirmation, status, submit);
      body.append(form);
      let pending = false;
      const onSubmit = async (event: SubmitEvent): Promise<void> => {
        event.preventDefault();
        const selected = file.files?.[0];
        if (!selected || !checkbox.checked || pending) return;
        pending = true;
        submit.disabled = true;
        file.disabled = true;
        checkbox.disabled = true;
        status.textContent = "正在读取 Excel 文件…";
        try {
          if (!/\.xlsx$/i.test(selected.name)) throw new Error("请选择 .xlsx 文件");
          await importExcel(selected, { signal: controller.signal });
          close();
        } catch (error) {
          if (!controller.signal.aborted) status.textContent = `导入失败：${error instanceof Error ? error.message : String(error)}`;
        } finally {
          pending = false;
          file.disabled = false;
          checkbox.disabled = false;
          refresh();
        }
      };
      form.addEventListener("submit", onSubmit);
      const cleanup = (): void => {
        controller.abort();
        form.removeEventListener("submit", onSubmit);
        file.removeEventListener("change", refresh);
        checkbox.removeEventListener("change", refresh);
        form.remove();
        dialogs.delete(close);
      };
      dialogs.add(close);
      return cleanup;
    }
  };
  return {
    toolbarItems: [item],
    attach: target => {
      if (disposed) throw new Error("Excel import plugin is disposed");
      if (workbook && workbook !== target) throw new Error("Create a separate Excel import plugin for each workbook");
      if (typeof target.replaceSheets !== "function") throw new Error("This cyexcel version does not support replaceSheets(); update cyexcel before attaching the plugin");
      workbook = target;
    },
    importExcel,
    dispose: () => {
      disposed = true;
      generation += 1;
      workbook = undefined;
      for (const close of [...dialogs]) close();
    }
  };
}
