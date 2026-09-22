import { rowsToWorkbookSnapshot, type WorkbookSnapshot } from "@cyexcel/core";

export const performanceColumns = [
  { width: 120, header: "编号", sortable: false, filterable: false },
  { width: 220, header: "名称", sortable: false, filterable: false },
  { width: 120, header: "区域", sortable: false, filterable: false },
  { width: 100, header: "价格", sortable: false, filterable: false },
  { width: 90, header: "库存", sortable: false, filterable: false }
];

export const performanceFields = ["id", "name", "region", "price", "stock"];

const NAMES = [
  "工业温度传感器",
  "便携标签打印机",
  "空气质量探头",
  "仓储扫码终端",
  "低功耗控制器",
  "车载充电底座"
];
const REGIONS = ["华东", "华南", "华北", "西南", "东北"];

export function buildPerformanceRows(count: number) {
  return Array.from({ length: count }, (_unused, index) => ({
    id: `P-${String(index + 1).padStart(7, "0")}`,
    name: NAMES[index % NAMES.length]!,
    region: REGIONS[index % REGIONS.length]!,
    price: Math.round((48 + ((index * 37) % 520)) * 100) / 100,
    stock: index % 5 === 0 ? "无" : "有"
  }));
}

export function buildPerformanceSnapshot(count: number): WorkbookSnapshot {
  return rowsToWorkbookSnapshot(buildPerformanceRows(count), performanceFields);
}
