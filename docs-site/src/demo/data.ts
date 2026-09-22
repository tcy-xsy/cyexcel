import { rowsToWorkbookSnapshot, type WorkbookSnapshot } from "@cyexcel/core";

export const columns = [
  { width: 130, header: "SKU", sortable: false, filterable: false },
  { width: 220, header: "名称", wrap: true, sortable: false, filterable: false },
  { width: 110, header: "区域", sortable: false, filterable: false },
  { width: 100, header: "价格", sortable: false, filterable: false },
  { width: 90, header: "库存", sortable: false, filterable: false }
];

export const fields = ["sku", "name", "region", "price", "stock"];

export const rows = [
  { sku: "CY-100001", name: "工业温度传感器", region: "East", price: 128, stock: "有" },
  { sku: "CY-100002", name: "便携标签打印机", region: "North", price: 465, stock: "有" },
  { sku: "CY-100003", name: "空气质量探头", region: "East", price: 209.5, stock: "无" },
  { sku: "CY-100004", name: "仓储扫码终端", region: "South", price: 780, stock: "有" },
  { sku: "CY-100005", name: "低功耗控制器", region: "North", price: 356, stock: "有" },
  { sku: "CY-100006", name: "车载充电底座", region: "West", price: 188.8, stock: "无" },
  { sku: "CY-100007", name: "智能门锁套件", region: "South", price: 929, stock: "有" },
  { sku: "CY-100008", name: "户外电源模组", region: "East", price: 612.4, stock: "有" }
];

export const snapshot: WorkbookSnapshot = rowsToWorkbookSnapshot(rows, fields);

export const summaryColumns = [
  { width: 120, header: "区域", sortable: false, filterable: false },
  { width: 110, header: "订单数", sortable: false, filterable: false },
  { width: 130, header: "销售额", sortable: false, filterable: false },
  { width: 120, header: "客单价", sortable: false, filterable: false }
];

export const summaryFields = ["region", "orders", "amount", "average"];

export const summaryRows = [
  { region: "华东", orders: 3, amount: 949.9, average: 316.63 },
  { region: "华北", orders: 1, amount: 465, average: 465 },
  { region: "华南", orders: 0, amount: 0, average: 0 },
  { region: "西南", orders: 1, amount: 780, average: 780 },
  { region: "华中", orders: 1, amount: 188.8, average: 188.8 },
  { region: "西北", orders: 1, amount: 929, average: 929 },
  { region: "东北", orders: 1, amount: 612.4, average: 612.4 }
];

export const summarySnapshot: WorkbookSnapshot = rowsToWorkbookSnapshot(
  summaryRows,
  summaryFields
);
