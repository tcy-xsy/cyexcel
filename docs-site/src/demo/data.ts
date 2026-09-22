import { matrixToWorkbookSnapshot, type WorkbookSnapshot } from "@cyexcel/core";

export const columns = [
  { width: 130, header: "SKU", sortable: false, filterable: false },
  { width: 220, header: "名称", wrap: true, sortable: false, filterable: false },
  { width: 110, header: "区域", sortable: false, filterable: false },
  { width: 100, header: "价格", sortable: false, filterable: false },
  { width: 90, header: "库存", sortable: false, filterable: false }
];

export const rows = [
  ["CY-100001", "工业温度传感器", "East", 128, "有"],
  ["CY-100002", "便携标签打印机", "North", 465, "有"],
  ["CY-100003", "空气质量探头", "East", 209.5, "无"],
  ["CY-100004", "仓储扫码终端", "South", 780, "有"],
  ["CY-100005", "低功耗控制器", "North", 356, "有"],
  ["CY-100006", "车载充电底座", "West", 188.8, "无"],
  ["CY-100007", "智能门锁套件", "South", 929, "有"],
  ["CY-100008", "户外电源模组", "East", 612.4, "有"]
];

export const snapshot: WorkbookSnapshot = matrixToWorkbookSnapshot(rows);

export const summaryColumns = [
  { width: 120, header: "区域", sortable: false, filterable: false },
  { width: 110, header: "订单数", sortable: false, filterable: false },
  { width: 130, header: "销售额", sortable: false, filterable: false },
  { width: 120, header: "客单价", sortable: false, filterable: false }
];

export const summaryRows = [
  ["华东", 3, 949.9, 316.63],
  ["华北", 1, 465, 465],
  ["华南", 0, 0, 0],
  ["西南", 1, 780, 780],
  ["华中", 1, 188.8, 188.8],
  ["西北", 1, 929, 929],
  ["东北", 1, 612.4, 612.4]
];

export const summarySnapshot: WorkbookSnapshot = matrixToWorkbookSnapshot(summaryRows);
