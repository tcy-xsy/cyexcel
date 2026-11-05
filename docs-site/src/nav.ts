import type { PageModule } from "./router";
import overviewPage from "./pages/overview";
import quickStartPage from "./pages/quick-start";
import basicExamplePage from "./pages/examples/basic";
import commandsExamplePage from "./pages/examples/commands";
import stylesExamplePage from "./pages/examples/styles";
import linksAndImagesExamplePage from "./pages/examples/links-and-images";
import searchExamplePage from "./pages/examples/search";
import validationExamplePage from "./pages/examples/validation";
import formulasExamplePage from "./pages/examples/formulas";
import toolbarExamplePage from "./pages/examples/toolbar";
import workbookExamplePage from "./pages/examples/workbook";
import performanceExamplePage from "./pages/examples/performance";
import createSheetApiPage from "./pages/api/create-sheet";
import coreApiPage from "./pages/api/core";
import snapshotApiPage from "./pages/api/snapshot";
import workbookApiPage from "./pages/api/workbook";
import dataStructurePage from "./pages/architecture/data-structure";
import hooksPage from "./pages/extend/hooks";
import pluginsPage from "./pages/extend/plugins";

export interface NavGroup {
  readonly title: string;
  readonly links: readonly {
    readonly label: string;
    readonly route: string;
    readonly module: PageModule;
  }[];
}

export const navGroups: readonly NavGroup[] = [
  {
    title: "开始",
    links: [
      { label: "Overview", route: "/", module: overviewPage },
      { label: "Quick Start", route: "/quick-start", module: quickStartPage }
    ]
  },
  {
    title: "示例",
    links: [
      { label: "基础表格", route: "/examples/basic", module: basicExamplePage },
      { label: "编辑与历史", route: "/examples/commands", module: commandsExamplePage },
      { label: "样式与格式", route: "/examples/styles", module: stylesExamplePage },
      { label: "链接与图片单元格", route: "/examples/links-and-images", module: linksAndImagesExamplePage },
      { label: "查找与替换", route: "/examples/search", module: searchExamplePage },
      { label: "数据验证", route: "/examples/validation", module: validationExamplePage },
      { label: "公式计算", route: "/examples/formulas", module: formulasExamplePage },
      { label: "工具栏", route: "/examples/toolbar", module: toolbarExamplePage },
      { label: "多 Sheet 工作簿", route: "/examples/workbook", module: workbookExamplePage },
      { label: "十万行性能", route: "/examples/performance", module: performanceExamplePage }
    ]
  },
  {
    title: "API",
    links: [
      { label: "createSheet()", route: "/api/create-sheet", module: createSheetApiPage },
      { label: "CyExcelCore", route: "/api/core", module: coreApiPage },
      { label: "matrixToWorkbookSnapshot()", route: "/api/snapshot", module: snapshotApiPage },
      { label: "createWorkbook()", route: "/api/workbook", module: workbookApiPage }
    ]
  },
  {
    title: "架构与扩展",
    links: [
      { label: "数据结构", route: "/architecture/data-structure", module: dataStructurePage },
      { label: "Hooks", route: "/extend/hooks", module: hooksPage },
      { label: "Plugins", route: "/extend/plugins", module: pluginsPage }
    ]
  }
];
