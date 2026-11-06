import type { CyExcelPlugin } from "@cyexcel/core";
import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import {
  actionButton,
  apiTable,
  callout,
  codeBlock,
  demoBlock,
  pageHeader,
  paragraph,
  section
} from "../../components/doc";

export default function pluginsPage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "EXTEND",
      "Plugins",
      "Plugin 拿到的是受控能力面，不是内部 Store；注销时会自动清理它注册的所有能力。"
    )
  );

  const intro = section("插件形状");
  intro.body.append(
    codeBlock(`const auditPlugin: CyExcelPlugin = {
  name: "audit",
  setup(context) {
    context.events.on("change", (payload) => {
      console.log("change", payload);
    });

    context.registerShortcut("Mod+Shift+L", () => {
      console.log("shortcut from plugin");
    });

    return () => {
      console.log("audit plugin disposed");
    };
  }
};

const sheet = createSheet(container, {
  columns,
  snapshot,
  plugins: [auditPlugin]
});`)
  );
  root.append(intro.root);

  const install = section("怎么添加插件");
  install.body.append(
    paragraph("插件只能在 CyExcelCore 构造时通过 plugins 数组安装；当前公开 API 不支持构造后再动态 use()。"),
    codeBlock(`import { CyExcelCore, type CyExcelPlugin } from "cyexcel/core";

const helloPlugin: CyExcelPlugin = {
  name: "hello",
  setup(context) {
    // 通过 context 注册的事件/hook/快捷键会在插件卸载时自动清理。
    const disposeChange = context.events.on("change", () => {
      console.log("committed change");
    });

    const disposeShortcut = context.registerShortcut("Mod+Shift+H", () => {
      console.log("hello from plugin");
    });

    // 可选：释放 setup 创建的资源。
    return () => {
      disposeChange();
      disposeShortcut();
    };
  }
};

    const sheet = createSheet(container, {
  columns,
  snapshot,
  plugins: [helloPlugin]
});`)
  );
  install.body.append(
    apiTable([
      { name: "name", type: "string", description: "插件唯一名称，用于诊断和归属注册项。" },
      { name: "setup", type: "(context: PluginContext) => void | Disposer", description: "安装时调用一次；返回 Disposer 时会在插件清理阶段额外调用。" },
      { name: "plugins", type: "readonly CyExcelPlugin[]", description: "CyExcelCoreOptions 的构造参数，按数组顺序安装。" }
    ])
  );
  root.append(install.root);

  const capabilities = section("PluginContext 能力");
  root.append(capabilities.root);
  root.append(
    apiTable([
      { name: "events.on()", type: "(event, listener) => Disposer", description: "观察已完成提交、选区变化、错误等事件。" },
      { name: "hooks.tap()", type: "(hookName, meta, handler) => Disposer", description: "参与命令或结构操作前的流程。" },
      { name: "hooks.tapAsync()", type: "(hookName, meta, handler) => Disposer", description: "注册异步 tap；当前内置 Hook 都是同步，调用会抛 async-hook-not-supported。" },
      { name: "contextMenu.registerProvider()", type: "(meta, provider) => Disposer", description: "给右键菜单追加菜单项。" },
      { name: "getCell()", type: "(row, column) => unknown", description: "只读访问单元格。" },
      { name: "copy() / paste()", type: "Promise", description: "使用核心剪贴板能力。" },
      { name: "insertRows(start, count?)", type: "void", description: "通过公开命令接口插入空行，count 默认 1。" },
      { name: "deleteRows(start, count?)", type: "void", description: "通过公开命令接口删除行，count 默认 1。" },
      { name: "moveRows(start, count, target)", type: "void", description: "通过公开命令接口移动行。" },
      { name: "insertColumns(start, columns)", type: "void", description: "插入一列或多列；每项都是完整 ColumnDefinition。" },
      { name: "deleteColumns(start, count?)", type: "void", description: "通过公开命令接口删除列，count 默认 1。" },
      { name: "moveColumns(start, count, target)", type: "void", description: "通过公开命令接口移动列。" },
      { name: "hideColumn(column)", type: "void", description: "隐藏指定下标的列。" },
      { name: "registerShortcut()", type: "(combo, handler) => Disposer", description: "注册键盘快捷键。" }
    ])
  );

  const logRef = { current: undefined as HTMLDivElement | undefined };
  const setLog = (text: string): void => {
    if (logRef.current) {
      logRef.current.textContent = text;
    }
  };
  let changeCount = 0;
  const auditPlugin: CyExcelPlugin = {
    name: "docs-audit",
    setup(context) {
      context.events.on("change", () => {
        changeCount += 1;
        setLog(`audit plugin：已捕获 ${changeCount} 次提交`);
      });
      return () => {
        setLog("audit plugin：已销毁");
      };
    }
  };

  const setCell = actionButton("写入单元格", () => {
    sheet.setCell(0, 1, `插件观察 #${changeCount + 1}`);
  });
  const insertRow = actionButton("插入一行", () => {
    sheet.insertRows(sheet.getUsedRowCount(), 1);
  });

  const demo = demoBlock({
    title: "插件事件审计",
    description: "每次写入或插入行均提交一次 Command，并触发 change 事件。",
    actions: [setCell, insertRow]
  });
  root.append(demo.root);
  if (demo.log === undefined) {
    throw new Error("demo log missing");
  }
  logRef.current = demo.log;

  demo.mount.classList.add("with-toolbar");
  const sheet = createSheet(demo.mount, {
    columns,
    snapshot,
    plugins: [auditPlugin],
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 30
  });

  root.append(callout("PluginManager 会在插件卸载时自动清理 events、hooks、菜单和快捷键注册。"));

  return () => sheet.dispose();
}
