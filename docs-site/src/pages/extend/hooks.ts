import type { DocElement } from "../../components/content";
import { createSheet } from "@cyexcel/createSheet";
import { columns, snapshot } from "../../demo/data";
import {
  setText,
  actionButton,
  apiTable,
  callout,
  codeBlock,
  demoBlock,
  pageHeader,
  section
} from "../../components/doc";

export default function hooksPage(root: DocElement): void {
  let sheet: ReturnType<typeof createSheet>;
  root.append(
    pageHeader(
      "EXTEND",
      "Hooks",
      "Hook 参与操作流程，可以观察、转换或取消一次提交；Event 只能观察已经完成的结果。"
    )
  );

  const modes = section("执行模式");
  modes.body.append(
    codeBlock(`sheet.hooks.tap(
  "beforeCommand",
  { name: "audit" },
  (context) => {
    const { command } = context as { command: { type: string } };
    if (command.type === "delete-rows") {
      context.cancel?.("文档示例禁止删除行");
    }
  }
);`)
  );
  root.append(modes.root);
  root.append(
    apiTable([
      { name: "series", type: "mode", description: "按顺序执行所有 tap。" },
      { name: "bail", type: "mode", description: "按顺序执行，某个 tap 取消后停止。" },
      { name: "waterfall", type: "mode", description: "上一个 tap 的返回值是下一个 tap 的输入。" },
      { name: "parallel", type: "mode", description: "tap 之间不依赖彼此副作用。" }
    ])
  );

  const signature = section("注册签名");
  signature.body.append(
    codeBlock(`const dispose = sheet.hooks.tap(
  "beforeCommand",
  { name: "audit", priority: 10 },
  (context) => {
    // series/bail: handler 只处理 context，返回值被忽略。
    // waterfall: handler 的第二个参数是 context，返回值传给下一个 tap。
  }
);

// 不需要时手动移除；sheet dispose 也会统一清理。
dispose();`)
  );
  signature.body.append(
    apiTable([
      { name: "hookName", type: "string", description: "注册表中的固定名称；未知名称会抛 unknown-hook。" },
      { name: "meta", type: "{ name: string; priority?: number }", description: "name 用于日志和错误定位；priority 越大越先执行，默认 0。" },
      { name: "handler", type: "context => void 或 (value, context) => TValue", description: "只有 waterfall 使用第二个参数并消费返回值；其他模式返回值忽略。" },
      { name: "返回值", type: "() => void", description: "Disposer，调用后注销这个 tap。" }
    ])
  );
  root.append(signature.root);

  const names = section("当前 Hook");
  root.append(names.root);
  root.append(
    apiTable([
      {
        name: "beforeInit",
        type: "series",
        description: "context: {}。插件 setup 期间、核心完成主要装配前触发。handler 返回 void；抛错会中止初始化。"
      },
      {
        name: "beforeDestroy",
        type: "series",
        description: "context: {}。dispose() 前触发。handler 返回 void；单个 tap 异常被隔离并转发到 onHookError。"
      },
      {
        name: "beforeCommand",
        type: "bail",
        description: "context: { command: { type, params }, transactionId, cancel(reason?), cancelled, reason }。handler 返回 void；调用 cancel() 会阻止 Command 进入计划。"
      },
      {
        name: "beforeTransactionCommit",
        type: "bail",
        description: "context: { commands, patches, model, transactionId, cancel(reason?), cancelled, reason }。handler 返回 void；取消后整个事务不提交。"
      },
      {
        name: "beforeRowsChange",
        type: "bail",
        description: "context: { operation: 'insert' | 'delete' | 'move', start, count, target?, cancel(reason?), cancelled, reason }。move 时有 target。返回 void；cancel() 阻止行结构变更。"
      },
      {
        name: "beforeColumnsChange",
        type: "bail",
        description: "context: { operation: 'insert' | 'delete' | 'move', start, count, target?, cancel(reason?), cancelled, reason }。返回 void；cancel() 阻止列结构变更。"
      },
      {
        name: "resolveColumn",
        type: "waterfall",
        description: "handler: (column: RuntimeColumn, { definition, index }) => RuntimeColumn。必须返回列配置；异常被隔离，当前列回退到上一个值。"
      },
      {
        name: "beforeContextMenuOpen",
        type: "bail",
        description: "context: { context: ContextMenuContext, cancel(reason?), cancelled, reason }。返回 void；cancel() 后菜单不打开。"
      },
      {
        name: "extendContextMenu",
        type: "waterfall",
        description: "handler: (items: ContextMenuItem[], context: ContextMenuContext) => ContextMenuItem[]。返回最终菜单项数组；异常被隔离。"
      },
      {
        name: "onHookError",
        type: "series",
        description: "context: { error: unknown, hookName: string, handlerName: string }。返回 void；自身异常不会再递归进入本 Hook。"
      }
    ])
  );

  const logRef = { current: undefined as DocElement | undefined };
  const setLog = (text: string): void => {
    if (logRef.current) {
      setText(logRef.current, text);
    }
  };
  const setCell = actionButton("写入单元格", () => {
    sheet.setCell(0, 1, "Hook 已观察这次写入");
  });
  const insertRow = actionButton("插入一行", () => {
    sheet.insertRows(sheet.getUsedRowCount(), 1);
  });

  const demo = demoBlock({
    title: "beforeCommand 审计",
    description: "每次写入或插入行均先触发 beforeCommand。",
    actions: [setCell, insertRow]
  });
  root.append(demo.root);
  if (demo.log === undefined) {
    throw new Error("demo log missing");
  }
  logRef.current = demo.log;

  demo.mount.addClass("with-toolbar");
  demo.mount.onMount = () => {
    sheet = createSheet(demo.mount.element, {
      columns,
      snapshot,
      headerHeight: 28,
      rowHeaderWidth: 48,
      rowHeight: 30
    });
    sheet.hooks.tap("beforeCommand", { name: "docs-audit" }, (context: unknown) => {
      const typed = context as { command: { type: string } };
      setLog(`beforeCommand：${typed.command.type}`);
    });
    return () => sheet.dispose();
  };

  root.append(callout("当前注册表中的 Hook 均为同步 Hook；tapAsync() 会抛 async-hook-not-supported。bail Hook 通过 context.cancel() 取消，waterfall Hook 必须返回转换后的值。Event 只能观察已完成结果，不会回滚或改写。"));
}
