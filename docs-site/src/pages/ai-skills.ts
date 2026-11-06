import type { DocElement } from "../components/content";
import { callout, codeBlock, pageHeader, paragraph, section } from "../components/doc";

export default function aiSkillsPage(root: DocElement): void {
  root.append(pageHeader("GUIDES", "AI 编程助手", "通过 cyexcel Skill，为 Claude Code 和 Codex 提供组件接入、配置与扩展指南。"));

  const install = section("安装 Skill");
  install.body.append(
    paragraph("在需要接入组件的项目目录中执行以下命令，并按提示完成安装。添加 --global 可安装到用户级目录，供多个项目使用。"),
    codeBlock("npx skills add tcy-xsy/cyexcel --skill cyexcel --agent claude-code codex", "shell"),
    paragraph("Skill 源码位于仓库的 skills/cyexcel/，包含 SKILL.md、Codex 展示元数据及按需读取的参考资料。安装不依赖源码仓库中的内部设计文档。")
  );
  root.append(install.root);

  const invoke = section("调用方式");
  invoke.body.append(
    paragraph("安装后重新启动对应助手会话。Claude Code 使用 /cyexcel 调用，Codex 使用 $cyexcel 调用，并在后面描述目标。助手也可根据任务内容匹配此 Skill。"),
    codeBlock("/cyexcel 创建一个采购工作簿，包含数量验证、金额公式和 XLSX 导出。", "text"),
    codeBlock("$cyexcel 在当前 React 页面接入表格，并在组件卸载时释放实例。", "text")
  );
  root.append(invoke.root);

  const scope = section("适用范围");
  scope.body.append(
    paragraph("Skill 包含表格与工作簿创建、快照导入、框架生命周期、公式、数据验证、右键菜单、样式、图片、插件和导出说明。助手应以项目实际安装版本的类型定义为准。"),
    callout("Skill 提供编程指引，不安装组件运行时，也不提供远程控制表格的服务。当前仓库包仍为私有开发版本；未发布 npm 版本时，需构建并打包源码，再将生成的包安装到业务项目。"),
    paragraph("当前公式仅支持同工作表引用及 12 个基础函数，复制与填充按值处理；含公式的工作表限制行列结构操作与排序。XLSX 导入暂未实现。")
  );
  root.append(scope.root);
}
