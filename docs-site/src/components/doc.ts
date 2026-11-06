import { translate } from "../i18n";
import { DocElement } from "./content";

export function setText(node: DocElement, text: string): void {
  node.setText(translate(text));
}

export function element(
  tag: keyof HTMLElementTagNameMap,
  className?: string,
  text?: string
): DocElement {
  const node = new DocElement(tag);
  if (className !== undefined) {
    node.addClass(className);
  }
  if (text !== undefined) {
    setText(node, text);
  }
  return node;
}

export function pageHeader(eyebrow: string, title: string, description: string): DocElement {
  const header = element("header", "docs-page-header");
  header.append(
    element("p", "page-eyebrow", eyebrow),
    element("h1", "page-title", title),
    element("p", "page-description", description)
  );
  return header;
}

export function section(title: string): { root: DocElement; body: DocElement } {
  const root = element("section", "docs-section");
  const body = element("div", "prose");
  root.append(element("h2", "section-title", title), body);
  return { root, body };
}

export function paragraph(text: string): DocElement {
  return element("p", "docs-paragraph", text);
}

export function featureGrid(items: ReadonlyArray<{ title: string; text: string }>): DocElement {
  const grid = element("div", "feature-grid");
  for (const item of items) {
    const card = element("article", "feature-card");
    card.append(element("h3", undefined, item.title), element("p", undefined, item.text));
    grid.append(card);
  }
  return grid;
}

export function codeBlock(code: string, language = "ts"): DocElement {
  const localizedCode = translate(code);
  const root = element("figure", "code-block");
  const head = element("div", "code-head");
  const label = element("span", undefined, language);
  const button = element("button", "copy-button", "复制");
  button.attributes.type = "button";
  button.onClick = async () => {
    await navigator.clipboard.writeText(localizedCode);
    setText(button, "已复制");
    window.setTimeout(() => {
      setText(button, "复制");
    }, 1400);
  };
  const pre = new DocElement("pre");
  pre.setText(localizedCode);
  head.append(label, button);
  root.append(head, pre);
  return root;
}

export function demoBlock(options: {
  title: string;
  description?: string;
  actions?: DocElement[];
}): { root: DocElement; mount: DocElement; log?: DocElement } {
  const root = element("section", "demo-block");
  const head = element("div", "demo-head");
  head.append(element("h3", "demo-title", options.title));
  if (options.description !== undefined) {
    head.append(element("p", "demo-description", options.description));
  }
  const mount = element("div", "demo-mount");
  root.append(head);
  if (options.actions !== undefined && options.actions.length > 0) {
    const toolbar = element("div", "demo-toolbar");
    toolbar.append(...options.actions);
    root.append(toolbar);
  }
  root.append(mount);

  const hasLog = options.actions !== undefined && options.actions.length > 0;
  if (hasLog) {
    const log = element("div", "demo-log");
    setText(log, "等待操作…");
    root.append(log);
    return { root, mount, log };
  }
  return { root, mount };
}

export function actionButton(label: string, onClick: () => void | Promise<void>): DocElement {
  const button = element("button", "action-button", label);
  button.attributes.type = "button";
  button.onClick = onClick;
  return button;
}

export type ApiRow = ReadonlyArray<{ name: string; type?: string; description: string }>;

export function apiTable(rows: ApiRow): DocElement {
  const table = element("table", "api-table");
  const thead = element("thead");
  const headRow = element("tr");
  for (const label of ["名称", "类型", "说明"]) {
    headRow.append(element("th", undefined, label));
  }
  thead.append(headRow);
  const tbody = element("tbody");
  for (const row of rows) {
    const tr = element("tr");
    tr.append(
      element("td", "api-name", row.name),
      element("td", "api-type", row.type ?? "-"),
      element("td", undefined, row.description)
    );
    tbody.append(tr);
  }
  table.append(thead, tbody);
  return table;
}

export function callout(text: string): DocElement {
  const root = element("aside", "callout");
  root.append(element("p", undefined, text));
  return root;
}
