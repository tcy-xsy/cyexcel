import { translate } from "../i18n";

export function setText(node: HTMLElement, text: string): void {
  node.textContent = translate(text);
}

export function element<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== undefined) {
    node.className = className;
  }
  if (text !== undefined) {
    setText(node, text);
  }
  return node;
}

export function pageHeader(eyebrow: string, title: string, description: string): HTMLElement {
  const header = element("header", "docs-page-header");
  header.append(
    element("p", "page-eyebrow", eyebrow),
    element("h1", "page-title", title),
    element("p", "page-description", description)
  );
  return header;
}

export function section(title: string): { root: HTMLElement; body: HTMLElement } {
  const root = element("section", "docs-section");
  const body = element("div", "prose");
  root.append(element("h2", "section-title", title), body);
  return { root, body };
}

export function paragraph(text: string): HTMLParagraphElement {
  return element("p", "docs-paragraph", text);
}

export function featureGrid(items: ReadonlyArray<{ title: string; text: string }>): HTMLElement {
  const grid = element("div", "feature-grid");
  for (const item of items) {
    const card = element("article", "feature-card");
    card.append(element("h3", undefined, item.title), element("p", undefined, item.text));
    grid.append(card);
  }
  return grid;
}

export function codeBlock(code: string, language = "ts"): HTMLElement {
  const localizedCode = translate(code);
  const root = element("figure", "code-block");
  const head = element("div", "code-head");
  const label = element("span", undefined, language);
  const button = element("button", "copy-button", "复制");
  button.type = "button";
  button.addEventListener("click", async () => {
    await navigator.clipboard.writeText(localizedCode);
    setText(button, "已复制");
    window.setTimeout(() => {
      setText(button, "复制");
    }, 1400);
  });
  const pre = document.createElement("pre");
  pre.textContent = localizedCode;
  head.append(label, button);
  root.append(head, pre);
  return root;
}

export function demoBlock(options: {
  title: string;
  description?: string;
  actions?: HTMLButtonElement[];
}): { root: HTMLElement; mount: HTMLDivElement; log?: HTMLDivElement } {
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

export function actionButton(label: string, onClick: () => void | Promise<void>): HTMLButtonElement {
  const button = element("button", "action-button", label);
  button.type = "button";
  button.addEventListener("click", () => {
    void onClick();
  });
  return button;
}

export type ApiRow = ReadonlyArray<{ name: string; type?: string; description: string }>;

export function apiTable(rows: ApiRow): HTMLElement {
  const table = document.createElement("table");
  table.className = "api-table";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (const label of ["名称", "类型", "说明"]) {
    headRow.append(element("th", undefined, label));
  }
  thead.append(headRow);
  const tbody = document.createElement("tbody");
  for (const row of rows) {
    const tr = document.createElement("tr");
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

export function callout(text: string): HTMLElement {
  const root = element("aside", "callout");
  root.append(element("p", undefined, text));
  return root;
}
