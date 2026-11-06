export type Cleanup = () => void;

export class DocElement {
  readonly children: DocElement[] = [];
  readonly attributes: Record<string, string> = {};
  onClick?: () => void | Promise<void>;
  onMount?: () => void | Cleanup;
  private mounted: HTMLElement | undefined;
  private text = "";

  constructor(readonly tag: keyof HTMLElementTagNameMap) {}

  append(...children: DocElement[]): void {
    this.children.push(...children);
  }

  addClass(name: string): void {
    this.attributes.class = [this.attributes.class, name].filter(Boolean).join(" ");
  }

  setText(text: string): void {
    this.text = text;
    if (this.mounted !== undefined) this.mounted.textContent = text;
  }

  get textContent(): string {
    return this.text + this.children.map(child => child.textContent).join("");
  }

  get element(): HTMLElement {
    if (this.mounted === undefined) throw new Error("Document content has not been mounted");
    return this.mounted;
  }

  findByClass(name: string): DocElement | undefined {
    if (this.attributes.class?.split(" ").includes(name)) return this;
    for (const child of this.children) {
      const found = child.findByClass(name);
      if (found !== undefined) return found;
    }
    return undefined;
  }

  toHtml(): string {
    const attributes = Object.entries(this.attributes).map(([key, value]) => ` ${key}="${escapeHtml(value)}"`).join("");
    return `<${this.tag}${attributes}>${escapeHtml(this.text)}${this.children.map(child => child.toHtml()).join("")}</${this.tag}>`;
  }

  mount(container: HTMLElement): Cleanup {
    const initializers: Array<() => void | Cleanup> = [];
    const cleanups: Cleanup[] = [];
    const nodes: DocElement[] = [];
    const create = (content: DocElement): HTMLElement => {
      const node = document.createElement(content.tag);
      content.mounted = node;
      nodes.push(content);
      for (const [key, value] of Object.entries(content.attributes)) node.setAttribute(key, value);
      node.textContent = content.text;
      if (content.onClick !== undefined) node.addEventListener("click", content.onClick);
      node.append(...content.children.map(create));
      if (content.onMount !== undefined) initializers.push(content.onMount);
      return node;
    };
    const dispose = (): void => {
      for (const cleanup of cleanups.reverse()) cleanup();
      cleanups.length = 0;
      for (const node of nodes) node.mounted = undefined;
    };
    container.replaceChildren(create(this));
    try {
      for (const initialize of initializers) {
        const cleanup = initialize();
        if (typeof cleanup === "function") cleanups.push(cleanup);
      }
    } catch (error) {
      dispose();
      throw error;
    }
    return dispose;
  }
}

export function escapeHtml(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
}
