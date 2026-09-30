import { describe, expect, it } from "vitest";

import { EditorManager } from "../../../../src/core/editors";
import type { CellEditor, EditValidationOutcome, ValidationPresenter } from "../../../../src/core/editors";
import { outcomeFromResult } from "../../../../src/core/editors";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { CellValue } from "../../../../src/core/sheet/types";
import type { DataValidationResult } from "../../../../src/core/sheet/validations";

/**
 * 编辑器提交前的校验流程（数据验证设计稿 §7.1）。呈现器用假的：`EditorManager` 只管
 * "什么时候要提示、要什么"，气泡长什么样是 DOM 的事，Node 里没有 DOM。
 */

function createFakeContainer() {
  const listeners = new Map<string, Set<EventListener>>();
  const element = {
    addEventListener: (type: string, listener: EventListener) => {
      if (!listeners.has(type)) {
        listeners.set(type, new Set());
      }
      listeners.get(type)?.add(listener);
    },
    removeEventListener: (type: string, listener: EventListener) => listeners.get(type)?.delete(listener),
    contains: () => false
  } as unknown as HTMLElement;
  return {
    element,
    dispatch(type: string, event: unknown = {}) {
      for (const listener of listeners.get(type) ?? []) {
        (listener as (event: unknown) => void)(event);
      }
    }
  };
}

interface PresenterCall {
  readonly method: "show" | "hide" | "reposition" | "dispose";
  readonly request?: Parameters<ValidationPresenter["show"]>[0];
}

function createFakePresenter() {
  const calls: PresenterCall[] = [];
  let last: Parameters<ValidationPresenter["show"]>[0] | undefined;
  const presenter: ValidationPresenter = {
    show: (request) => {
      last = request;
      calls.push({ method: "show", request });
    },
    reposition: () => calls.push({ method: "reposition" }),
    hide: () => calls.push({ method: "hide" }),
    dispose: () => calls.push({ method: "dispose" })
  };
  return {
    presenter,
    calls,
    shows: () => calls.filter((call) => call.method === "show").map((call) => call.request),
    last: () => last
  };
}

function setup(outcome: EditValidationOutcome | ((value: CellValue | undefined) => EditValidationOutcome)) {
  const layout = new LayoutEngine({ rowCount: 10, columnCount: 5, rowHeight: 20, columnWidths: [50, 50, 50, 50, 50] });
  const scene = new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, scrollTop: 0, scrollLeft: 0, width: 300, height: 200 })
  });
  const events: string[] = [];
  const editor: CellEditor = {
    mount: () => events.push("mount"),
    updateGeometry: () => {},
    focus: () => events.push("focus"),
    getValue: () => "typed",
    destroy: () => events.push("destroy")
  };
  const container = createFakeContainer();
  const fake = createFakePresenter();
  const commits: unknown[] = [];
  const validated: Array<[number, number, unknown]> = [];
  const manager = new EditorManager({
    container: container.element,
    getScene: () => scene,
    getCellValue: () => "old",
    commitValue: (_row, _column, value) => {
      commits.push(value);
    },
    editorFactory: () => editor,
    validationPresenter: fake.presenter,
    validateCommit: (row, column, value) => {
      validated.push([row, column, value]);
      return typeof outcome === "function" ? outcome(value) : outcome;
    }
  });
  return { manager, container, fake, events, commits, validated };
}

describe("EditorManager validation — accept", () => {
  it("commits normally and shows nothing", async () => {
    const { manager, fake, commits, validated } = setup({ kind: "accept" });
    manager.startEdit(2, 1);

    await manager.commitEdit();

    expect(validated).toEqual([[2, 1, "typed"]]);
    expect(commits).toEqual(["typed"]);
    expect(manager.isEditing()).toBe(false);
    expect(fake.shows()).toEqual([]);
  });

  it("behaves exactly as before when no validateCommit is configured", async () => {
    const layout = new LayoutEngine({ rowCount: 2, columnCount: 2, rowHeight: 20, columnWidths: [50, 50] });
    const scene = new SceneSnapshot({
      version: 1,
      skeleton: SkeletonSnapshot.fromLayout(layout, 1),
      viewport: new ViewportSnapshot({ version: 1, scrollTop: 0, scrollLeft: 0, width: 100, height: 40 })
    });
    const commits: unknown[] = [];
    const manager = new EditorManager({
      container: createFakeContainer().element,
      getScene: () => scene,
      getCellValue: () => undefined,
      commitValue: (_r, _c, value) => {
        commits.push(value);
      },
      editorFactory: () => ({
        mount: () => {},
        updateGeometry: () => {},
        focus: () => {},
        getValue: () => "v",
        destroy: () => {}
      })
    });
    manager.startEdit(0, 0);
    await manager.commitEdit();
    expect(commits).toEqual(["v"]);
  });
});

describe("EditorManager validation — stop (reject)", () => {
  const reject: EditValidationOutcome = { kind: "reject", title: "输入无效", message: "需要整数" };

  it("keeps the editor open, refocuses it, commits nothing, and shows the message under the cell", async () => {
    const { manager, fake, events, commits } = setup(reject);
    manager.startEdit(2, 1);
    events.length = 0;

    await manager.commitEdit();

    expect(commits).toEqual([]);
    expect(manager.isEditing()).toBe(true);
    expect(events).toEqual(["focus"]); // 没有 destroy
    const shown = fake.last();
    expect(shown).toMatchObject({ kind: "reject", title: "输入无效", message: "需要整数" });
    expect(shown?.rect).toEqual({ x: 50, y: 40, width: 50, height: 20 }); // 第 2 行第 1 列
  });

  it("hides the bubble once the user edits again, and commits when the next attempt is accepted", async () => {
    let accept = false;
    const { manager, container, fake, commits } = setup(() => (accept ? { kind: "accept" } : reject));
    manager.startEdit(0, 0);
    await manager.commitEdit();
    expect(fake.calls.some((c) => c.method === "hide")).toBe(false);

    container.dispatch("input");
    expect(fake.calls.filter((c) => c.method === "hide")).toHaveLength(1);

    accept = true;
    await manager.commitEdit();
    expect(commits).toEqual(["typed"]);
    expect(manager.isEditing()).toBe(false);
  });

  it("Enter and focus-out both go through the same gate", async () => {
    const { manager, container, commits } = setup(reject);
    manager.startEdit(0, 0);

    container.dispatch("keydown", { key: "Enter", preventDefault: () => {} });
    container.dispatch("focusout", { relatedTarget: null });
    await Promise.resolve();
    await Promise.resolve();

    expect(commits).toEqual([]);
    expect(manager.isEditing()).toBe(true);
  });

  it("Escape still cancels and hides the bubble", async () => {
    const { manager, container, fake } = setup(reject);
    manager.startEdit(0, 0);
    await manager.commitEdit();

    container.dispatch("keydown", { key: "Escape", preventDefault: () => {} });

    expect(manager.isEditing()).toBe(false);
    expect(fake.calls.at(-1)?.method).toBe("hide");
  });
});

describe("EditorManager validation — warning (confirm)", () => {
  const confirm: EditValidationOutcome = { kind: "confirm", message: "值偏大" };

  it("waits for the answer: nothing committed, editor stays, later commit attempts are ignored", async () => {
    const { manager, container, fake, commits, validated } = setup(confirm);
    manager.startEdit(0, 0);

    await manager.commitEdit();
    container.dispatch("focusout", { relatedTarget: null });
    await Promise.resolve();

    expect(fake.shows()).toHaveLength(1);
    expect(fake.last()).toMatchObject({ kind: "confirm", message: "值偏大" });
    expect(commits).toEqual([]);
    expect(manager.isEditing()).toBe(true);
    expect(validated).toHaveLength(1); // 失焦没有再触发一轮校验
  });

  it("commits when the user chooses to continue", async () => {
    const { manager, fake, commits } = setup(confirm);
    manager.startEdit(0, 0);
    await manager.commitEdit();

    fake.last()?.onContinue?.();
    await Promise.resolve();
    await Promise.resolve();

    expect(commits).toEqual(["typed"]);
    expect(manager.isEditing()).toBe(false);
  });

  it("returns to the editor, uncommitted, when the user cancels", async () => {
    const { manager, fake, commits, events } = setup(confirm);
    manager.startEdit(0, 0);
    await manager.commitEdit();
    events.length = 0;

    fake.last()?.onCancel?.();

    expect(commits).toEqual([]);
    expect(manager.isEditing()).toBe(true);
    expect(events).toEqual(["focus"]);
    // 取消之后可以再次提交（不再被挡住）
    await manager.commitEdit();
    expect(fake.shows()).toHaveLength(2);
  });

  it("typing while the prompt is open dismisses it and re-validates on the next commit", async () => {
    const { manager, container, fake, validated } = setup(confirm);
    manager.startEdit(0, 0);
    await manager.commitEdit();

    container.dispatch("input");
    await manager.commitEdit();

    expect(validated).toHaveLength(2);
    expect(fake.shows()).toHaveLength(2);
  });
});

describe("EditorManager validation — information (notify)", () => {
  it("commits first, then shows the message", async () => {
    const { manager, fake, commits } = setup({ kind: "notify", message: "已记录" });
    manager.startEdit(1, 1);

    await manager.commitEdit();

    expect(commits).toEqual(["typed"]);
    expect(manager.isEditing()).toBe(false);
    expect(fake.last()).toMatchObject({ kind: "notify", message: "已记录" });
    // 提示是提交完成之后才出现的：show 之前的最后一次 hide 来自 finishCommit
    const methods = fake.calls.map((c) => c.method);
    expect(methods.lastIndexOf("show")).toBeGreaterThan(methods.indexOf("hide"));
  });
});

describe("EditorManager validation — lifecycle", () => {
  it("disposes the presenter with the manager", () => {
    const { manager, fake } = setup({ kind: "accept" });
    manager.dispose();
    expect(fake.calls.at(-1)?.method).toBe("dispose");
  });

  it("moves the bubble when the scene is republished", async () => {
    const { manager, fake } = setup({ kind: "reject", message: "x" });
    manager.startEdit(0, 0);
    manager.updateGeometry();
    expect(fake.calls.some((c) => c.method === "reposition")).toBe(true);
  });
});

describe("outcomeFromResult", () => {
  const base: DataValidationResult = { valid: false, ruleId: "r", message: "不对" };

  it("maps a valid result to accept", () => {
    expect(outcomeFromResult({ valid: true })).toEqual({ kind: "accept" });
  });

  it.each([
    ["stop", "reject"],
    ["warning", "confirm"],
    ["information", "notify"]
  ] as const)("maps errorStyle %s to %s", (errorStyle, kind) => {
    expect(outcomeFromResult({ ...base, errorStyle })).toEqual({ kind, message: "不对" });
  });

  it("defaults to stop, and carries the title", () => {
    expect(outcomeFromResult({ ...base, title: "标题" })).toEqual({ kind: "reject", title: "标题", message: "不对" });
  });
});
