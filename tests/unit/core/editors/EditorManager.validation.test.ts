import { describe, expect, it } from "vitest";

import { EditorManager, outcomeFromResult } from "../../../../src/core/editors";
import type { CellEditor, EditValidationOutcome, ValidationPresenter } from "../../../../src/core/editors";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { CellValue } from "../../../../src/core/sheet/types";
import type { DataValidationResult } from "../../../../src/core/sheet/validations";

/**
 * 编辑器提交前的校验流程（数据验证设计稿 §7.1）。与 Univer 一致：只有停止级会拒绝输入，
 * 弹模态对话框；警告、信息放行。呈现器用假的：`EditorManager` 只管"什么时候要弹、弹什么"，
 * 对话框长什么样是 DOM 的事，Node 里没有 DOM。
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

type Shown = Parameters<ValidationPresenter["show"]>[0];

function createFakePresenter() {
  const calls: string[] = [];
  const shown: Shown[] = [];
  const presenter: ValidationPresenter = {
    show: (request) => {
      shown.push(request);
      calls.push("show");
    },
    hide: () => calls.push("hide"),
    dispose: () => calls.push("dispose")
  };
  return { presenter, calls, shown, last: () => shown.at(-1) };
}

function setup(
  outcome: EditValidationOutcome | ((value: CellValue | undefined) => EditValidationOutcome),
  options: { readonly focusoutOnHide?: boolean } = {}
) {
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
  if (options.focusoutOnHide) {
    // 浏览器行为：对话框里持有焦点的「确定」按钮被摘掉时会发一次 focusout，relatedTarget 为空。
    // 见 EditorManager 里 onClose 的注释。
    const hide = fake.presenter.hide;
    fake.presenter.hide = () => {
      hide();
      container.dispatch("focusout", { relatedTarget: null });
    };
  }
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

const reject: EditValidationOutcome = { kind: "reject", title: "数量无效", message: "需要整数" };

describe("EditorManager validation — accept", () => {
  it("commits normally and shows nothing", async () => {
    const { manager, fake, commits, validated } = setup({ kind: "accept" });
    manager.startEdit(2, 1);

    await manager.commitEdit();

    expect(validated).toEqual([[2, 1, "typed"]]);
    expect(commits).toEqual(["typed"]);
    expect(manager.isEditing()).toBe(false);
    expect(fake.shown).toEqual([]);
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
  it("commits nothing, keeps the editor open, and shows a dialog with the title and message", async () => {
    const { manager, fake, events, commits } = setup(reject);
    manager.startEdit(2, 1);
    events.length = 0;

    await manager.commitEdit();

    expect(commits).toEqual([]);
    expect(manager.isEditing()).toBe(true);
    expect(events).toEqual([]); // 没有 destroy
    expect(fake.last()).toMatchObject({ title: "数量无效", message: "需要整数" });
  });

  it("abandons the invalid draft on close and allows a fresh valid edit", async () => {
    let accept = false;
    const { manager, fake, events, commits } = setup(() => (accept ? { kind: "accept" } : reject));
    manager.startEdit(0, 0);
    await manager.commitEdit();
    events.length = 0;

    fake.last()?.onClose();
    expect(events).toEqual(["destroy"]);
    expect(manager.isEditing()).toBe(false);
    expect(commits).toEqual([]);

    accept = true;
    manager.startEdit(0, 0);
    await manager.commitEdit();
    expect(commits).toEqual(["typed"]);
    expect(manager.isEditing()).toBe(false);
  });

  it("while the dialog is open, Enter, focus-out and repeated commits are ignored", async () => {
    const { manager, container, fake, validated } = setup(reject);
    manager.startEdit(0, 0);
    await manager.commitEdit();

    container.dispatch("keydown", { key: "Enter", preventDefault: () => {} });
    container.dispatch("focusout", { relatedTarget: null });
    await manager.commitEdit();
    await Promise.resolve();

    expect(fake.shown).toHaveLength(1);
    expect(validated).toHaveLength(1);
  });

  it("closing the dialog does not raise it again when the removed OK button fires a focusout", async () => {
    // 回归：点「确定」→ 按钮被摘掉 → 浏览器发 focusout → 曾被当成"点到别处去了"而重新校验、
    // 再弹出同一个对话框，看上去就是点了没反应。
    const { manager, container, fake, commits, validated } = setup(reject, { focusoutOnHide: true });
    manager.startEdit(0, 0);
    await manager.commitEdit();

    fake.last()?.onClose();
    await Promise.resolve();
    await Promise.resolve();

    expect(fake.shown).toHaveLength(1);
    expect(validated).toHaveLength(1);
    expect(commits).toEqual([]);
    expect(manager.isEditing()).toBe(false);

    // Leaving the cell after acknowledging must not resubmit the rejected draft.
    container.dispatch("focusout", { relatedTarget: null });
    await manager.commitEdit();
    expect(fake.shown).toHaveLength(1);
  });

  it("Escape still cancels the edit and hides the dialog", async () => {
    const { manager, container, fake } = setup(reject);
    manager.startEdit(0, 0);
    await manager.commitEdit();

    container.dispatch("keydown", { key: "Escape", preventDefault: () => {} });

    expect(manager.isEditing()).toBe(false);
    expect(fake.calls.at(-1)).toBe("hide");
  });

  it("after cancelling, a fresh edit is validated normally again", async () => {
    const { manager, container, validated } = setup(reject);
    manager.startEdit(0, 0);
    await manager.commitEdit();
    container.dispatch("keydown", { key: "Escape", preventDefault: () => {} });

    manager.startEdit(1, 1);
    await manager.commitEdit();

    expect(validated).toHaveLength(2);
  });
});

describe("EditorManager validation — lifecycle", () => {
  it("disposes the presenter with the manager", () => {
    const { manager, fake } = setup({ kind: "accept" });
    manager.dispose();
    expect(fake.calls.at(-1)).toBe("dispose");
  });
});

describe("outcomeFromResult", () => {
  const base: DataValidationResult = { valid: false, ruleId: "r", message: "不对" };

  it("accepts a valid result", () => {
    expect(outcomeFromResult({ valid: true })).toEqual({ kind: "accept" });
  });

  it("rejects only stop-level rules — the default", () => {
    expect(outcomeFromResult({ ...base, errorStyle: "stop" })).toEqual({ kind: "reject", message: "不对" });
    expect(outcomeFromResult(base)).toEqual({ kind: "reject", message: "不对" });
  });

  it("lets warning and information values through (they are only marked), like Univer", () => {
    expect(outcomeFromResult({ ...base, errorStyle: "warning" })).toEqual({ kind: "accept" });
    expect(outcomeFromResult({ ...base, errorStyle: "information" })).toEqual({ kind: "accept" });
  });

  it("carries the title, and has a fallback message", () => {
    expect(outcomeFromResult({ ...base, title: "标题" })).toEqual({ kind: "reject", title: "标题", message: "不对" });
    expect(outcomeFromResult({ valid: false, errorStyle: "stop" })).toMatchObject({ kind: "reject", message: expect.any(String) });
  });
});
