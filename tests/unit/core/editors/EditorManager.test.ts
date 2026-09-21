import { describe, expect, it } from "vitest";

import { EditorManager } from "../../../../src/core/editors";
import type { CellEditor, CellEditorContext, ValidationResult } from "../../../../src/core/editors";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { CellRange } from "../../../../src/core/types";
import type { CellValue } from "../../../../src/core/sheet/types";

interface FakeContainer {
  readonly element: HTMLElement;
  dispatch(type: string, event: unknown): void;
}

function createFakeContainer(containsNext = false): FakeContainer {
  const listeners = new Map<string, Set<EventListener>>();
  const element = {
    addEventListener: (type: string, listener: EventListener) => {
      if (!listeners.has(type)) {
        listeners.set(type, new Set());
      }
      listeners.get(type)?.add(listener);
    },
    removeEventListener: (type: string, listener: EventListener) => {
      listeners.get(type)?.delete(listener);
    },
    contains: () => containsNext
  } as unknown as HTMLElement;

  return {
    element,
    dispatch(type: string, event: unknown) {
      for (const listener of listeners.get(type) ?? []) {
        (listener as (event: unknown) => void)(event);
      }
    }
  };
}

function createLayout(): LayoutEngine {
  return new LayoutEngine({
    rowCount: 10,
    columnCount: 5,
    rowHeight: 20,
    columnWidths: [50, 50, 50, 50, 50]
  });
}


const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 300, height: 200 };

function createScene(layout: LayoutEngine, viewport: ViewportState = VIEWPORT): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

function createRecordingEditor(
  overrides: {
    readonly value?: CellValue;
    readonly validate?: () => ValidationResult;
  } = {}
): {
  readonly editor: CellEditor;
  readonly calls: string[];
  getLastContext: () => CellEditorContext | undefined;
  getLastGeometry: () => CellEditorContext["rect"] | undefined;
} {
  const calls: string[] = [];
  let lastContext: CellEditorContext | undefined;
  let lastGeometry: CellEditorContext["rect"] | undefined;
  const value = overrides.value ?? "typed-value";

  const base = {
    mount(_container: HTMLElement, context: CellEditorContext): void {
      calls.push("mount");
      lastContext = context;
    },
    updateGeometry(rect: CellEditorContext["rect"]): void {
      calls.push("updateGeometry");
      lastGeometry = rect;
    },
    focus(): void {
      calls.push("focus");
    },
    getValue(): CellValue {
      calls.push("getValue");
      return value;
    },
    destroy(): void {
      calls.push("destroy");
    }
  };

  const editor: CellEditor = overrides.validate
    ? {
        ...base,
        validate: () => {
          calls.push("validate");
          return overrides.validate?.() as ValidationResult;
        }
      }
    : base;

  return {
    editor,
    calls,
    getLastContext: () => lastContext,
    getLastGeometry: () => lastGeometry
  };
}

function createManager(
  editorOverrides: Parameters<typeof createRecordingEditor>[0] = {},
  commitValue: (row: number, column: number, value: unknown) => void | Promise<void> = () => {},
  containerOverrides?: { readonly containsNext?: boolean }
) {
  const layout = createLayout();
  const recording = createRecordingEditor(editorOverrides);
  const container = createFakeContainer(containerOverrides?.containsNext ?? false);
  let scene = createScene(layout);
  const manager = new EditorManager({
    container: container.element,
    getScene: () => scene,
    getCellValue: () => "current-value",
    commitValue,
    editorFactory: () => recording.editor
  });
  return {
    manager,
    layout,
    dispatch: container.dispatch,
    setViewport: (viewport: ViewportState) => {
      scene = createScene(layout, viewport);
    },
    ...recording
  };
}

describe("EditorManager.startEdit", () => {
  it("mounts and focuses the editor, and reports isEditing", () => {
    const { manager, calls } = createManager();

    manager.startEdit(2, 1);

    expect(manager.isEditing()).toBe(true);
    expect(calls).toEqual(["mount", "focus"]);
  });

  it("positions the editor using LayoutEngine's cell rect, not its own math", () => {
    const { manager, layout, getLastContext } = createManager();

    manager.startEdit(3, 2);

    expect(getLastContext()?.rect).toEqual(layout.getCellRect(3, 2));
  });

  it("positions the editor relative to the current scroll offset, not the absolute content position", () => {
    const { manager, layout, getLastContext, setViewport } = createManager();

    setViewport({ scrollTop: 40, scrollLeft: 10, width: 200, height: 100 });
    manager.startEdit(3, 2);

    const absolute = layout.getCellRect(3, 2);
    expect(getLastContext()?.rect).toEqual({
      x: absolute.x - 10,
      y: absolute.y - 40,
      width: absolute.width,
      height: absolute.height
    });
  });

  it("limits a partially scrolled cell editor to the canonical body pane", () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 50],
      freeze: { rows: 1, columns: 1 }
    });
    const recording = createRecordingEditor();
    const container = createFakeContainer();
    const manager = new EditorManager({
      container: container.element,
      getScene: () => createScene(layout, {
        scrollTop: 10,
        scrollLeft: 20,
        width: 100,
        height: 40
      }),
      getCellValue: () => "current-value",
      commitValue: () => {},
      editorFactory: () => recording.editor
    });

    manager.startEdit(1, 1);

    expect(recording.getLastContext()?.rect).toEqual({
      x: 50,
      y: 20,
      width: 30,
      height: 10
    });
  });

  it("repositions the active editor from a newly published Scene", () => {
    const { manager, layout, getLastGeometry, setViewport } = createManager();
    manager.startEdit(3, 2);

    setViewport({ scrollTop: 40, scrollLeft: 10, width: 200, height: 100 });
    manager.updateGeometry();

    const absolute = layout.getCellRect(3, 2);
    expect(getLastGeometry()).toEqual({
      x: absolute.x - 10,
      y: absolute.y - 40,
      width: absolute.width,
      height: absolute.height
    });
  });

  it("cancels when the edited cell is absent from the new Scene panes", () => {
    const { manager, calls, setViewport } = createManager();
    manager.startEdit(3, 2);

    setViewport({ scrollTop: 0, scrollLeft: 0, width: 20, height: 20 });
    manager.updateGeometry();

    expect(manager.isEditing()).toBe(false);
    expect(calls).toContain("destroy");
  });

  it("does not create an editor for a cell outside every pane clip", () => {
    const { manager, calls, setViewport } = createManager();
    setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });

    manager.startEdit(9, 4);

    expect(manager.isEditing()).toBe(false);
    expect(calls).toEqual([]);
  });

  it("ignores a second startEdit call while already editing", () => {
    const { manager, calls } = createManager();

    manager.startEdit(0, 0);
    manager.startEdit(5, 5);

    expect(calls).toEqual(["mount", "focus"]);
  });
});

describe("EditorManager merge-spanning editor overlay", () => {
  // A 3-row-by-3-column merge whose anchor is not the origin.
  const MERGE: CellRange = { startRow: 2, endRow: 4, startColumn: 1, endColumn: 3 };
  const findMergeAt = (row: number, column: number): CellRange | undefined =>
    row >= MERGE.startRow && row <= MERGE.endRow && column >= MERGE.startColumn && column <= MERGE.endColumn
      ? MERGE
      : undefined;

  function createMergeAwareManager() {
    const layout = createLayout();
    const recording = createRecordingEditor();
    const container = createFakeContainer();
    let scene = createScene(layout);
    const manager = new EditorManager({
      container: container.element,
      getScene: () => scene,
      findMergeAt,
      getCellValue: () => "current-value",
      commitValue: () => {},
      editorFactory: () => recording.editor
    });
    return {
      manager,
      layout,
      ...recording,
      setViewport: (viewport: ViewportState) => {
        scene = createScene(layout, viewport);
      }
    };
  }

  it("spans the whole merge's rect when starting an edit at its anchor", () => {
    const { manager, getLastContext } = createMergeAwareManager();

    manager.startEdit(2, 1);

    expect(getLastContext()?.rect).toEqual({ x: 50, y: 40, width: 150, height: 60 });
  });

  it("keeps spanning the whole merge's rect after a Scene reprojection", () => {
    const { manager, getLastGeometry, setViewport } = createMergeAwareManager();
    manager.startEdit(2, 1);

    setViewport({ scrollTop: 0, scrollLeft: 10, width: 300, height: 200 });
    manager.updateGeometry();

    expect(getLastGeometry()).toEqual({ x: 40, y: 40, width: 150, height: 60 });
  });

  it("uses only the single cell's rect when the cell is not merged", () => {
    const { manager, getLastContext } = createMergeAwareManager();

    manager.startEdit(0, 0);

    expect(getLastContext()?.rect).toEqual({ x: 0, y: 0, width: 50, height: 20 });
  });
});

describe("EditorManager.commitEdit", () => {
  it("calls getValue, then commitValue, then destroy, exactly once each and in order", async () => {
    const commits: Array<{ row: number; column: number; value: unknown }> = [];
    const { manager, calls } = createManager({ value: "hello" }, (row, column, value) => {
      commits.push({ row, column, value });
    });

    manager.startEdit(1, 1);
    await manager.commitEdit();

    expect(calls).toEqual(["mount", "focus", "getValue", "destroy"]);
    expect(commits).toEqual([{ row: 1, column: 1, value: "hello" }]);
    expect(manager.isEditing()).toBe(false);
  });

  it("keeps the editor, value, and focus when validate() fails", async () => {
    let commitValueCalled = false;
    const { manager, calls } = createManager(
      { validate: () => ({ valid: false, message: "nope" }) },
      () => {
        commitValueCalled = true;
      }
    );

    manager.startEdit(0, 0);
    await manager.commitEdit();

    expect(calls).toEqual(["mount", "focus", "validate"]);
    expect(commitValueCalled).toBe(false);
    expect(manager.isEditing()).toBe(true);
  });

  it("does not tear down the editor when commitValue rejects", async () => {
    const { manager, calls } = createManager({}, () => Promise.reject(new Error("boom")));

    manager.startEdit(0, 0);

    await expect(manager.commitEdit()).rejects.toThrow("boom");
    expect(calls).not.toContain("destroy");
    expect(manager.isEditing()).toBe(true);
  });

  it("is a no-op when nothing is being edited", async () => {
    const { manager } = createManager();

    await expect(manager.commitEdit()).resolves.toBeUndefined();
    expect(manager.isEditing()).toBe(false);
  });
});

describe("EditorManager.cancelEdit", () => {
  it("destroys the editor without reading its value or committing", () => {
    let commitValueCalled = false;
    const { manager, calls } = createManager({}, () => {
      commitValueCalled = true;
    });

    manager.startEdit(4, 4);
    manager.cancelEdit();

    expect(calls).toEqual(["mount", "focus", "destroy"]);
    expect(commitValueCalled).toBe(false);
    expect(manager.isEditing()).toBe(false);
  });

  it("is a no-op when nothing is being edited", () => {
    const { manager } = createManager();

    expect(() => manager.cancelEdit()).not.toThrow();
    expect(manager.isEditing()).toBe(false);
  });
});

describe("EditorManager keydown handling", () => {
  it("Escape cancels the edit", () => {
    const { manager, dispatch, calls } = createManager();
    manager.startEdit(0, 0);

    let defaultPrevented = false;
    dispatch("keydown", { key: "Escape", preventDefault: () => (defaultPrevented = true) });

    expect(defaultPrevented).toBe(true);
    expect(calls).toContain("destroy");
    expect(manager.isEditing()).toBe(false);
  });

  it("Enter commits when not composing", async () => {
    const commits: unknown[] = [];
    const { manager, dispatch } = createManager({}, (_row, _column, value) => {
      commits.push(value);
    });
    manager.startEdit(0, 0);

    dispatch("keydown", { key: "Enter", preventDefault: () => {} });
    await Promise.resolve();
    await Promise.resolve();

    expect(commits).toHaveLength(1);
    expect(manager.isEditing()).toBe(false);
  });

  it("Enter does not commit while an IME composition is in progress", async () => {
    const commits: unknown[] = [];
    const { manager, dispatch } = createManager({}, (_row, _column, value) => {
      commits.push(value);
    });
    manager.startEdit(0, 0);

    dispatch("compositionstart", {});
    dispatch("keydown", { key: "Enter", preventDefault: () => {} });
    await Promise.resolve();

    expect(commits).toHaveLength(0);
    expect(manager.isEditing()).toBe(true);

    dispatch("compositionend", {});
    dispatch("keydown", { key: "Enter", preventDefault: () => {} });
    await Promise.resolve();
    await Promise.resolve();

    expect(commits).toHaveLength(1);
  });

  it("ignores keys other than Escape/Enter", () => {
    const { manager, dispatch, calls } = createManager();
    manager.startEdit(0, 0);

    dispatch("keydown", { key: "a", preventDefault: () => {} });

    expect(calls).toEqual(["mount", "focus"]);
    expect(manager.isEditing()).toBe(true);
  });

  it("does nothing when a key event arrives while nothing is being edited", () => {
    const { dispatch } = createManager();

    expect(() => dispatch("keydown", { key: "Enter", preventDefault: () => {} })).not.toThrow();
  });
});

describe("EditorManager focusout handling", () => {
  it("commits when focus moves outside the container", async () => {
    const commits: unknown[] = [];
    const { manager, dispatch } = createManager({}, (_row, _column, value) => {
      commits.push(value);
    });
    manager.startEdit(0, 0);

    dispatch("focusout", { relatedTarget: null });
    await Promise.resolve();
    await Promise.resolve();

    expect(commits).toHaveLength(1);
    expect(manager.isEditing()).toBe(false);
  });

  it("does not commit when focus moves to another node still inside the container", () => {
    const commits: unknown[] = [];
    const { manager, dispatch } = createManager({}, (_row, _column, value) => {
      commits.push(value);
    }, { containsNext: true });
    manager.startEdit(0, 0);

    dispatch("focusout", { relatedTarget: {} });

    expect(commits).toHaveLength(0);
    expect(manager.isEditing()).toBe(true);
  });

  it("does nothing when focusout arrives while nothing is being edited", () => {
    const { dispatch } = createManager();

    expect(() => dispatch("focusout", { relatedTarget: null })).not.toThrow();
  });
});

describe("EditorManager construction", () => {
  it("defaults to the built-in text editor when no editorFactory is provided", () => {
    const layout = createLayout();
    const container = createFakeContainer();

    expect(
      () =>
        new EditorManager({
          container: container.element,
          getScene: () => createScene(layout),
          getCellValue: () => "x",
          commitValue: () => {}
        })
    ).not.toThrow();
  });
});

describe("EditorManager concurrent cancel during a pending commit", () => {
  it("does not re-destroy the editor or fire onEditCommit once commitValue resolves", async () => {
    let resolveCommit: () => void = () => {};
    const commitPromise = new Promise<void>((resolve) => {
      resolveCommit = resolve;
    });
    const onEditCommitCalls: Array<readonly [number, number]> = [];
    const layout = createLayout();
    const recording = createRecordingEditor();
    const container = createFakeContainer();
    const manager = new EditorManager({
      container: container.element,
      getScene: () => createScene(layout),
      getCellValue: () => "current-value",
      commitValue: () => commitPromise,
      editorFactory: () => recording.editor,
      onEditCommit: (row, column) => onEditCommitCalls.push([row, column])
    });

    manager.startEdit(0, 0);
    const pendingCommit = manager.commitEdit();
    manager.cancelEdit();

    resolveCommit();
    await pendingCommit;

    expect(onEditCommitCalls).toHaveLength(0);
    expect(recording.calls.filter((call) => call === "destroy")).toHaveLength(1);
  });
});

describe("EditorManager.dispose cleans up listeners", () => {
  it("stops reacting to events after dispose", () => {
    const commits: unknown[] = [];
    const { manager, dispatch } = createManager({}, (_row, _column, value) => {
      commits.push(value);
    });
    manager.startEdit(0, 0);
    manager.dispose();

    dispatch("keydown", { key: "Enter", preventDefault: () => {} });
    dispatch("focusout", { relatedTarget: null });

    expect(commits).toHaveLength(0);
  });
});

describe("EditorManager.dispose", () => {
  it("is safe to call when nothing is being edited", () => {
    const { manager } = createManager();

    expect(() => manager.dispose()).not.toThrow();
  });

  it("forces the active edit to cancel rather than commit", () => {
    let commitValueCalled = false;
    const { manager, calls } = createManager({}, () => {
      commitValueCalled = true;
    });

    manager.startEdit(0, 0);
    manager.dispose();

    expect(calls).toEqual(["mount", "focus", "destroy"]);
    expect(commitValueCalled).toBe(false);
    expect(manager.isEditing()).toBe(false);
  });
});

describe("EditorManager initial input", () => {
  it("passes the typed character through as initialInput", () => {
    const { manager, getLastContext } = createManager();

    manager.startEdit(0, 0, "5");

    expect(getLastContext()?.initialInput).toBe("5");
    expect(getLastContext()?.initialValue).toBe("current-value");
  });
});

describe("EditorManager.resolveEditor — per-column editor resolution", () => {
  it("uses the editor resolveEditor returns instead of the default editorFactory", () => {
    const layout = createLayout();
    const container = createFakeContainer();
    const defaultRecording = createRecordingEditor({ value: "default" });
    const resolvedRecording = createRecordingEditor({ value: "resolved" });

    const manager = new EditorManager({
      container: container.element,
      getScene: () => createScene(layout),
      getCellValue: () => "current-value",
      commitValue: () => {},
      editorFactory: () => defaultRecording.editor,
      resolveEditor: () => resolvedRecording.editor
    });

    manager.startEdit(0, 0);

    expect(resolvedRecording.calls).toEqual(["mount", "focus"]);
    expect(defaultRecording.calls).toEqual([]);
  });

  it("falls back to the default editorFactory when resolveEditor returns undefined for that cell", () => {
    const layout = createLayout();
    const container = createFakeContainer();
    const defaultRecording = createRecordingEditor({ value: "default" });
    const resolvedRecording = createRecordingEditor({ value: "resolved" });

    const manager = new EditorManager({
      container: container.element,
      getScene: () => createScene(layout),
      getCellValue: () => "current-value",
      commitValue: () => {},
      editorFactory: () => defaultRecording.editor,
      // Only column 1 gets a resolved editor — proves this is a per-cell
      // decision, not a one-time global fallback check.
      resolveEditor: (_row, column) => (column === 1 ? resolvedRecording.editor : undefined)
    });

    manager.startEdit(0, 0);
    expect(defaultRecording.calls).toEqual(["mount", "focus"]);
    expect(resolvedRecording.calls).toEqual([]);

    manager.cancelEdit();
    manager.startEdit(0, 1);
    expect(resolvedRecording.calls).toEqual(["mount", "focus"]);
  });

  it("works normally when resolveEditor is not provided at all", () => {
    const { manager, calls } = createManager();

    manager.startEdit(0, 0);

    expect(calls).toEqual(["mount", "focus"]);
  });
});

describe("EditorManager.serializeCellValue — initial edit text", () => {
  it("passes serializeCellValue's return value through as initialText", () => {
    const layout = createLayout();
    const container = createFakeContainer();
    const recording = createRecordingEditor();

    const manager = new EditorManager({
      container: container.element,
      getScene: () => createScene(layout),
      // 46037 = 2026-01-15：日期在格子里是序列号，编辑时要看到的是日期串，
      // 这正是 serializeCellValue 存在的理由。
      getCellValue: () => 46037,
      commitValue: () => {},
      editorFactory: () => recording.editor,
      serializeCellValue: (_row, _column, value) => `serial:${String(value)}`
    });

    manager.startEdit(0, 0);

    expect(recording.getLastContext()?.initialText).toBe("serial:46037");
  });

  it("leaves initialText undefined when serializeCellValue is not provided", () => {
    const { manager, getLastContext } = createManager();

    manager.startEdit(0, 0);

    expect(getLastContext()?.initialText).toBeUndefined();
  });
});
