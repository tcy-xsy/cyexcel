import { describe, expect, it } from "vitest";
import {
  snapshotColumnDefinitions,
  type CanvasCellRenderer,
  type DomHeaderRendererFactory
} from "../../../../src/core/cells";

describe("snapshotColumnDefinitions", () => {
  it("owns declarative nested metadata while preserving behavior object identity", () => {
    const renderer: CanvasCellRenderer = { kind: "canvas", paint: () => {} };
    const headerRenderer: DomHeaderRendererFactory = {
      kind: "dom",
      pointerOwnership: "surface",
      create: () => ({ mount: () => {}, update: () => {}, destroy: () => {} })
    };
    const createEditor = () => ({
      mount: () => {},
      updateGeometry: () => {},
      focus: () => {},
      getValue: () => "",
      destroy: () => {}
    });
    const header = { title: "Initial", renderer: headerRenderer };
    const option = { value: "a", label: "A" };
    const options = [option];
    const clipboard = { copy: true, paste: false };
    const input = [{ header, options, clipboard, renderer, createEditor }];

    const snapshot = snapshotColumnDefinitions(input);
    const column = snapshot[0];

    expect(snapshot).not.toBe(input);
    expect(column).not.toBe(input[0]);
    expect(column?.header).not.toBe(header);
    expect(column?.options).not.toBe(options);
    expect(column?.options?.[0]).not.toBe(option);
    expect(column?.clipboard).not.toBe(clipboard);
    expect(column?.renderer).toBe(renderer);
    expect(typeof column?.header === "object" ? column.header.renderer : undefined).toBe(headerRenderer);
    expect(column?.createEditor).toBe(createEditor);
    expect([
      snapshot,
      column,
      column?.header,
      column?.options,
      column?.options?.[0],
      column?.clipboard
    ].every((value) => Object.isFrozen(value))).toBe(true);

    header.title = "Mutated";
    option.label = "Mutated";
    clipboard.paste = true;
    expect(column?.header).toMatchObject({ title: "Initial" });
    expect(column?.options?.[0]).toEqual({ value: "a", label: "A" });
    expect(column?.clipboard).toEqual({ copy: true, paste: false });
  });
});
