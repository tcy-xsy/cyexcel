import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { Disposer } from "../types";

/** Structurally compatible with a real KeyboardEvent — CyExcelCore passes one directly; unit tests pass a plain object. */
export interface ShortcutMatchInput {
  readonly key: string;
  readonly ctrlKey: boolean;
  readonly metaKey: boolean;
  readonly shiftKey: boolean;
  readonly altKey: boolean;
}

/** Returning `false` explicitly falls through to the next older handler registered for the same combo; any other return value (including `undefined`) stops the chain. */
export type ShortcutHandler = () => void | false;

export interface ParsedShortcut {
  readonly mod: boolean;
  readonly shift: boolean;
  readonly alt: boolean;
  readonly key: string;
}

const MODIFIER_TOKENS = new Set(["mod", "shift", "alt"]);

/**
 * Parses a `"Mod+Shift+Z"`-style combo string. `Mod` means ctrlKey||metaKey
 * (the same collapse `CyExcelCore`'s pre-existing undo/redo check already
 * did). Modifier order doesn't matter; the key is always the last segment.
 * Throws CyExcelErrorCode.InvalidOptions (reusing the existing code, not a
 * new one) for anything that isn't exactly "zero or more known modifier
 * tokens, then a key".
 */
export function parseShortcut(combo: string): ParsedShortcut {
  const segments = combo.split("+").map((segment) => segment.trim());
  if (combo.trim().length === 0 || segments.some((segment) => segment.length === 0)) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidOptions,
      `Invalid shortcut combo ${JSON.stringify(combo)}: expected "[Mod+][Shift+][Alt+]<Key>"`
    );
  }

  const key = segments[segments.length - 1] as string;
  const modifierSegments = segments.slice(0, -1);
  let mod = false;
  let shift = false;
  let alt = false;
  for (const segment of modifierSegments) {
    const token = segment.toLowerCase();
    if (!MODIFIER_TOKENS.has(token)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `Invalid shortcut combo ${JSON.stringify(combo)}: unknown modifier ${JSON.stringify(segment)}`
      );
    }
    if (token === "mod") mod = true;
    else if (token === "shift") shift = true;
    else alt = true;
  }

  return { mod, shift, alt, key: key.toLowerCase() };
}

function toStackKey(parsed: ParsedShortcut): string {
  return `${parsed.mod ? 1 : 0}|${parsed.shift ? 1 : 0}|${parsed.alt ? 1 : 0}|${parsed.key}`;
}

function matches(parsed: ParsedShortcut, input: ShortcutMatchInput): boolean {
  return (
    parsed.key === input.key.toLowerCase() &&
    parsed.mod === (input.ctrlKey || input.metaKey) &&
    parsed.shift === input.shiftKey &&
    parsed.alt === input.altKey
  );
}

/**
 * A stack of handlers per normalized combo. `dispatch()` tries the
 * most-recently-registered handler first; a handler returning `false`
 * explicitly falls through to the next older one for the same combo — see
 * doc/progress.md's Step 18 entry for why this is the chosen conflict
 * policy (later registration, e.g. a plugin, overrides an earlier one, e.g.
 * a built-in, by default).
 */
export class ShortcutRegistry {
  readonly #stacks = new Map<string, { readonly parsed: ParsedShortcut; readonly handler: ShortcutHandler }[]>();

  register(combo: string, handler: ShortcutHandler): Disposer {
    const parsed = parseShortcut(combo);
    const stackKey = toStackKey(parsed);
    const entry = { parsed, handler };
    const stack = this.#stacks.get(stackKey) ?? [];
    stack.push(entry);
    this.#stacks.set(stackKey, stack);

    let disposed = false;
    return () => {
      if (disposed) {
        return;
      }
      disposed = true;
      const current = this.#stacks.get(stackKey);
      if (!current) {
        return;
      }
      const index = current.indexOf(entry);
      if (index >= 0) {
        current.splice(index, 1);
      }
      if (current.length === 0) {
        this.#stacks.delete(stackKey);
      }
    };
  }

  /** Returns whether some handler ran and did not fall through (i.e. the caller should treat the input as handled, e.g. call preventDefault()). */
  dispatch(input: ShortcutMatchInput): boolean {
    for (const stack of this.#stacks.values()) {
      if (stack.length === 0 || !matches(stack[0]!.parsed, input)) {
        continue;
      }
      for (let i = stack.length - 1; i >= 0; i -= 1) {
        const result = stack[i]!.handler();
        if (result !== false) {
          return true;
        }
      }
      return false;
    }
    return false;
  }

  dispose(): void {
    this.#stacks.clear();
  }
}
