import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { Disposer } from "../types";
import type { HookDefinition, HookHandlerMeta, HookRegistry } from "./types";

type AnyHandler = (...args: never[]) => unknown;

interface StoredTap {
  readonly seq: number;
  readonly name: string;
  readonly priority: number;
  readonly handler: AnyHandler;
  readonly isAsync: boolean;
}

/** Sentinel distinguishing "a waterfall tap threw and was isolated" from "a tap legitimately returned undefined". */
const SKIPPED = Symbol("hook-skipped");

/**
 * Generic Hook engine — knows nothing about any specific hook name's
 * meaning, only what `registry` (doc/customization-and-hooks.md §4.3's
 * fixed per-name mode/async/isolate table, see `./registry.ts`) says about
 * it. Real hook names and their call sites live in the modules that fire
 * them (CommandManager, ColumnManager, ContextMenuService, CyExcelCore).
 */
export interface HookManagerOptions {
  readonly registry: HookRegistry;
}

export class HookManager {
  readonly #registry: HookRegistry;
  readonly #taps = new Map<string, StoredTap[]>();
  #seq = 0;

  constructor(options: HookManagerOptions) {
    this.#registry = options.registry;
  }

  tap<TContext = unknown, TValue = unknown>(
    hookName: string,
    meta: HookHandlerMeta,
    handler: ((context: TContext) => void) | ((value: TValue, context: TContext) => TValue)
  ): Disposer {
    this.#definitionFor(hookName);
    return this.#register(hookName, meta, handler as AnyHandler, false);
  }

  tapAsync<TContext = unknown, TValue = unknown>(
    hookName: string,
    meta: HookHandlerMeta,
    handler:
      | ((context: TContext, signal: AbortSignal) => Promise<void>)
      | ((value: TValue, context: TContext, signal: AbortSignal) => Promise<TValue>)
  ): Disposer {
    const definition = this.#definitionFor(hookName);
    if (!definition.async) {
      throw new CyExcelError(
        CyExcelErrorCode.AsyncHookNotSupported,
        `Hook "${hookName}" does not accept async taps — use tap() instead`
      );
    }
    return this.#register(hookName, meta, handler as AnyHandler, true);
  }

  /** All taps run in order; a non-isolated hook's exception propagates and stops the remaining taps. */
  runSeries<TContext>(hookName: string, context: TContext): void {
    const definition = this.#definitionFor(hookName);
    for (const tap of this.#orderedTaps(hookName)) {
      this.#invoke(hookName, tap, definition, () => (tap.handler as (c: TContext) => void)(context));
    }
  }

  /** Taps run in order until `context.cancelled` becomes true (see createCancellableContext). */
  runBail<TContext extends { readonly cancelled: boolean }>(hookName: string, context: TContext): void {
    const definition = this.#definitionFor(hookName);
    for (const tap of this.#orderedTaps(hookName)) {
      this.#invoke(hookName, tap, definition, () => (tap.handler as (c: TContext) => void)(context));
      if (context.cancelled) {
        return;
      }
    }
  }

  /** Each tap's return value feeds the next; a tap isolated after throwing leaves the value unchanged. */
  runWaterfall<TValue, TContext>(hookName: string, initialValue: TValue, context: TContext): TValue {
    const definition = this.#definitionFor(hookName);
    let value = initialValue;
    for (const tap of this.#orderedTaps(hookName)) {
      const result = this.#invoke(hookName, tap, definition, () =>
        (tap.handler as (v: TValue, c: TContext) => TValue)(value, context)
      );
      if (result !== SKIPPED) {
        value = result;
      }
    }
    return value;
  }

  /** Taps run independently against the same context — none may rely on another's mutations. */
  runParallel<TContext>(hookName: string, context: TContext): void {
    const definition = this.#definitionFor(hookName);
    for (const tap of this.#orderedTaps(hookName)) {
      this.#invoke(hookName, tap, definition, () => (tap.handler as (c: TContext) => void)(context));
    }
  }

  async runSeriesAsync<TContext>(hookName: string, context: TContext, signal: AbortSignal): Promise<void> {
    const definition = this.#definitionFor(hookName);
    for (const tap of this.#orderedTaps(hookName)) {
      await this.#invokeAsync(hookName, tap, definition, () =>
        (tap.handler as (c: TContext, s: AbortSignal) => void | Promise<void>)(context, signal)
      );
    }
  }

  async runBailAsync<TContext extends { readonly cancelled: boolean }>(
    hookName: string,
    context: TContext,
    signal: AbortSignal
  ): Promise<void> {
    const definition = this.#definitionFor(hookName);
    for (const tap of this.#orderedTaps(hookName)) {
      await this.#invokeAsync(hookName, tap, definition, () =>
        (tap.handler as (c: TContext, s: AbortSignal) => void | Promise<void>)(context, signal)
      );
      if (context.cancelled || signal.aborted) {
        return;
      }
    }
  }

  async runWaterfallAsync<TValue, TContext>(
    hookName: string,
    initialValue: TValue,
    context: TContext,
    signal: AbortSignal
  ): Promise<TValue> {
    const definition = this.#definitionFor(hookName);
    let value = initialValue;
    for (const tap of this.#orderedTaps(hookName)) {
      if (signal.aborted) {
        return value;
      }
      const result = await this.#invokeAsync(hookName, tap, definition, () =>
        (tap.handler as (v: TValue, c: TContext, s: AbortSignal) => TValue | Promise<TValue>)(value, context, signal)
      );
      if (result !== SKIPPED) {
        value = result;
      }
    }
    return value;
  }

  dispose(): void {
    this.#taps.clear();
  }

  #definitionFor(hookName: string): HookDefinition {
    const definition = this.#registry[hookName];
    if (!definition) {
      throw new CyExcelError(CyExcelErrorCode.UnknownHook, `Unknown hook "${hookName}"`);
    }
    return definition;
  }

  #register(hookName: string, meta: HookHandlerMeta, handler: AnyHandler, isAsync: boolean): Disposer {
    const list = this.#taps.get(hookName) ?? [];
    const tap: StoredTap = { seq: this.#seq, name: meta.name, priority: meta.priority ?? 0, handler, isAsync };
    this.#seq += 1;
    list.push(tap);
    this.#taps.set(hookName, list);

    let disposed = false;
    return () => {
      if (disposed) {
        return;
      }
      disposed = true;
      const current = this.#taps.get(hookName);
      const index = current?.indexOf(tap) ?? -1;
      if (index >= 0) {
        current?.splice(index, 1);
      }
    };
  }

  /** Higher priority first; same priority preserves registration order (stable sort by insertion sequence). */
  #orderedTaps(hookName: string): readonly StoredTap[] {
    const list = this.#taps.get(hookName);
    if (!list || list.length === 0) {
      return [];
    }
    return [...list].sort((a, b) => b.priority - a.priority || a.seq - b.seq);
  }

  #invoke<T>(hookName: string, tap: StoredTap, definition: HookDefinition, fn: () => T): T | typeof SKIPPED {
    if (!definition.isolate) {
      return fn();
    }
    try {
      return fn();
    } catch (error) {
      this.#reportError(hookName, tap.name, error);
      return SKIPPED;
    }
  }

  async #invokeAsync<T>(
    hookName: string,
    tap: StoredTap,
    definition: HookDefinition,
    fn: () => T | Promise<T>
  ): Promise<T | typeof SKIPPED> {
    if (!definition.isolate) {
      return fn();
    }
    try {
      return await fn();
    } catch (error) {
      this.#reportError(hookName, tap.name, error);
      return SKIPPED;
    }
  }

  /**
   * `onHookError` is itself a series/isolate hook (see ./registry.ts) —
   * routing through runSeries() here would let a throwing onHookError
   * handler re-enter this same method forever. The guard below is the
   * doc/customization-and-hooks.md §8 "onHookError 自身失败：记录到开发
   * 模式控制台，避免递归调用" requirement.
   */
  #reportError(hookName: string, handlerName: string, error: unknown): void {
    if (hookName === "onHookError") {
      console.error(`[cyexcel] "onHookError" handler "${handlerName}" itself threw`, error);
      return;
    }
    if (this.#registry.onHookError) {
      this.runSeries("onHookError", { error, hookName, handlerName });
    }
  }
}
