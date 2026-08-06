/**
 * Fixed per-name execution semantics, per doc/customization-and-hooks.md §4.3.
 * - series: every tap runs, in order.
 * - bail: taps run in order until one cancels (via a cancellable context).
 * - waterfall: each tap's return value feeds the next tap.
 * - parallel: taps run independently; none may rely on another's side effects.
 */
export type HookMode = "series" | "bail" | "waterfall" | "parallel";

/**
 * `isolate` is orthogonal to `mode` — it's each hook name's own error
 * policy (doc §8), not implied by series/bail/waterfall/parallel:
 * `isolate: true` catches a tap's exception, reports it via `onHookError`,
 * and keeps going (display/transform hooks, safe to fall back on);
 * `isolate: false` lets the exception propagate to the caller, which is
 * expected to abort/roll back whatever operation is in progress (data
 * commit hooks).
 */
export interface HookDefinition {
  readonly mode: HookMode;
  readonly async: boolean;
  readonly isolate: boolean;
}

export type HookRegistry = Readonly<Record<string, HookDefinition>>;

export interface HookHandlerMeta {
  readonly name: string;
  /** Higher runs first. Same priority preserves registration order. Defaults to 0. */
  readonly priority?: number;
}
