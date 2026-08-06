export interface CancellableContext {
  cancel(reason?: string): void;
  readonly cancelled: boolean;
  readonly reason: string | undefined;
}

/**
 * Bail-mode hook contexts (doc/customization-and-hooks.md §4.4) each
 * declare their own `cancel(reason?)` method inline, but the mechanics
 * (flip a flag, remember the reason, expose both back to the caller) are
 * identical every time — this is the one shared implementation.
 * `HookManager.runBail()` checks `context.cancelled` after every tap to
 * decide whether to stop.
 */
export function createCancellableContext<TBase extends object>(base: TBase): TBase & CancellableContext {
  let cancelled = false;
  let reason: string | undefined;
  return {
    ...base,
    cancel(nextReason?: string) {
      cancelled = true;
      reason = nextReason;
    },
    get cancelled() {
      return cancelled;
    },
    get reason() {
      return reason;
    }
  };
}
