import type { Disposer } from "../types";

/**
 * Plain pub/sub for completed-operation notifications
 * (doc/customization-and-hooks.md §4.1: "Event 只能观察结果，不能改变已经提交的
 * 结果"). Unlike HookManager there's no mode/priority/cancellation — every
 * listener just gets called, in registration order, and a listener's own
 * exception never stops the others or the emitter.
 */
export class EventBus {
  readonly #listeners = new Map<string, Array<(payload: unknown) => void>>();

  on<TPayload = unknown>(eventName: string, listener: (payload: TPayload) => void): Disposer {
    const list = this.#listeners.get(eventName) ?? [];
    const entry = listener as (payload: unknown) => void;
    list.push(entry);
    this.#listeners.set(eventName, list);

    let disposed = false;
    return () => {
      if (disposed) {
        return;
      }
      disposed = true;
      const current = this.#listeners.get(eventName);
      const index = current?.indexOf(entry) ?? -1;
      if (index >= 0) {
        current?.splice(index, 1);
      }
    };
  }

  emit<TPayload = unknown>(eventName: string, payload: TPayload): void {
    const list = this.#listeners.get(eventName);
    if (!list || list.length === 0) {
      return;
    }
    // Snapshot before iterating: a listener disposing itself (or another)
    // mid-emit shouldn't skip or duplicate calls for this round.
    for (const listener of [...list]) {
      try {
        listener(payload);
      } catch (error) {
        this.#reportError(eventName, error);
      }
    }
  }

  dispose(): void {
    this.#listeners.clear();
  }

  /** Forwards to the "error" event; an "error" listener that itself throws only logs, never re-enters emit(). */
  #reportError(eventName: string, error: unknown): void {
    if (eventName === "error") {
      console.error(`[cyexcel] "error" event listener itself threw`, error);
      return;
    }
    this.emit("error", { error, eventName });
  }
}
