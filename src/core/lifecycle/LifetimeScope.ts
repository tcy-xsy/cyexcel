export type LifetimeState = "active" | "disposing" | "disposed";
export type LifetimeDisposer = () => void;

export interface LifetimeScopeOptions {
  readonly onDisposeError?: (error: unknown) => void;
}

/** Owns resources in LIFO order and isolates cleanup failures. */
export class LifetimeScope {
  readonly #onDisposeError: ((error: unknown) => void) | undefined;
  readonly #abortController = new AbortController();
  #state: LifetimeState = "active";
  #disposers: LifetimeDisposer[] = [];

  constructor(options: LifetimeScopeOptions = {}) {
    this.#onDisposeError = options.onDisposeError;
  }

  get state(): LifetimeState {
    return this.#state;
  }

  get signal(): AbortSignal {
    return this.#abortController.signal;
  }

  defer(disposer: LifetimeDisposer): LifetimeDisposer {
    if (this.#state !== "active") {
      disposer();
      return () => {};
    }

    let active = true;
    const tracked = () => {
      if (!active) {
        return;
      }
      active = false;
      disposer();
    };
    this.#disposers.push(tracked);
    return tracked;
  }

  fork(): LifetimeScope {
    const child = new LifetimeScope(
      this.#onDisposeError ? { onDisposeError: this.#onDisposeError } : {}
    );
    this.defer(() => child.dispose());
    return child;
  }

  dispose(): void {
    if (this.#state !== "active") {
      return;
    }

    this.#state = "disposing";
    this.#abortController.abort();
    const disposers = this.#disposers;
    this.#disposers = [];

    for (let index = disposers.length - 1; index >= 0; index -= 1) {
      try {
        disposers[index]?.();
      } catch (error) {
        try {
          this.#onDisposeError?.(error);
        } catch {
          // Error reporters are observers and cannot interrupt cleanup.
        }
      }
    }

    this.#state = "disposed";
  }
}
