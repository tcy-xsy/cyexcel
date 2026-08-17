export type InitializationDisposer = () => void;

/** LIFO cleanup for resources acquired by a constructor before it commits. */
export class InitializationRollback {
  #active = true;
  #disposers: InitializationDisposer[] = [];

  defer(disposer: InitializationDisposer): void {
    if (!this.#active) {
      throw new Error("Initialization rollback is no longer active");
    }
    this.#disposers.push(disposer);
  }

  commit(): void {
    if (!this.#active) {
      return;
    }
    this.#active = false;
    this.#disposers = [];
  }

  rollback(): void {
    if (!this.#active) {
      return;
    }
    this.#active = false;
    const disposers = this.#disposers;
    this.#disposers = [];
    for (let index = disposers.length - 1; index >= 0; index -= 1) {
      try {
        disposers[index]?.();
      } catch {
        // Cleanup failures cannot replace the constructor's original error.
      }
    }
  }
}
