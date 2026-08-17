import type { ModelRevision } from "../model";
import type { Mutation, PreparedMutationParticipant, StateMutationOptions } from "./types";

class StateMutation implements Mutation {
  readonly type: string;
  readonly changes;
  readonly #inverseChanges;
  readonly #apply;
  readonly #revert;
  readonly #estimatedBytes;

  constructor(options: StateMutationOptions) {
    this.type = options.type;
    this.changes = Object.freeze([...(options.changes ?? [])]);
    this.#inverseChanges = Object.freeze([...(options.inverseChanges ?? options.changes ?? [])]);
    this.#apply = options.apply;
    this.#revert = options.revert;
    this.#estimatedBytes = options.estimatedBytes ?? 64;
  }

  prepare(context: { readonly baseRevision: ModelRevision }): PreparedMutationParticipant {
    void context.baseRevision;
    const inverse = new StateMutation({
      type: `${this.type}:inverse`,
      changes: this.#inverseChanges,
      inverseChanges: this.changes,
      apply: this.#revert,
      revert: this.#apply,
      estimatedBytes: this.#estimatedBytes
    });
    let commitStarted = false;
    let rolledBack = false;

    return {
      hasEffect: true,
      forward: this,
      inverse,
      estimatedBytes: this.#estimatedBytes,
      commit: () => {
        commitStarted = true;
        this.#apply();
      },
      rollback: () => {
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        this.#revert();
      },
      dispose: () => {}
    };
  }
}

export function createStateMutation(options: StateMutationOptions): Mutation {
  return new StateMutation(options);
}
