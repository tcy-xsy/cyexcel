import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { AxisChangeDescriptor } from "../layout";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";

export interface PreparedStructureIntegration<TRoot, TPayload> {
  readonly nextRoot: TRoot;
  readonly forwardPayload?: TPayload;
  readonly inverseChange: AxisChangeDescriptor;
  readonly inversePayload: TPayload;
  readonly estimatedBytes?: number;
}

/** Owner bridge kept deliberately narrow: prepare is pure, publish is a root swap. */
export interface StructureIntegrationOwner<TRoot, TPayload> {
  getRoot(): TRoot;
  getPlanningRoot?(root: TRoot, context: MutationPrepareContext): TRoot;
  isCurrent?(root: TRoot): boolean;
  prepare(
    root: TRoot,
    change: AxisChangeDescriptor,
    payload: TPayload,
    context: MutationPrepareContext
  ): PreparedStructureIntegration<TRoot, TPayload>;
  replaceRoot(root: TRoot): void;
}

/**
 * Keeps the non-DataSource owners of a structural edit in the same mutation
 * journal as the data participant. All expensive resolution happens in
 * prepare(); commit()/rollback() only publish an already-built root.
 */
export class StructureIntegrationMutation<TRoot, TPayload> implements Mutation {
  readonly type: string;
  readonly changes = Object.freeze([]);

  readonly #owner: StructureIntegrationOwner<TRoot, TPayload>;
  readonly #change: AxisChangeDescriptor;
  readonly #payload: TPayload;

  constructor(
    owner: StructureIntegrationOwner<TRoot, TPayload>,
    change: AxisChangeDescriptor,
    payload: TPayload
  ) {
    this.type = `${change.axis}-structure-integration`;
    this.#owner = owner;
    this.#change = Object.freeze({ ...change, targetSpace: "pre-removal" });
    this.#payload = payload;
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const ownerKey = this.#owner as object;
    const stagedRoot = context.roots.get(ownerKey, () => this.#owner.getRoot());
    const previousRoot = this.#owner.getPlanningRoot?.(stagedRoot, context) ?? stagedRoot;
    const prepared = this.#owner.prepare(previousRoot, this.#change, this.#payload, context);
    context.roots.set(ownerKey, prepared.nextRoot);
    const forward = new StructureIntegrationMutation(
      this.#owner,
      this.#change,
      prepared.forwardPayload ?? this.#payload
    );
    const inverse = new StructureIntegrationMutation(
      this.#owner,
      prepared.inverseChange,
      prepared.inversePayload
    );
    let commitStarted = false;
    let rolledBack = false;

    return {
      hasEffect: prepared.nextRoot !== previousRoot,
      forward,
      inverse,
      estimatedBytes: prepared.estimatedBytes ?? 128,
      commit: () => {
        const current = this.#owner.getRoot();
        const isCurrent = this.#owner.isCurrent
          ? this.#owner.isCurrent(previousRoot)
          : current === previousRoot;
        if (!isCurrent) {
          throw new CyExcelError(
            CyExcelErrorCode.MutationConflict,
            "Structure root changed after mutation prepare"
          );
        }
        // replaceRoot() may publish more than one coupled owner and throw
        // after a partial swap. Mark the attempt first so coordinator
        // rollback always restores the captured pre-commit composite root.
        commitStarted = true;
        this.#owner.replaceRoot(prepared.nextRoot);
      },
      rollback: () => {
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        this.#owner.replaceRoot(previousRoot);
      },
      dispose: () => {}
    };
  }
}
