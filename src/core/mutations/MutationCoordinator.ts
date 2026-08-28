import { createChangeSet } from "../data/changes";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { RevisionGate, type ModelRevision } from "../model";
import type {
  CommitReceipt,
  Mutation,
  MutationCoordinatorState,
  MutationPrepareContext,
  MutationPrepareSession,
  MutationPlanningResources,
  PreparedMutationParticipant,
  PreparedTransaction
} from "./types";

export class MutationCoordinator {
  readonly #gate: RevisionGate;
  #state: MutationCoordinatorState = "idle";
  #prepared: PreparedTransaction | null = null;
  #planningToken: object | null = null;
  #discardPlanning: (() => void) | null = null;

  constructor(gate = new RevisionGate()) {
    this.#gate = gate;
  }

  get revisionGate(): RevisionGate {
    return this.#gate;
  }

  get state(): MutationCoordinatorState {
    return this.#state;
  }

  begin(
    transactionId: string,
    createResources?: (baseRevision: ModelRevision) => MutationPlanningResources
  ): MutationPrepareSession {
    this.#assertIdle();
    this.#gate.assertCanCommit();
    this.#state = "preparing";
    const baseRevision = this.#gate.revision;
    const token = {};
    this.#planningToken = token;
    const participants: PreparedMutationParticipant[] = [];
    let resources: MutationPlanningResources;

    try {
      resources = createResources?.(baseRevision) ?? this.#createEmptyPlanningResources(baseRevision);
    } catch (error) {
      this.#planningToken = null;
      this.#state = "idle";
      throw error;
    }

    let active = true;
    let resourcesDisposed = false;
    const disposeResources = () => {
      if (resourcesDisposed) {
        return;
      }
      resourcesDisposed = true;
      try {
        resources.dispose();
      } catch {
        // Scratch cleanup cannot turn a completed commit into a failed one.
      }
    };
    const discardPlanning = () => {
      if (!active) {
        return;
      }
      active = false;
      this.#disposeParticipants(participants);
      disposeResources();
      if (this.#planningToken === token) {
        this.#planningToken = null;
        this.#discardPlanning = null;
        if (this.#state !== "disposed") {
          this.#state = "idle";
        }
      }
    };
    this.#discardPlanning = discardPlanning;
    const context: MutationPrepareContext = {
      baseRevision,
      model: resources.model,
      prepareData: resources.prepareData,
      roots: this.#createRootDraftRegistry()
    };

    return {
      transactionId,
      baseRevision,
      model: resources.model,
      stage: (mutation) => {
        if (!active || this.#state !== "preparing" || this.#planningToken !== token) {
          throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Mutation planning session is no longer active");
        }
        try {
          participants.push(mutation.prepare(context));
        } catch (error) {
          discardPlanning();
          throw error;
        }
      },
      seal: () => {
        if (!active || this.#state !== "preparing" || this.#planningToken !== token) {
          throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Mutation planning session cannot be sealed");
        }
        active = false;
        const effectiveParticipants = participants.filter((participant) => participant.hasEffect);
        this.#disposeParticipants(participants.filter((participant) => !participant.hasEffect));
        let preparedDisposed = false;
        const prepared: PreparedTransaction = Object.freeze({
          transactionId,
          baseRevision,
          preview: resources.model,
          hasEffect: effectiveParticipants.length > 0,
          participants: Object.freeze(effectiveParticipants),
          changes: Object.freeze(effectiveParticipants.flatMap((participant) => participant.forward.changes)),
          estimatedHistoryBytes: effectiveParticipants.reduce(
            (total, participant) => total + participant.estimatedBytes,
            0
          ),
          dispose: () => {
            if (preparedDisposed) {
              return;
            }
            preparedDisposed = true;
            disposeResources();
          }
        });
        this.#planningToken = null;
        this.#discardPlanning = null;
        this.#prepared = prepared;
        this.#state = "prepared";
        return prepared;
      },
      discard: discardPlanning
    };
  }

  prepare(transactionId: string, mutations: readonly Mutation[]): PreparedTransaction {
    const session = this.begin(transactionId);
    try {
      for (const mutation of mutations) {
        session.stage(mutation);
      }
      return session.seal();
    } catch (error) {
      session.discard();
      throw error;
    }
  }

  commit(prepared: PreparedTransaction): CommitReceipt {
    if (this.#state !== "prepared" || this.#prepared !== prepared) {
      throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Prepared transaction is no longer current");
    }
    if (this.#gate.revision !== prepared.baseRevision) {
      this.discard(prepared);
      throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Model revision changed after mutation prepare");
    }
    const started: PreparedMutationParticipant[] = [];

    let afterRevision = prepared.baseRevision;
    if (prepared.hasEffect) {
      this.#gate.assertCanCommit();
      this.#state = "committing";
      try {
        for (const participant of prepared.participants) {
          // Mark before entering commit: a throwing participant may already have written.
          started.push(participant);
          participant.commit();
        }
        afterRevision = this.#gate.commit();
      } catch (commitError) {
        const rollbackErrors: unknown[] = [];
        for (let index = started.length - 1; index >= 0; index -= 1) {
          try {
            started[index]?.rollback();
          } catch (rollbackError) {
            rollbackErrors.push(rollbackError);
          }
        }
        this.#disposeParticipants(prepared.participants);
        prepared.dispose();
        this.#prepared = null;
        if (rollbackErrors.length > 0) {
          this.#state = "faulted";
          throw new CyExcelError(
            CyExcelErrorCode.MutationRuntimeFaulted,
            `Mutation commit failed and ${rollbackErrors.length} rollback participant(s) also failed`
          );
        }
        this.#state = "idle";
        throw commitError;
      }
    }

    const changeSet = createChangeSet(afterRevision, prepared.changes);
    const patches = changeSet.changes;
    const receipt: CommitReceipt = Object.freeze({
      transactionId: prepared.transactionId,
      noop: !prepared.hasEffect,
      beforeRevision: prepared.baseRevision,
      afterRevision,
      forward: Object.freeze(prepared.participants.map((participant) => participant.forward)),
      inverse: Object.freeze([...prepared.participants].reverse().map((participant) => participant.inverse)),
      patches,
      historyBytes: prepared.estimatedHistoryBytes,
      changeSet
    });
    this.#disposeParticipants(prepared.participants);
    prepared.dispose();
    this.#prepared = null;
    this.#state = "idle";
    return receipt;
  }

  discard(prepared: PreparedTransaction): void {
    if (this.#prepared !== prepared) {
      return;
    }
    this.#disposeParticipants(prepared.participants);
    prepared.dispose();
    this.#prepared = null;
    this.#state = "idle";
  }

  execute(transactionId: string, mutations: readonly Mutation[]): CommitReceipt {
    return this.commit(this.prepare(transactionId, mutations));
  }

  dispose(): void {
    this.#discardPlanning?.();
    if (this.#prepared) {
      this.#disposeParticipants(this.#prepared.participants);
      this.#prepared.dispose();
      this.#prepared = null;
    }
    this.#state = "disposed";
  }

  #createEmptyPlanningResources(baseRevision: ModelRevision): MutationPlanningResources {
    const unsupported = () => {
      throw new CyExcelError(CyExcelErrorCode.MutationNotSupported, "No sheet transaction is attached");
    };
    return {
      model: {
        baseRevision,
        getRowCount: () => 0,
        getColumnCount: () => 0,
        getCell: unsupported
      },
      prepareData: unsupported,
      dispose: () => {}
    };
  }

  #createRootDraftRegistry(): MutationPrepareContext["roots"] {
    const roots = new Map<object, unknown>();
    return Object.freeze({
      get: <TRoot>(owner: object, fallback: () => TRoot): TRoot => {
        if (roots.has(owner)) {
          return roots.get(owner) as TRoot;
        }
        const root = fallback();
        roots.set(owner, root);
        return root;
      },
      set: <TRoot>(owner: object, root: TRoot): void => {
        roots.set(owner, root);
      }
    });
  }

  #assertIdle(): void {
    if (this.#state === "faulted") {
      throw new CyExcelError(CyExcelErrorCode.MutationRuntimeFaulted, "Mutation runtime is faulted");
    }
    if (this.#state === "disposed") {
      throw new CyExcelError(CyExcelErrorCode.MutationRuntimeDisposed, "Mutation runtime is disposed");
    }
    if (this.#state !== "idle") {
      throw new CyExcelError(
        CyExcelErrorCode.ReentrantMutation,
        `Cannot start a mutation while the coordinator is ${this.#state}`
      );
    }
  }

  #disposeParticipants(participants: readonly PreparedMutationParticipant[]): void {
    for (let index = participants.length - 1; index >= 0; index -= 1) {
      try {
        participants[index]?.dispose();
      } catch {
        // Prepared resources are best-effort cleanup; they never change live model state.
      }
    }
  }
}
