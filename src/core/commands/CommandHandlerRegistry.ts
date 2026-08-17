import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { Command, CommandContext, Disposer } from "../types";

export interface CommandHandler<TCommand extends Command = Command> {
  readonly type: TCommand["type"];
  plan(params: TCommand["params"], context: CommandContext): void;
}

export interface CommandPlanner {
  plan(command: Command, context: CommandContext): void;
}

/** Maps immutable command data to the handler that plans its mutations. */
export class CommandHandlerRegistry implements CommandPlanner {
  readonly #handlers = new Map<string, CommandHandler>();

  register<TCommand extends Command>(handler: CommandHandler<TCommand>): Disposer {
    if (this.#handlers.has(handler.type)) {
      throw new CyExcelError(CyExcelErrorCode.InvalidOptions, `Command handler "${handler.type}" already exists`);
    }
    this.#handlers.set(handler.type, handler as CommandHandler);
    return () => {
      if (this.#handlers.get(handler.type) === handler) {
        this.#handlers.delete(handler.type);
      }
    };
  }

  plan(command: Command, context: CommandContext): void {
    const handler = this.#handlers.get(command.type);
    if (!handler) {
      throw new CyExcelError(CyExcelErrorCode.InvalidOptions, `No handler is registered for command "${command.type}"`);
    }
    handler.plan(command.params, context);
  }
}
