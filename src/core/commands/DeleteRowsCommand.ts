import type { Command } from "../types";

export interface DeleteRowsCommandParams {
  readonly start: number;
  readonly count: number;
}

export class DeleteRowsCommand implements Command<"delete-rows", DeleteRowsCommandParams> {
  readonly type = "delete-rows";
  readonly params: DeleteRowsCommandParams;

  constructor(start: number, count: number) {
    this.params = Object.freeze({ start, count });
    Object.freeze(this);
  }
}
