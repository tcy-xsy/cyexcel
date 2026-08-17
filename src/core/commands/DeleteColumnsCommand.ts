import type { Command } from "../types";

export interface DeleteColumnsCommandParams {
  readonly start: number;
  readonly count: number;
}

export class DeleteColumnsCommand implements Command<"delete-columns", DeleteColumnsCommandParams> {
  readonly type = "delete-columns";
  readonly params: DeleteColumnsCommandParams;

  constructor(start: number, count: number) {
    this.params = Object.freeze({ start, count });
    Object.freeze(this);
  }
}
