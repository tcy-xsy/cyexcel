export interface CyExcelCoreOptions {
  readonly id?: string;
}

export class CyExcelCore {
  readonly #options: Readonly<CyExcelCoreOptions>;
  #disposed = false;

  constructor(options: CyExcelCoreOptions = {}) {
    this.#options = Object.freeze({ ...options });
  }

  get options(): Readonly<CyExcelCoreOptions> {
    return this.#options;
  }

  get disposed(): boolean {
    return this.#disposed;
  }

  dispose(): void {
    this.#disposed = true;
  }
}

export function createCyExcelCore(options: CyExcelCoreOptions = {}): CyExcelCore {
  return new CyExcelCore(options);
}
