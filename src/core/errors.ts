/**
 * Stable error codes. Callers must branch on `code`, not on `message` text —
 * message wording may change between releases without being a breaking
 * change. New codes are appended as later steps need them; existing codes
 * are never renumbered or reused for a different meaning.
 */
export const CyExcelErrorCode = {
  InvalidCoordinate: "invalid-coordinate",
  InvalidRange: "invalid-range",
  InvalidOptions: "invalid-options"
} as const;

export type CyExcelErrorCode =
  (typeof CyExcelErrorCode)[keyof typeof CyExcelErrorCode];

export class CyExcelError extends Error {
  readonly code: CyExcelErrorCode;

  constructor(code: CyExcelErrorCode, message: string) {
    super(message);
    this.name = "CyExcelError";
    this.code = code;
  }
}
