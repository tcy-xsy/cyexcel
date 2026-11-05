import { expect, it } from "vitest";
import { shiftFormulaRows } from "../../../src/core/formula/references";

it("shifts relative and absolute row references but preserves strings, identifiers and functions", () => {
  expect(shiftFormulaRows('=SUM($A$1:B12)+LOG10(A1)+"A1"+"say ""B2"""', 1))
    .toBe('=SUM($A$2:B13)+LOG10(A2)+"A1"+"say ""B2"""');
  expect(shiftFormulaRows("='A1 sheet'!A1+Sheet1!B2+my_A1", 1))
    .toBe("='A1 sheet'!A2+Sheet1!B3+my_A1");
  expect(shiftFormulaRows("=A1048576", 1)).toBe("=#REF!");
  expect(shiftFormulaRows("=ABC1!A1", 1)).toBe("=ABC1!A2");
  expect(shiftFormulaRows("=A1", 0)).toBe("=A1");
});
