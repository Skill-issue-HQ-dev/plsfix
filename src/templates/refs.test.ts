// The spelling helpers the template grids build their formulas with: the same
// `{r+dr,c+dc}` placeholders template-cells.ts resolves, written by number.
import { describe, expect, it } from "vitest";
import { placeholderRefs, resolveFormula } from "../template-cells";
import { at, span } from "./refs";

// Column index 2, row index 4: the corner every resolution below lands on.
const C5 = { row: 4, column: 2 };

describe("at", () => {
  it("spells an offset from the block's corner", () => {
    expect(at(2, 3)).toBe("{r+2,c+3}");
  });

  it("drops a zero offset to the bare axis", () => {
    expect(at(0, 0)).toBe("{r,c}");
    expect(at(0, 4)).toBe("{r,c+4}");
    expect(at(5, 0)).toBe("{r+5,c}");
  });

  it("spells a negative offset with its minus", () => {
    expect(at(-1, -2)).toBe("{r-1,c-2}");
  });

  it("reads back through the grammar the writer resolves", () => {
    expect(placeholderRefs(at(3, 1))).toEqual([{ row: 3, column: 1 }]);
    expect(placeholderRefs(at(0, 0))).toEqual([{ row: 0, column: 0 }]);
    expect(resolveFormula(`=${at(1, 1)}`, C5)).toBe("=D6");
    expect(resolveFormula(`=${at(0, 0)}`, C5)).toBe("=C5");
  });
});

describe("span", () => {
  it("joins two corners into a range", () => {
    expect(span(2, 6, 7, 6)).toBe("{r+2,c+6}:{r+7,c+6}");
  });

  it("resolves to an A1 range", () => {
    expect(resolveFormula(`=SUM(${span(1, 0, 3, 0)})`, C5)).toBe("=SUM(C6:C8)");
  });
});
