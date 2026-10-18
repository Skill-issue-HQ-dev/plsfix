// The capex share line of the capex-and-depreciation block: capex over revenue,
// "n.m." where there is no revenue to divide by (a zero, a negative or text
// typed over it), and an error in revenue still showing instead of being hidden.
import { describe, expect, it } from "vitest";
import { errorsIn } from "../../test/block-grid";
import { evaluateTemplate, ExcelError } from "../../test/formula-eval";
import { capexDepreciation } from "./capex-depreciation";

// With the block's corner at A1: revenue is row 5, the share row 7, years in B to F.
const YEARS = ["B", "C", "D", "E", "F"] as const;
const block = capexDepreciation();

describe("the capex share of revenue", () => {
  it("shows capex as a share of each year's revenue", () => {
    const sheet = evaluateTemplate(block);
    // by hand: 30 / 300 = 10 %, 33 / 320 = 10.3125 %, 36 / 340 = 10.5882 %,
    // 38 / 360 = 10.5556 %, 40 / 380 = 10.5263 %.
    const shares = [0.1, 0.103125, 36 / 340, 38 / 360, 40 / 380];
    YEARS.forEach((column, index) => {
      expect(sheet.value(`${column}7`), column).toBeCloseTo(
        shares[index] as number,
        12,
      );
    });
  });

  it("shows n.m. for capex over a year with no revenue, and nothing else breaks", () => {
    for (const none of [0, -50]) {
      expect(errorsIn(block, { D5: none })).toEqual([]);
      const sheet = evaluateTemplate(block, { D5: none });
      expect(sheet.value("D7"), String(none)).toBe("n.m.");
      expect(sheet.value("C7")).toBeCloseTo(33 / 320, 12);
    }
  });

  it("shows n.m. where revenue is text, as a pasted n.a. would be, instead of #VALUE!", () => {
    // "n.a." compares greater than 0 and cannot be divided by; N() reads it as 0.
    expect(errorsIn(block, { D5: "n.a." })).toEqual([]);
    expect(evaluateTemplate(block, { D5: "n.a." }).value("D7")).toBe("n.m.");
  });

  it("lets an error typed into revenue show, not hide behind n.m.", () => {
    const na = new ExcelError("#N/A");
    const sheet = evaluateTemplate(block, { D5: na });
    expect(sheet.value("D7")).toEqual(na);
    expect(sheet.value("C7")).toBeCloseTo(33 / 320, 12);
  });
});
