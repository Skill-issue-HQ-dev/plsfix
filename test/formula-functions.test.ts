// CHOOSE and ROUND held to Excel's reference: a pick that counts from one, reads
// only the value it picks and answers #VALUE! out of range; rounding that works
// on what a cell displays (15 significant digits), half away from zero; and a
// call with the wrong argument count refused by name instead of read around.
import { describe, expect, it } from "vitest";
import { blank, build, calc, input } from "../src/template-cells";
import { at } from "../src/templates/refs";
import { type Computed, evaluateTemplate, ExcelError } from "./formula-eval";
import { ARITY } from "./formula-functions";

// The formula under test sits in A1; the cells below it are A2, A3 ...
function sheet(
  formula: string,
  below: ReturnType<typeof input>[] = [],
  overrides: Record<string, Computed> = {},
) {
  const rows = [[calc(formula, "decimal")], ...below.map((cell) => [cell])];
  return evaluateTemplate(build("t", "T", "T", rows), overrides);
}

const refused = (formula: string) => () => sheet(formula).value("A1");
const VALUE = new ExcelError("#VALUE!");

describe("CHOOSE", () => {
  it("picks the value its index names, counting from one", () => {
    expect(sheet("=CHOOSE(1,10,20,30)").value("A1")).toBe(10);
    expect(sheet("=CHOOSE(2,10,20,30)").value("A1")).toBe(20);
    expect(sheet("=CHOOSE(3,10,20,30)").value("A1")).toBe(30);
  });

  it("reads its index and its values from cells", () => {
    const below = [
      input(2, "whole"),
      input(5, "whole"),
      input(6, "whole"),
      input(7, "whole"),
    ];
    const formula = `=CHOOSE(${at(1, 0)},${at(2, 0)},${at(3, 0)},${at(4, 0)})`;
    expect(sheet(formula, below).value("A1")).toBe(6);
  });

  it("drops the fraction of its index, as Excel does", () => {
    expect(sheet("=CHOOSE(2.9,10,20,30)").value("A1")).toBe(20);
  });

  it("answers #VALUE! for an index below one, past the last value or text", () => {
    for (const index of ["0", "-1", "4", '"x"']) {
      expect(sheet(`=CHOOSE(${index},10,20,30)`).value("A1"), index).toEqual(
        VALUE,
      );
    }
  });

  it("reads a blank index as zero, which is #VALUE!", () => {
    const grid = sheet(`=CHOOSE(${at(1, 0)},10,20,30)`, [blank()]);
    expect(grid.value("A1")).toEqual(VALUE);
  });

  it("reads only the value it picks, so an error elsewhere in the list stays hidden", () => {
    expect(sheet("=CHOOSE(1,5,1/0)").value("A1")).toBe(5);
    expect(sheet("=CHOOSE(2,5,1/0)").value("A1")).toEqual(
      new ExcelError("#DIV/0!"),
    );
  });

  it("hands text back as text", () => {
    const names = '"Base","Upside","Downside"';
    expect(sheet(`=CHOOSE(3,${names})`).value("A1")).toBe("Downside");
  });

  it("is refused with no value to pick from, instead of counting its arguments", () => {
    expect(refused("=CHOOSE(1)")).toThrow(/CHOOSE takes 2 to 255/);
  });
});

describe("ROUND", () => {
  it("keeps the digits it is given and rounds half away from zero", () => {
    expect(sheet("=ROUND(1234.5678,2)").value("A1")).toBe(1234.57);
    expect(sheet("=ROUND(2.5,0)").value("A1")).toBe(3);
    expect(sheet("=ROUND(-2.5,0)").value("A1")).toBe(-3);
    expect(sheet("=ROUND(1234.5678,-2)").value("A1")).toBe(1200);
  });

  it("rounds what Excel displays, not the binary value, at the half", () => {
    // Excel reads a number as 15 significant digits: 1.005 is 1.005, so to two
    // places it is 1.01, where the double 1.00499999999999989... would give 1.
    expect(sheet("=ROUND(1.005,2)").value("A1")).toBe(1.01);
    expect(sheet("=ROUND(0.285,2)").value("A1")).toBe(0.29);
    expect(sheet("=ROUND(-1.005,2)").value("A1")).toBe(-1.01);
  });

  it("wipes out the noise a floating-point difference leaves, with no minus zero", () => {
    // In floating point 0.1+0.2-0.3 is +5.6e-17 and 0.3-0.1-0.2 is -2.8e-17.
    for (const noise of ["0.1+0.2-0.3", "0.3-0.1-0.2"]) {
      const rounded = sheet(`=ROUND(${noise},6)`).value("A1");
      expect(Object.is(rounded, 0), noise).toBe(true);
    }
  });

  it("leaves a real difference alone", () => {
    expect(sheet("=ROUND(0.25-0.1,6)").value("A1")).toBe(0.15);
  });

  it("is refused without its digits, instead of rounding to none", () => {
    expect(refused("=ROUND(1)")).toThrow(/ROUND takes 2 to 2/);
    expect(refused("=ROUND(1,2,3)")).toThrow(/ROUND/);
  });
});

describe("the function table", () => {
  it("has an answer for every function it lists, so none falls into COUNT", () => {
    for (const [name, [fewest]] of Object.entries(ARITY)) {
      const args = Array.from({ length: fewest }, () => "1").join(",");
      expect(() => sheet(`=${name}(${args})`).value("A1"), name).not.toThrow();
    }
  });

  it("refuses a function it does not list, by name", () => {
    expect(refused("=CHOSE(1,2)")).toThrow(/does not know the function CHOSE/);
  });
});
