// The test evaluator held to Excel's own answers where a banker's block meets
// them: what an aggregate reads from a cell, what N, ISTEXT and PERCENTILE.INC
// return, and that a call with the wrong argument count is refused by name.
// Each expectation is what Excel's function reference states, not our guess.
import { describe, expect, it } from "vitest";
import { blank, build, calc, input, label } from "../src/template-cells";
import { at, span } from "../src/templates/refs";
import { type Computed, evaluateTemplate, ExcelError } from "./formula-eval";

// The formula under test sits in A1; the cells below it are A2, A3 ...
function sheet(
  formula: string,
  below: ReturnType<typeof input>[] = [],
  overrides: Record<string, Computed> = {},
) {
  const rows = [[calc(formula, "decimal")], ...below.map((cell) => [cell])];
  return evaluateTemplate(build("t", "T", "T", rows), overrides);
}

const DIV0 = new ExcelError("#DIV/0!");
const NUM = new ExcelError("#NUM!");
const VALUE = new ExcelError("#VALUE!");

describe("what an aggregate reads, by Excel's rules", () => {
  const four = [input(4, "whole"), blank()];
  const pair = `${at(1, 0)},${at(2, 0)}`;

  it("skips a blank, empty text and a logical handed over as one cell", () => {
    expect(sheet(`=AVERAGE(${pair})`, four).value("A1")).toBe(4);
    expect(sheet(`=AVERAGE(${pair})`, four, { A3: "" }).value("A1")).toBe(4);
    expect(sheet(`=SUM(${pair})`, four, { A3: true }).value("A1")).toBe(4);
    expect(sheet(`=MAX(${at(1, 0)},0)`, [label("n.a.")]).value("A1")).toBe(0);
    expect(sheet(`=MIN(${pair})`, four, { A3: "n.a." }).value("A1")).toBe(4);
  });

  it("answers what Excel answers when every cell it reads is skipped", () => {
    const text = [label("n.a.")];
    expect(sheet(`=AVERAGE(${at(1, 0)})`, text).value("A1")).toEqual(DIV0);
    expect(sheet(`=MEDIAN(${at(1, 0)})`, text).value("A1")).toEqual(NUM);
    expect(sheet(`=MAX(${at(1, 0)})`, text).value("A1")).toBe(0);
    const mixed = [label("x"), input(2, "whole")];
    expect(sheet(`=COUNT(${pair})`, mixed).value("A1")).toBe(1);
  });

  it("coerces a value typed into the call and refuses a word", () => {
    expect(sheet('=SUM("3",1)').value("A1")).toBe(4);
    expect(sheet('=SUM("abc",1)').value("A1")).toEqual(VALUE);
    expect(sheet('=COUNT("abc",1,"2")').value("A1")).toBe(2);
  });

  it("lets an error in a range travel, except through COUNT", () => {
    const column = [input(1, "whole"), input(2, "whole"), blank()];
    const bad = { A4: DIV0 };
    const range = span(1, 0, 3, 0);
    expect(sheet(`=SUM(${range})`, column, bad).value("A1")).toEqual(DIV0);
    expect(sheet(`=MEDIAN(${range})`, column, bad).value("A1")).toEqual(DIV0);
    expect(sheet(`=COUNT(${range})`, column, bad).value("A1")).toBe(2);
    expect(sheet(`=COUNT(${at(3, 0)})`, column, bad).value("A1")).toBe(0);
  });
});

describe("N and ISTEXT", () => {
  it("reads N as the number a cell holds and zero for the rest", () => {
    const one = (cell: ReturnType<typeof input>, over = {}) =>
      sheet(`=N(${at(1, 0)})`, [cell], over).value("A1");
    expect(one(input(5, "whole"))).toBe(5);
    expect(one(label("n.a."))).toBe(0);
    expect(one(blank())).toBe(0);
    expect(one(blank(), { A2: true })).toBe(1);
    expect(one(blank(), { A2: "3" })).toBe(0);
  });

  it("passes a genuine error through N", () => {
    const grid = sheet(`=N(${at(1, 0)})`, [blank()], {
      A2: new ExcelError("#N/A"),
    });
    expect(grid.value("A1")).toEqual(new ExcelError("#N/A"));
  });

  it("tells text from the rest, an empty result and an error included", () => {
    const one = (cell: ReturnType<typeof input>, over = {}) =>
      sheet(`=ISTEXT(${at(1, 0)})`, [cell], over).value("A1");
    expect(sheet('=ISTEXT("")').value("A1")).toBe(true);
    expect(one(label("n.a."))).toBe(true);
    expect(one(input(5, "whole"))).toBe(false);
    expect(one(blank())).toBe(false);
    expect(one(blank(), { A2: DIV0 })).toBe(false);
  });
});

describe("PERCENTILE.INC", () => {
  const four = [1, 2, 3, 4].map((n) => input(n, "whole"));
  const range = span(1, 0, 4, 0);
  const percentile = (k: string, cells = four) =>
    sheet(`=PERCENTILE.INC(${range},${k})`, cells).value("A1");

  // The Excel reference's own example: 1, 2, 3, 4 at k = 0.3 is 1.9.
  it("interpolates between the two nearest values", () => {
    expect(percentile("0.3")).toBeCloseTo(1.9, 12);
    expect(percentile("0.25")).toBe(1.75);
    expect(percentile("0.75")).toBe(3.25);
    expect(percentile("0")).toBe(1);
    expect(percentile("1")).toBe(4);
  });

  it("skips text and blanks, and answers #NUM! out of range or empty", () => {
    const mixed = [
      input(10, "whole"),
      label("n.a."),
      input(20, "whole"),
      blank(),
    ];
    expect(percentile("0.5", mixed)).toBe(15);
    expect(percentile("1.5", mixed)).toEqual(NUM);
    expect(percentile("-0.1", mixed)).toEqual(NUM);
    expect(percentile("0.5", [label("a"), blank(), blank(), blank()])).toEqual(
      NUM,
    );
  });
});

describe("^ and a reference to an empty cell, by Excel's rules", () => {
  it("raises to a power as usual, a negative base to a whole power too", () => {
    expect(sheet("=2^3").value("A1")).toBe(8);
    expect(sheet("=(-2)^3").value("A1")).toBe(-8);
    expect(sheet("=4^0.5").value("A1")).toBe(2);
  });

  it("answers #NUM! for a root of a negative number, 0^0 and an overflow", () => {
    expect(sheet("=(-8)^(1/3)").value("A1")).toEqual(NUM);
    expect(sheet("=0^0").value("A1")).toEqual(NUM);
    expect(sheet("=10^400").value("A1")).toEqual(NUM);
  });

  it("answers #DIV/0! for zero to a negative power", () => {
    expect(sheet("=0^-1").value("A1")).toEqual(DIV0);
    expect(sheet("=0^3").value("A1")).toBe(0);
  });

  it("shows 0 for a formula that is only a reference to an empty cell", () => {
    expect(sheet(`=${at(1, 0)}`, [blank()]).value("A1")).toBe(0);
    // ... and a word stays the word, a typed zero stays zero.
    expect(sheet(`=${at(1, 0)}`, [label("n.a.")]).value("A1")).toBe("n.a.");
    expect(sheet(`=${at(1, 0)}`, [input(0, "whole")]).value("A1")).toBe(0);
  });

  it("counts that formula as a number, because Excel does", () => {
    const rows = [
      [calc(`=COUNT(${span(1, 0, 2, 0)})`, "whole")],
      [calc(`=${at(2, 0)}`, "whole")],
      [blank()],
    ];
    const grid = evaluateTemplate(build("t", "T", "T", rows));
    expect(grid.value("A2")).toBe(0);
    expect(grid.value("A1")).toBe(1);
  });
});

describe("a call with the wrong number of arguments", () => {
  it("is refused by name, as Excel refuses to accept the formula", () => {
    const refused = (formula: string) => () => sheet(formula).value("A1");
    expect(refused("=IF(1>0)")).toThrow(/IF/);
    expect(refused("=IF(1>0,1,2,3)")).toThrow(/IF/);
    expect(refused("=IFERROR(1)")).toThrow(/IFERROR/);
    expect(refused("=ABS()")).toThrow(/ABS/);
    expect(refused("=N(1,2)")).toThrow(/\bN\b/);
    expect(refused("=PERCENTILE.INC(1)")).toThrow(/PERCENTILE/);
    expect(refused("=AVERAGE()")).toThrow(/AVERAGE/);
  });
});
