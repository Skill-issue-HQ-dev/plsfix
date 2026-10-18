// The test-only formula evaluator the template suites compute their blocks
// with: the language subset a block uses, held to Excel's own answers for the
// cases a banker's block meets - precedence, blanks, text in a range, a lazy IF,
// IFERROR, an error travelling through a reference.
import { describe, expect, it } from "vitest";
import {
  blank,
  build,
  calc,
  input,
  inputText,
  label,
} from "../src/template-cells";
import { at, span } from "../src/templates/refs";
import { evaluateTemplate, ExcelError } from "./formula-eval";

// A one-column sheet: the formula under test sits in A1, so every case reads
// `value("A1")`; inputs below it are A2, A3 ... through at(row, 0).
function sheet(
  formula: string,
  below: ReturnType<typeof input>[] = [],
  overrides: Record<string, number> = {},
) {
  const rows = [[calc(formula, "decimal")], ...below.map((cell) => [cell])];
  return evaluateTemplate(build("t", "T", "T", rows), overrides);
}

describe("arithmetic", () => {
  it("multiplies before it adds and honours brackets", () => {
    expect(sheet("=1+2*3").value("A1")).toBe(7);
    expect(sheet("=(1+2)*3").value("A1")).toBe(9);
  });

  it("binds a minus tighter than a power, as Excel does", () => {
    expect(sheet("=-2^2").value("A1")).toBe(4);
    expect(sheet("=2^-1").value("A1")).toBe(0.5);
    expect(sheet("=2^3^2").value("A1")).toBe(64);
  });

  it("divides and reports a zero divisor as #DIV/0!", () => {
    expect(sheet("=10/4").value("A1")).toBe(2.5);
    expect(sheet("=1/0").value("A1")).toEqual(new ExcelError("#DIV/0!"));
  });
});

describe("references", () => {
  it("reads another cell's literal and computes through it", () => {
    const grid = sheet(`=${at(1, 0)}*${at(2, 0)}`, [
      input(6, "whole"),
      input(7, "whole"),
    ]);
    expect(grid.value("A1")).toBe(42);
  });

  it("takes a blank cell as zero in arithmetic", () => {
    const grid = sheet(`=${at(1, 0)}+5`, [blank()]);
    expect(grid.value("A1")).toBe(5);
  });

  it("refuses a text cell in arithmetic with #VALUE!", () => {
    const grid = sheet(`=${at(1, 0)}+5`, [label("n/a")]);
    expect(grid.value("A1")).toEqual(new ExcelError("#VALUE!"));
  });

  it("lets an error travel through a reference", () => {
    const rows = [[calc("=1/0", "whole")], [calc(`=${at(0, 0)}+1`, "whole")]];
    const grid = evaluateTemplate(build("t", "T", "T", rows));
    expect(grid.value("A2")).toEqual(new ExcelError("#DIV/0!"));
  });

  it("answers a typed override over a block's own input", () => {
    const grid = sheet(`=${at(1, 0)}*2`, [input(5, "whole")], { A2: 8 });
    expect(grid.value("A1")).toBe(16);
  });
});

describe("the functions a block uses", () => {
  const column = [
    input(4, "whole"),
    input(10, "whole"),
    blank(),
    label("x"),
    input(1, "whole"),
  ];
  const range = span(1, 0, 5, 0);

  it("sums, averages and counts only the numbers in a range", () => {
    expect(sheet(`=SUM(${range})`, column).value("A1")).toBe(15);
    expect(sheet(`=AVERAGE(${range})`, column).value("A1")).toBe(5);
    expect(sheet(`=COUNT(${range})`, column).value("A1")).toBe(3);
  });

  it("takes the median of an odd and of an even count", () => {
    expect(sheet(`=MEDIAN(${range})`, column).value("A1")).toBe(4);
    const even = [
      input(1, "whole"),
      input(2, "whole"),
      input(9, "whole"),
      input(4, "whole"),
    ];
    expect(sheet(`=MEDIAN(${span(1, 0, 4, 0)})`, even).value("A1")).toBe(3);
  });

  it("finds the high and the low, and zero when there is no number", () => {
    expect(sheet(`=MAX(${range})`, column).value("A1")).toBe(10);
    expect(sheet(`=MIN(${range})`, column).value("A1")).toBe(1);
    expect(
      sheet(`=MAX(${span(1, 0, 2, 0)})`, [label("a"), blank()]).value("A1"),
    ).toBe(0);
  });

  it("answers #DIV/0! for an average and #NUM! for a median of nothing", () => {
    const text = [label("a"), blank()];
    expect(sheet(`=AVERAGE(${span(1, 0, 2, 0)})`, text).value("A1")).toEqual(
      new ExcelError("#DIV/0!"),
    );
    expect(sheet(`=MEDIAN(${span(1, 0, 2, 0)})`, text).value("A1")).toEqual(
      new ExcelError("#NUM!"),
    );
  });

  it("evaluates only the branch of an IF it takes", () => {
    expect(sheet("=IF(1>0,5,1/0)").value("A1")).toBe(5);
    expect(sheet("=IF(1<0,1/0,7)").value("A1")).toBe(7);
  });

  it("reads a blank as zero in a comparison, so a blank divisor is skipped", () => {
    const grid = sheet(`=IF(${at(1, 0)}>0,10/${at(1, 0)},"")`, [blank()]);
    expect(grid.value("A1")).toBe("");
  });

  it("falls back from an error with IFERROR and not otherwise", () => {
    expect(sheet('=IFERROR(1/0,"n/a")').value("A1")).toBe("n/a");
    expect(sheet('=IFERROR(3,"n/a")').value("A1")).toBe(3);
  });

  it("joins conditions with AND and OR, and reads a blank cell as false", () => {
    expect(sheet("=IF(AND(1>0,2>1),1,0)").value("A1")).toBe(1);
    expect(sheet("=IF(AND(1>0,2<1),1,0)").value("A1")).toBe(0);
    expect(sheet("=IF(OR(1<0,2>1),1,0)").value("A1")).toBe(1);
    expect(sheet("=IF(OR(1<0,2<1),1,0)").value("A1")).toBe(0);
    const blankCell = sheet(`=IF(AND(${at(1, 0)}>0,1>0),1,0)`, [blank()]);
    expect(blankCell.value("A1")).toBe(0);
  });

  it("takes an absolute value", () => {
    expect(sheet("=ABS(-4.5)").value("A1")).toBe(4.5);
  });

  it("names a function it does not know instead of guessing", () => {
    expect(() => sheet("=PMT(1,2,3)").value("A1")).toThrow(/PMT/);
  });
});

describe("text and strings", () => {
  it("keeps a string result as text and ignores it in a range sum", () => {
    const rows = [
      [calc(`=SUM(${span(1, 0, 2, 0)})`, "whole")],
      [calc('=IF(1<0,1,"")', "whole")],
      [input(3, "whole")],
    ];
    const grid = evaluateTemplate(build("t", "T", "T", rows));
    expect(grid.value("A2")).toBe("");
    expect(grid.value("A1")).toBe(3);
  });

  it("holds a typed word as the word", () => {
    const grid = evaluateTemplate(
      build("t", "T", "T", [[inputText("Peer A")]]),
    );
    expect(grid.value("A1")).toBe("Peer A");
  });
});
