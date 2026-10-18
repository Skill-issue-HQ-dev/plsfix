// The helper the operating blocks' suites share, held to what it promises: a
// block cell by its A1 address, every error value a block computes and the static
// walk that finds a dependency loop. Reading every cell only proves no loop along
// the path the values take; the walk reads the formulas, so a loop behind a branch
// that is never taken cannot hide, which the planted blocks below prove.
import { describe, expect, it } from "vitest";
import { build, calc, input, label } from "../src/template-cells";
import { at } from "../src/templates/refs";
import { cellAt, errorsIn, findLoop } from "./block-grid";
import { ExcelError } from "./formula-eval";

// A1 = "Name", B1 = 4; A2 = "Half", B2 = B1 / 2. A placeholder is an offset from
// the block's corner, so B1 is `at(0, 1)` wherever the formula sits.
const small = build("t", "T", "T", [
  [label("Name"), input(4, "whole")],
  [label("Half"), calc(`=${at(0, 1)}/2`, "whole")],
]);

// A column of formulas, A1 downwards; at(1, 0) is A2 whichever cell holds it.
const column = (...formulas: string[]) =>
  build(
    "loop",
    "Loop",
    "A planted loop.",
    formulas.map((formula) => [calc(formula, "decimal")]),
  );

describe("cellAt", () => {
  it("finds a cell by its A1 address, the block's corner being A1", () => {
    expect(cellAt(small, "A1")).toMatchObject({ v: "Name", kind: "label" });
    expect(cellAt(small, "B1")).toMatchObject({ v: 4, kind: "input" });
    expect(cellAt(small, "A2")?.v).toBe("Half");
    expect(cellAt(small, "B2")?.f).toBe("={r,c+1}/2");
  });

  it("answers undefined past the block's edge instead of another cell", () => {
    expect(cellAt(small, "C1")).toBeUndefined();
    expect(cellAt(small, "A3")).toBeUndefined();
  });

  it("refuses text that is no address", () => {
    expect(() => cellAt(small, "the corner")).toThrow(/cell reference/);
  });
});

describe("errorsIn", () => {
  it("is empty for a block that computes cleanly", () => {
    expect(errorsIn(small)).toEqual([]);
  });

  it("names every cell that shows an error, with the error", () => {
    expect(errorsIn(small, { B1: "x" })).toEqual(["B2 #VALUE!"]);
    const broken = build("t", "T", "T", [
      [calc("=1/0", "whole")],
      [calc(`=${at(0, 0)}+1`, "whole")],
    ]);
    expect(errorsIn(broken)).toEqual(["A1 #DIV/0!", "A2 #DIV/0!"]);
  });

  it("computes with the typed overrides it is given", () => {
    expect(errorsIn(small, { B1: 10 })).toEqual([]);
  });

  it("takes an error value as an override, the way a cell can hold one", () => {
    const na = new ExcelError("#N/A");
    expect(errorsIn(small, { B1: na })).toEqual(["B1 #N/A", "B2 #N/A"]);
  });

  it("throws on a circular reference the values walk into, rather than skipping the cell", () => {
    const loop = build("t", "T", "T", [
      [calc(`=${at(1, 0)}+1`, "whole")],
      [calc(`=${at(0, 0)}+1`, "whole")],
    ]);
    expect(() => errorsIn(loop)).toThrow(/circular reference/);
  });
});

describe("findLoop", () => {
  it("finds two formulas that read each other", () => {
    // A1 = A2 + 1 and A2 = A1 * 2.
    const loop = findLoop(column(`=${at(1, 0)}+1`, `=${at(0, 0)}*2`));
    expect(loop).toEqual(["A1", "A2", "A1"]);
  });

  it("finds a formula that reads itself", () => {
    expect(findLoop(column(`=${at(0, 0)}+1`))).toEqual(["A1", "A1"]);
  });

  it("finds a loop of three, however far round it runs", () => {
    // A1 -> A2 -> A3 -> A1.
    const rows = [`=${at(1, 0)}`, `=${at(2, 0)}`, `=${at(0, 0)}`];
    expect(findLoop(column(...rows))).toEqual(["A1", "A2", "A3", "A1"]);
  });

  it("finds a loop that runs through the middle of a range", () => {
    // A3 = SUM(A1:A5) sums itself, though neither end of the range is A3.
    const rows = ["=1", "=2", `=SUM(${at(0, 0)}:${at(4, 0)})`, "=3", "=4"];
    expect(findLoop(column(...rows))).toEqual(["A3", "A3"]);
  });

  it("lets a chain and a shared precedent through", () => {
    // A1 reads A3 and A2, A2 reads A3, A3 reads A4: a diamond is not a loop.
    const rows = [
      `=${at(2, 0)}+${at(1, 0)}`,
      `=${at(2, 0)}`,
      `=${at(3, 0)}`,
      "=5",
    ];
    expect(findLoop(column(...rows))).toEqual([]);
  });

  it("reads a formula with no precedent as the end of the walk", () => {
    expect(findLoop(column("=1+1"))).toEqual([]);
  });

  it("reads no loop into a block that is only labels and typed numbers", () => {
    expect(findLoop(small)).toEqual([]);
  });
});

// Each block below computes without a complaint, because the branch that closes
// its loop is never taken: the values never walk it, the formulas still do.
describe("a loop behind a branch the values never take", () => {
  const hidden: [string, string[]][] = [
    ["an IF that is false", [`=IF(1>2,${at(1, 0)},1)`, `=${at(0, 0)}+1`]],
    [
      "a CHOOSE value that is not picked",
      [`=CHOOSE(1,5,${at(1, 0)})`, `=${at(0, 0)}+1`],
    ],
    [
      "an IFERROR fallback that is not needed",
      [`=IFERROR(1,${at(1, 0)})`, `=${at(0, 0)}+1`],
    ],
  ];

  it.each(hidden)(
    "%s: the values compute, the walk names the loop",
    (_name, formulas) => {
      const block = column(...formulas);
      expect(errorsIn(block)).toEqual([]);
      expect(findLoop(block)).toEqual(["A1", "A2", "A1"]);
    },
  );
});
