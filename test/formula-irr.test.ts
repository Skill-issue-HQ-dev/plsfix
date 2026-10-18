// The IRR the template suites' formula evaluator computes, held to the closed
// form wherever the cash flows have one, to the two answers Excel's own help
// page gives, and to the net present value it has to zero. The Returns block's
// own test leans on this being right, so it is proven apart from that block.
import { describe, expect, it } from "vitest";
import { blank, build, calc, input, label } from "../src/template-cells";
import { span } from "../src/templates/refs";
import { evaluateTemplate, ExcelError } from "./formula-eval";
import { irr } from "./formula-irr";

// The net present value at a rate, written out again here so a check on the
// solver does not lean on the solver.
const npv = (rate: number, flows: number[]): number =>
  flows.reduce((sum, flow, period) => sum + flow / (1 + rate) ** period, 0);

describe("irr", () => {
  it("finds the rate that grows one outflow into one inflow", () => {
    // by hand: 200 back on 100 in, five years later: 2.0x, so the rate r has
    // (1 + r)^5 = 2 and r = 2^(1/5) - 1 = 0.1486983550.
    expect(irr([-100, 0, 0, 0, 0, 200])).toBeCloseTo(2 ** 0.2 - 1, 12);
  });

  it("finds a negative rate when the money comes back short", () => {
    // by hand: 90 back on 100 in after three years, (1 + r)^3 = 0.9,
    // r = 0.9^(1/3) - 1 = -0.0345106154.
    expect(irr([-100, 0, 0, 90])).toBeCloseTo(0.9 ** (1 / 3) - 1, 12);
  });

  it("finds a rate above 100 %, the way a 50x outcome needs", () => {
    // by hand: 50x in five years, r = 50^(1/5) - 1 = 1.1867.
    expect(irr([-1, 0, 0, 0, 0, 50])).toBeCloseTo(50 ** 0.2 - 1, 10);
  });

  it("matches the two answers Excel's help page gives", () => {
    // Excel help, IRR: -70 000 then 12 000, 15 000, 18 000, 21 000 is -2.12 %
    // after four years and 8.66 % once 26 000 follows in year five.
    const four = irr([-70_000, 12_000, 15_000, 18_000, 21_000]);
    const five = irr([-70_000, 12_000, 15_000, 18_000, 21_000, 26_000]);
    expect(four).toBeCloseTo(-0.0212, 4);
    expect(five).toBeCloseTo(0.0866, 4);
  });

  it("zeroes the net present value of whatever it finds", () => {
    const cases = [
      [-225, 0, 0, 0, 0, 561],
      [-1000, 300, 400, 500, 250],
      [-50, -50, 40, 40, 40, 40, 40],
      // Two sign changes, two roots (10 % and 20 %): any one zeroes the NPV.
      [-100, 230, -132],
    ];
    for (const flows of cases) {
      const rate = irr(flows);
      expect(rate, flows.join(",")).not.toBeNull();
      expect(npv(rate as number, flows), flows.join(",")).toBeCloseTo(0, 8);
    }
  });

  it("gives up where Excel does: a near-total loss has no answer from the default guess", () => {
    // Excel steps from 10 % to -322 % to -24 511 % and answers #NUM!; the
    // evaluator may not do better with a search Excel does not run.
    expect(irr([-400, 0, 0, 0, 0, 40])).toBeNull();
  });

  it("finds that same loss from a guess at the root, in one step", () => {
    // by hand: 40 back on 400 in is 0.1x over five years, r = 0.1^(1/5) - 1 =
    // -0.369043; started there, the first step is already under 0.00001 %.
    const rate = irr([-400, 0, 0, 0, 0, 40], 0.1 ** 0.2 - 1);
    expect(rate).toBeCloseTo(-0.369043, 6);
  });

  it("stops after 20 tries, as Excel documents, and answers null", () => {
    // A 100 000 000x outcome needs 25 steps from the default guess, 100 000x
    // needs 17; the first is no answer, the second is, and a guess at the root
    // brings the first within reach.
    expect(irr([-1, 0, 0, 0, 0, 1e8])).toBeNull();
    expect(irr([-1, 0, 0, 0, 0, 1e5])).toBeCloseTo(1e5 ** 0.2 - 1, 8);
    expect(irr([-1, 0, 0, 0, 0, 1e8], 1e8 ** 0.2 - 1)).toBeCloseTo(
      1e8 ** 0.2 - 1,
      8,
    );
  });

  it("answers null, never a guess, when the flows never change sign", () => {
    expect(irr([100, 50])).toBeNull();
    expect(irr([-100, -5])).toBeNull();
    expect(irr([0, 0, 0])).toBeNull();
    expect(irr([])).toBeNull();
  });
});

// The same function through the evaluator: a one-row sheet, the formula in A1
// and the cash flows in B1 onward.
function row(formula: string, ...flows: ReturnType<typeof input>[]) {
  return evaluateTemplate(
    build("t", "T", "T", [[calc(formula, "percent"), ...flows]]),
  );
}

const range = span(0, 1, 0, 5);
const flows = [
  input(-100, "decimal"),
  input(0, "decimal"),
  input(0, "decimal"),
  input(0, "decimal"),
  input(200, "decimal"),
];

describe("=IRR in a formula", () => {
  it("reads a range of cash flows", () => {
    expect(row(`=IRR(${range})`, ...flows).value("A1")).toBeCloseTo(
      2 ** 0.25 - 1,
      12,
    );
  });

  it("takes the optional guess and lands on the same root", () => {
    const answer = row(`=IRR(${range},0.5)`, ...flows).value("A1");
    expect(answer).toBeCloseTo(2 ** 0.25 - 1, 12);
  });

  it("answers #NUM! for a near-total loss unless the formula guesses near the root", () => {
    const loss = [-400, 0, 0, 0, 0, 40].map((flow) => input(flow, "decimal"));
    const six = span(0, 1, 0, 6);
    expect(row(`=IRR(${six})`, ...loss).value("A1")).toEqual(
      new ExcelError("#NUM!"),
    );
    // by hand: the guess is MOIC^(1/5) - 1 with MOIC = 40 / 400 = 0.1.
    const guessed = row(`=IRR(${six},0.1^(1/5)-1)`, ...loss).value("A1");
    expect(guessed).toBeCloseTo(-0.369043, 6);
  });

  it("answers #NUM! when the row has no sign change", () => {
    const positive = [input(100, "decimal"), input(5, "decimal")];
    expect(row(`=IRR(${span(0, 1, 0, 2)})`, ...positive).value("A1")).toEqual(
      new ExcelError("#NUM!"),
    );
  });

  it("lets IFERROR turn that #NUM! into a word", () => {
    const positive = [input(100, "decimal"), input(5, "decimal")];
    const guarded = `=IFERROR(IRR(${span(0, 1, 0, 2)}),"n/a")`;
    expect(row(guarded, ...positive).value("A1")).toBe("n/a");
  });

  it("skips text and blank cells in the range, as Excel does, squeezing the years", () => {
    // by hand: -100, a note, a blank, 200 is read as -100 then 200 one period
    // apart, so r = 200/100 - 1 = 1.0 - not the 2^(1/3) - 1 of three years.
    const squeezed = row(
      `=IRR(${span(0, 1, 0, 4)})`,
      input(-100, "decimal"),
      label("note"),
      blank(),
      input(200, "decimal"),
    );
    expect(squeezed.value("A1")).toBeCloseTo(1, 12);
  });
});
