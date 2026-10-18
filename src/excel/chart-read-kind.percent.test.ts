// A 100 % stack and a banker's total line: decideKind says which series of a
// combo are 100 % stacks, and foldTotals refuses to fold a total line into one,
// because the folded chart would be an absolute stack the sheet never drew. The
// rest of the grammar is in chart-read-kind.test.ts.

import { describe, expect, it } from "vitest";
import {
  comboVerdict,
  decideKind,
  foldTotals,
  UNKNOWN_FACTS,
  type Decided,
  type SeriesFacts,
  type TotalCandidate,
} from "./chart-read-kind";

const fact = (chartType: string, axisGroup = "Primary"): SeriesFacts => ({
  ...UNKNOWN_FACTS,
  chartType,
  axisGroup,
});

const candidate = (
  type: TotalCandidate["type"],
  values: (number | null)[],
  extra: Partial<TotalCandidate> = {},
): TotalCandidate => ({
  name: type,
  type,
  values,
  labels: values.map((one) => (one === null ? "" : String(one))),
  dataLabels: true,
  secondary: false,
  percent: false,
  ...extra,
});

describe("decideKind: which series of a combo are 100 % stacks", () => {
  it("names them, series by series", () => {
    const decision = decideKind("ColumnStacked100", [
      fact("ColumnStacked100"),
      fact("ColumnStacked100"),
      fact("Line"),
    ]);
    expect(decision).toEqual({
      kind: "combo",
      percent: false,
      types: ["stackedColumn", "stackedColumn", "line"],
      percents: [true, true, false],
    });
  });

  it("says none of them for a plain stack with a line", () => {
    const decision = decideKind("ColumnStacked", [
      fact("ColumnStacked"),
      fact("ColumnStacked"),
      fact("Line"),
    ]);
    expect(decision).toMatchObject({
      kind: "combo",
      percents: [false, false, false],
    });
  });

  it("says nothing of it for a chart that is one kind", () => {
    const decision = decideKind("ColumnStacked100", [
      fact("ColumnStacked100"),
      fact("ColumnStacked100"),
    ]);
    expect(decision).toEqual({
      kind: "stackedColumn",
      percent: true,
      types: null,
    });
  });
});

describe("foldTotals: a total line over a 100 % stack", () => {
  const sums = [100, 120, 140];
  const stack = (percent: boolean) => [
    candidate("stackedColumn", [60, 70, 80], { percent }),
    candidate("stackedColumn", [40, 50, 60], { percent }),
  ];

  it("folds into a plain stack", () => {
    const folded = foldTotals([...stack(false), candidate("line", sums)]);
    expect(folded).toEqual({ keep: [0, 1], totals: ["100", "120", "140"] });
  });

  it("refuses a 100 % stack: the chart stays a combo", () => {
    expect(foldTotals([...stack(true), candidate("line", sums)])).toBeNull();
  });

  it("refuses a stack with a single 100 % series among plain ones", () => {
    const mixed = [
      candidate("stackedColumn", [60, 70, 80], { percent: true }),
      candidate("stackedColumn", [40, 50, 60]),
    ];
    expect(foldTotals([...mixed, candidate("line", sums)])).toBeNull();
  });
});

// Whether the combo a decision names is one the slide can draw. A 100 % stack
// with a line is the case C2 kept as a combo so it would never be drawn as an
// absolute stack; since C5 it declines in words instead, and so does every
// other mix the combo layout has no shapes for.
describe("comboVerdict: the combos the slide can and cannot draw", () => {
  const decided = (head: string, facts: SeriesFacts[]): Decided => {
    const decision = decideKind(head, facts);
    if ("issue" in decision) throw new Error(decision.issue);
    return decision;
  };

  it("is null for a chart that is not a combo", () => {
    const plain = decided("ColumnClustered", [fact("ColumnClustered")]);
    expect(comboVerdict(plain, [fact("ColumnClustered")])).toBeNull();
  });

  it("is null for clustered columns with a line, on either axis", () => {
    for (const axis of ["Primary", "Secondary"]) {
      const facts = [fact("ColumnClustered"), fact("Line", axis)];
      expect(comboVerdict(decided("ColumnClustered", facts), facts)).toBeNull();
    }
  });

  it("is null for a plain stack with a line: a combo, or the stack and its totals", () => {
    const facts = [fact("ColumnStacked"), fact("ColumnStacked"), fact("Line")];
    expect(comboVerdict(decided("ColumnStacked", facts), facts)).toBeNull();
  });

  it("names a 100 % stack with a line, so it is never drawn as an absolute stack", () => {
    const facts = [
      fact("ColumnStacked100"),
      fact("ColumnStacked100"),
      fact("Line", "Secondary"),
    ];
    expect(comboVerdict(decided("ColumnStacked100", facts), facts)).toBe(
      "a 100 % stack in a combo chart is not drawn as shapes",
    );
  });

  it("names a series the combo layout has no shapes for", () => {
    const facts = [fact("ColumnClustered"), fact("Area")];
    expect(comboVerdict(decided("ColumnClustered", facts), facts)).toBe(
      "area series in a combo chart are not drawn as shapes",
    );
    const bars = [fact("BarClustered"), fact("Line")];
    expect(comboVerdict(decided("BarClustered", bars), bars)).toBe(
      "bar series in a combo chart are not drawn as shapes",
    );
  });

  it("names columns on the secondary axis beside a primary line", () => {
    const facts = [fact("ColumnClustered", "Secondary"), fact("Line")];
    expect(comboVerdict(decided("ColumnClustered", facts), facts)).toBe(
      "columns on the secondary axis are not drawn as shapes",
    );
  });
});
