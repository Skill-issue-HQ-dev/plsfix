// 100 % stacks (shares per category, the values stay absolute), the total label
// on top of a stack, stacked bars, and a full 120-point, 12-series chart: every
// primitive inside the box. Stacked columns are in chart-shapes-stack.test.ts.

import { describe, expect, it } from "vitest";
import { layoutStack } from "./chart-shapes-stack";
import type { Primitive, Rect, Text } from "./chart-shapes";
import type { ChartData, ChartSeries } from "./link/chart-model";

const BOX = { left: 100, top: 100, width: 480, height: 270 };
const base = {
  v: 1 as const,
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
  title: null,
};
const DARK = "#14213D";
const LIGHT = "#F2D16B";

function series(
  name: string,
  values: (number | null)[],
  color = DARK,
  extra: Partial<ChartSeries> = {},
): ChartSeries {
  return {
    name,
    values,
    labels: values.map((value) => (value === null ? "" : String(value))),
    colors: values.map(() => color),
    ...extra,
  };
}

function stack(
  kind: "stackedColumn" | "stackedBar",
  list: ChartSeries[],
  extra: Partial<ChartData> = {},
): ChartData {
  const points = list[0]!.values.length;
  return {
    ...base,
    kind,
    categories: Array.from({ length: points }, (_one, i) => `C${String(i)}`),
    series: list,
    ...extra,
  };
}

const bars = (out: Primitive[]) =>
  out.filter(
    (one): one is Rect => one.kind === "rect" && one.name.startsWith("bar "),
  );
const texts = (out: Primitive[], prefix: string) =>
  out.filter(
    (one): one is Text => one.kind === "text" && one.name.startsWith(prefix),
  );
const named = (out: Primitive[], name: string) =>
  out.find((one) => one.name === name)!;
const baseline = (out: Primitive[]) => named(out, "baseline");

function inside(one: Primitive): boolean {
  const b = one.box;
  return (
    Number.isFinite(b.left + b.top + b.width + b.height) &&
    b.width >= 0 &&
    b.height >= 0 &&
    b.left >= BOX.left - 0.01 &&
    b.top >= BOX.top - 0.01 &&
    b.left + b.width <= BOX.left + BOX.width + 0.01 &&
    b.top + b.height <= BOX.top + BOX.height + 0.01
  );
}

const TWO = [series("a", [60, 40, 80]), series("b", [40, 60, 20], LIGHT)];

describe("a 100 % stack", () => {
  const data = stack(
    "stackedColumn",
    [series("a", [60, 10, 50]), series("b", [40, 30, 50], LIGHT)],
    {
      percent: true,
    },
  );
  const out = layoutStack(data, BOX);

  it("makes every column the same height whatever it adds up to", () => {
    const drawn = bars(out);
    const total = (i: number) =>
      drawn[i]!.box.height + drawn[i + 3]!.box.height;
    expect(total(0)).toBeCloseTo(total(1));
    expect(total(1)).toBeCloseTo(total(2));
  });

  it("draws each segment as its share of its own category", () => {
    const drawn = bars(out);
    // 10 of 40 is 25 %, 30 of 40 is 75 %.
    expect(drawn[1]!.box.height / drawn[4]!.box.height).toBeCloseTo(10 / 30);
    expect(drawn[0]!.box.height / drawn[3]!.box.height).toBeCloseTo(60 / 40);
  });

  it("fills the plot from the baseline to its top", () => {
    const drawn = bars(out);
    const zero = baseline(out).box.top;
    const column = drawn[0]!.box.height + drawn[3]!.box.height;
    // Everything above the baseline is the column; its top is the plot's top.
    expect(zero - column).toBeLessThan(zero);
    expect(drawn[3]!.box.top + drawn[3]!.box.height).toBeCloseTo(
      drawn[0]!.box.top,
    );
  });

  it("still labels a segment with the value the sheet holds, not a percentage", () => {
    expect((named(out, "label 0.0") as Text).text).toBe("60");
    expect((named(out, "label 1.1") as Text).text).toBe("30");
  });

  it("draws a category of zeros without a bad box", () => {
    const zeros = layoutStack(
      stack(
        "stackedColumn",
        [series("a", [0, 5]), series("b", [0, 5], LIGHT)],
        { percent: true },
      ),
      BOX,
    );
    expect(zeros.every(inside)).toBe(true);
  });

  it("normalises mixed signs by the absolute sum: positives up, negatives down", () => {
    const mixed = layoutStack(
      stack(
        "stackedColumn",
        [series("up", [60, 60]), series("down", [-40, -40], LIGHT)],
        {
          percent: true,
        },
      ),
      BOX,
    );
    const [up, , down] = bars(mixed);
    const zero = baseline(mixed).box.top;
    expect(up!.box.height / down!.box.height).toBeCloseTo(60 / 40);
    expect(up!.box.top + up!.box.height).toBeCloseTo(zero);
    expect(down!.box.top).toBeCloseTo(zero);
  });

  it("ignores a total label, which a 100 % stack has no use for", () => {
    const withTotals = layoutStack(
      stack(
        "stackedColumn",
        [series("a", [60, 10]), series("b", [40, 30], LIGHT)],
        {
          percent: true,
          totals: ["100", "40"],
        },
      ),
      BOX,
    );
    expect(texts(withTotals, "total ")).toEqual([]);
  });
});

describe("a stack with totals", () => {
  const data = stack("stackedColumn", TWO, { totals: ["140", "160", "100"] });
  const out = layoutStack(data, BOX);

  it("writes each total above its column, as the sheet shows it", () => {
    const totals = texts(out, "total ");
    expect(totals.map((one) => one.text)).toEqual(["140", "160", "100"]);
    totals.forEach((one, i) => {
      const column = bars(out).filter((bar) =>
        bar.name.endsWith(`.${String(i)}`),
      );
      const top = Math.min(...column.map((bar) => bar.box.top));
      expect(one.box.top + one.box.height).toBeLessThanOrEqual(top + 0.01);
    });
  });

  it("keeps every primitive inside the box, the totals included", () => {
    expect(out.every(inside)).toBe(true);
    expect(new Set(out.map((one) => one.name)).size).toBe(out.length);
  });

  it("gives a category with no total label none", () => {
    const sparse = layoutStack(
      stack("stackedColumn", TWO, { totals: ["140", "", "100"] }),
      BOX,
    );
    expect(texts(sparse, "total ").map((one) => one.name)).toEqual([
      "total 0",
      "total 2",
    ]);
  });

  it("thins the totals once the columns are too narrow to hold them", () => {
    const many = Array.from({ length: 60 }, (_one, i) => 10 + (i % 7));
    const dense = layoutStack(
      stack("stackedColumn", [series("a", many), series("b", many, LIGHT)], {
        totals: many.map((value) => String(value * 2 * 1000)),
      }),
      BOX,
    );
    const shown = texts(dense, "total ");
    expect(shown.length).toBeGreaterThan(0);
    expect(shown.length).toBeLessThan(60);
    expect(shown[0]!.name).toBe("total 0");
    expect(dense.every(inside)).toBe(true);
  });
});

describe("a stacked bar chart", () => {
  const data = stack("stackedBar", [
    series("a", [60, 40, 80]),
    series("b", [40, 60, 20], LIGHT),
  ]);
  const out = layoutStack(data, BOX);

  it("turns the stack on its side: segments run left to right from a vertical baseline", () => {
    expect(baseline(out).box.width).toBe(0);
    const [a0, b0] = [named(out, "bar 0.0"), named(out, "bar 1.0")];
    expect(b0.box.left).toBeCloseTo(a0.box.left + a0.box.width);
    expect(a0.box.width / b0.box.width).toBeCloseTo(60 / 40);
    expect(a0.box.left).toBeCloseTo(baseline(out).box.left);
  });

  it("gives each category a row, in order from the top", () => {
    expect(named(out, "bar 0.0").box.top).toBeLessThan(
      named(out, "bar 0.1").box.top,
    );
    expect(named(out, "bar 0.1").box.top).toBeLessThan(
      named(out, "bar 0.2").box.top,
    );
    // Both segments of a category share the row.
    expect(named(out, "bar 1.0").box.top).toBeCloseTo(
      named(out, "bar 0.0").box.top,
    );
  });

  it("names the categories in the left column", () => {
    const names = texts(out, "category ");
    expect(names.map((one) => one.text)).toEqual(["C0", "C1", "C2"]);
    expect(names.every((one) => one.align === "r")).toBe(true);
  });

  it("writes a total at the end of its row", () => {
    const withTotals = layoutStack(
      { ...data, totals: ["140", "160", "100"] },
      BOX,
    );
    const totals = texts(withTotals, "total ");
    expect(totals.map((one) => one.text)).toEqual(["140", "160", "100"]);
    const end = named(withTotals, "bar 1.0").box;
    expect(totals[0]!.box.left).toBeGreaterThanOrEqual(
      end.left + end.width - 0.01,
    );
    expect(withTotals.every(inside)).toBe(true);
  });

  it("normalises a 100 % bar stack too, to the same row length", () => {
    const percent = layoutStack({ ...data, percent: true }, BOX);
    const row = (i: number) =>
      named(percent, `bar 0.${String(i)}`).box.width +
      named(percent, `bar 1.${String(i)}`).box.width;
    expect(row(0)).toBeCloseTo(row(2));
  });

  it("keeps every primitive inside the box", () => {
    expect(out.every(inside)).toBe(true);
    expect(new Set(out.map((one) => one.name)).size).toBe(out.length);
  });
});

describe("at the caps", () => {
  it("holds 120 points of 12 series inside the box, gaps and negatives included", () => {
    const points = 120;
    const list = Array.from({ length: 12 }, (_one, j) =>
      series(
        `s${String(j)}`,
        Array.from({ length: points }, (_p, i) =>
          (i + j) % 9 === 0 ? null : (i % 2 === 0 ? 1 : -1) * (j + 1),
        ),
        j % 2 === 0 ? DARK : LIGHT,
      ),
    );
    for (const kind of ["stackedColumn", "stackedBar"] as const) {
      for (const percent of [false, true]) {
        const out = layoutStack(
          stack(kind, list, percent ? { percent: true } : {}),
          BOX,
        );
        expect(out.length).toBeGreaterThan(0);
        expect(out.every(inside)).toBe(true);
      }
    }
  });
});
