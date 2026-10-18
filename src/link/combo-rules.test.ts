// Which combo charts a slide can draw, and in what roles: clustered or stacked
// columns on the primary axis, any number of lines on either axis. Everything
// else a combo can be made of - a bar, an area, a scatter, a 100 % stack,
// columns on the secondary axis, clustered and stacked columns together - is
// declined in words, by the one function both panes call, never drawn wrong.

import { describe, expect, it } from "vitest";
import type { ChartData } from "./chart-model";
import type { ChartSeries, SeriesType } from "./chart-series";
import {
  comboIssue,
  comboParts,
  comboRoles,
  type ComboPart,
} from "./combo-rules";

function part(type: SeriesType, extra: Partial<ComboPart> = {}): ComboPart {
  return { type, secondary: false, percent: false, ...extra };
}

describe("comboIssue: the combos that are drawn", () => {
  it("accepts clustered columns with a line on the primary axis", () => {
    expect(comboIssue([part("column"), part("line")])).toBeNull();
  });

  it("accepts a line on the secondary axis: the commonest banker combo", () => {
    expect(
      comboIssue([part("column"), part("line", { secondary: true })]),
    ).toBeNull();
  });

  it("accepts stacked columns with lines on both axes", () => {
    expect(
      comboIssue([
        part("stackedColumn"),
        part("stackedColumn"),
        part("line"),
        part("line", { secondary: true }),
      ]),
    ).toBeNull();
  });

  it("accepts lines alone on two axes", () => {
    expect(
      comboIssue([part("line"), part("line", { secondary: true })]),
    ).toBeNull();
  });

  it("takes a combo whose every series is on the secondary axis for the one axis it is", () => {
    expect(
      comboIssue([
        part("column", { secondary: true }),
        part("line", { secondary: true }),
      ]),
    ).toBeNull();
  });
});

describe("comboIssue: the combos that decline, in words", () => {
  it.each([
    ["bar", "bar"],
    ["stackedBar", "stacked bar"],
    ["area", "area"],
    ["stackedArea", "stacked area"],
    ["scatter", "scatter"],
    ["pie", "pie"],
    ["doughnut", "doughnut"],
  ] as const)(
    "names a %s series the slide has no combo shapes for",
    (type, words) => {
      expect(comboIssue([part("column"), part(type)])).toBe(
        `${words} series in a combo chart are not drawn as shapes`,
      );
    },
  );

  it("names the first series it cannot draw, in series order", () => {
    expect(comboIssue([part("column"), part("area"), part("scatter")])).toBe(
      "area series in a combo chart are not drawn as shapes",
    );
  });

  it("declines a 100 % stack, which a line beside it would be drawn over wrongly", () => {
    expect(
      comboIssue([
        part("stackedColumn", { percent: true }),
        part("stackedColumn", { percent: true }),
        part("line", { secondary: true }),
      ]),
    ).toBe("a 100 % stack in a combo chart is not drawn as shapes");
  });

  it("declines clustered and stacked columns in one chart", () => {
    expect(
      comboIssue([part("column"), part("stackedColumn"), part("line")]),
    ).toBe(
      "clustered and stacked columns in one combo chart are not drawn as shapes",
    );
  });

  it("declines columns on the secondary axis while another series sits on the primary one", () => {
    expect(
      comboIssue([part("column", { secondary: true }), part("line")]),
    ).toBe("columns on the secondary axis are not drawn as shapes");
  });
});

describe("comboRoles", () => {
  it("lists the columns and the lines of each axis by their place in the payload", () => {
    const roles = comboRoles([
      part("line"),
      part("column"),
      part("line", { secondary: true }),
      part("column"),
      part("line", { secondary: true }),
    ]);
    expect(roles).toEqual({
      columns: [1, 3],
      stacked: false,
      primaryLines: [0],
      secondaryLines: [2, 4],
    });
  });

  it("says stacked for stacked columns", () => {
    expect(
      comboRoles([part("stackedColumn"), part("stackedColumn"), part("line")])
        ?.stacked,
    ).toBe(true);
  });

  it("puts every series on the one axis when all of them say secondary", () => {
    expect(
      comboRoles([
        part("column", { secondary: true }),
        part("line", { secondary: true }),
      ]),
    ).toEqual({
      columns: [0],
      stacked: false,
      primaryLines: [1],
      secondaryLines: [],
    });
  });

  it("is null for a combo that declines", () => {
    expect(comboRoles([part("column"), part("area")])).toBeNull();
  });
});

describe("comboParts", () => {
  const series = (extra: Partial<ChartSeries>): ChartSeries => ({
    name: "s",
    values: [1, 2],
    labels: ["1", "2"],
    colors: ["#2EC4B6", "#2EC4B6"],
    ...extra,
  });
  const chart = (list: ChartSeries[]): ChartData => ({
    v: 1,
    kind: "combo",
    title: null,
    categories: ["a", "b"],
    series: list,
    font: "Arial",
    ink: "#111111",
    titleColor: "#000000",
  });

  it("reads each series' type and axis off the payload", () => {
    expect(
      comboParts(
        chart([
          series({ type: "column" }),
          series({ type: "line", axis: "secondary" }),
        ]),
      ),
    ).toEqual([
      { type: "column", secondary: false, percent: false },
      { type: "line", secondary: true, percent: false },
    ]);
  });

  it("gives a series that names no type the commonest one, a column", () => {
    expect(
      comboParts(chart([series({}), series({ type: "line" })]))[0]?.type,
    ).toBe("column");
  });
});
