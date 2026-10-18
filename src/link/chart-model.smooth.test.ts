// What the slide says about a smooth line: it is drawn as straight segments
// (PowerPoint's shape API has no curve through points), and the row that
// inserted or updated the chart says so, in one sentence, only for a chart that
// has one. Pure: no Office.js.

import { describe, expect, it } from "vitest";
import { isChartData as frozenIsChartData } from "../../test/frozen/chart-model.v2.9.25";
import {
  isChartData,
  labelledNote,
  SMOOTH_STRAIGHT,
  smoothNote,
  STACK_NEGATIVE,
  type ChartData,
  type ChartSeries,
} from "./chart-model";

function series(extra: Partial<ChartSeries> = {}): ChartSeries {
  return {
    name: "s",
    values: [1, 2, 3],
    labels: ["1", "2", "3"],
    colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
    ...extra,
  };
}

function chartOf(list: ChartSeries[]): ChartData {
  return {
    v: 1,
    kind: "line",
    title: null,
    categories: ["a", "b", "c"],
    series: list,
    font: "Arial",
    ink: "#333333",
    titleColor: "#14213D",
  };
}

describe("smoothNote", () => {
  it("is the one sentence the row shows", () => {
    expect(SMOOTH_STRAIGHT).toBe("smooth lines are drawn straight");
  });

  it("is that sentence for a chart with a smooth series, however many others it has", () => {
    expect(smoothNote(chartOf([series({ smooth: true })]))).toBe(
      SMOOTH_STRAIGHT,
    );
    expect(
      smoothNote(chartOf([series(), series({ smooth: true }), series()])),
    ).toBe(SMOOTH_STRAIGHT);
  });

  it("is nothing for a chart whose lines are straight, flagged or not", () => {
    expect(smoothNote(chartOf([series()]))).toBeUndefined();
    expect(smoothNote(chartOf([series({ smooth: false })]))).toBeUndefined();
  });
});

// An older pane has no such field: it reads the line, ignores the flag and
// draws it straight without a word, so nothing about smoothing needs a new
// payload version.
describe("a smooth line reaching a pane from before the flag", () => {
  it("is read by the frozen v2.9.25 validator, flag and all", () => {
    const smooth = chartOf([series({ smooth: true })]);
    expect(frozenIsChartData(smooth)).toBe(true);
    expect(isChartData(smooth)).toBe(true);
  });
});

// Under "Update all" a drawn chart's note is a fragment, so it sits in brackets
// after the label, as in Excel's toast; a note that is a sentence follows it.
describe("labelledNote", () => {
  const LABEL = "Sales";

  it("puts the smooth note and the stack note in brackets", () => {
    expect(labelledNote(LABEL, SMOOTH_STRAIGHT)).toBe(
      "Sales (smooth lines are drawn straight)",
    );
    expect(labelledNote(LABEL, STACK_NEGATIVE)).toBe(
      "Sales (values under zero in a stack are drawn approximately)",
    );
  });

  it("leaves a picture note and a deck note after the label as they stand", () => {
    const picture = "as a picture: this PowerPoint has no shape API";
    const deck =
      "The deck's copy of this link was not saved: the store is full";
    expect(labelledNote(LABEL, picture)).toBe(`${LABEL} ${picture}`);
    expect(labelledNote(LABEL, deck)).toBe(`${LABEL} ${deck}`);
  });

  it("brackets only the fragment when a deck note was joined to it", () => {
    const joined = `${SMOOTH_STRAIGHT}. The deck's copy of this link was not saved`;
    expect(labelledNote(LABEL, joined)).toBe(
      "Sales (smooth lines are drawn straight). The deck's copy of this link was not saved",
    );
  });
});
