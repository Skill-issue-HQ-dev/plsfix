// A chart type the host does not name. Excel for the web answers null for the
// type of a chart whose series differ (a combo; read on 05.10.2026) while the
// series still name themselves: the series decide what the chart is, and a chart
// nothing can decide keeps declining without ever printing the missing type
// ("null charts are not drawn as shapes"). Pure: no Office.js.

import { describe, expect, it } from "vitest";
import {
  decideKind,
  overlapsColumns,
  standingType,
  typeNamed,
  unsupportedType,
  type SeriesFacts,
} from "./chart-read-kind";

function fact(chartType: string, extra: Partial<SeriesFacts> = {}) {
  return {
    chartType,
    axisGroup: "Primary",
    hasDataLabels: true,
    markerStyle: "Automatic",
    ...extra,
  } satisfies SeriesFacts;
}

describe("decideKind when the host names no type for the chart", () => {
  const column = fact("ColumnClustered");
  const margin = fact("LineMarkers", { axisGroup: "Secondary" });
  const UNDECIDED = { issue: "this chart type is not drawn as shapes" };

  it.each([
    ["null", null],
    ["undefined", undefined],
    ["an empty string", ""],
  ])(
    "is a combo from the series' own types when the type is %s",
    (_what, head) => {
      expect(decideKind(head, [column, margin])).toEqual({
        kind: "combo",
        percent: false,
        types: ["column", "line"],
        percents: [false, false],
      });
    },
  );

  it("is the one kind the series share, a 100 % stack included", () => {
    expect(decideKind(null, [fact("Line"), fact("LineMarkers")])).toEqual({
      kind: "line",
      percent: false,
      types: null,
    });
    expect(
      decideKind(null, [fact("ColumnStacked100"), fact("ColumnStacked100")]),
    ).toEqual({ kind: "stackedColumn", percent: true, types: null });
  });

  it("gives a series that named nothing the type of the first one that did", () => {
    expect(
      decideKind(null, [fact(""), fact("ColumnClustered"), fact("")]),
    ).toEqual({ kind: "column", percent: false, types: null });
    expect(
      decideKind(null, [fact(""), fact("Line", { axisGroup: "Secondary" })]),
    ).toMatchObject({ kind: "combo", types: ["line", "line"] });
  });

  it("names a series type the slide has no shapes for", () => {
    expect(decideKind(null, [column, fact("Radar")])).toEqual({
      issue: "Radar series are not drawn as shapes",
    });
    expect(decideKind(null, [fact("Waterfall"), column])).toEqual({
      issue: "Waterfall series are not drawn as shapes",
    });
  });

  it("declines when no series can decide either, never printing the missing type", () => {
    for (const head of [null, undefined, ""]) {
      expect(decideKind(head, [fact(""), fact("")])).toEqual(UNDECIDED);
      expect(decideKind(head, [])).toEqual(UNDECIDED);
    }
  });
});

describe("typeNamed and standingType", () => {
  it("reads a type nobody gave as no type, never as the word null", () => {
    expect(typeNamed("ColumnClustered")).toBe("ColumnClustered");
    for (const nothing of [null, undefined, "", "  ", 7, {}]) {
      expect(typeNamed(nothing)).toBe("");
    }
  });

  it("stands the chart's own type for the chart, else the first series' that named one", () => {
    expect(standingType("ColumnClustered", [fact("Line")])).toBe(
      "ColumnClustered",
    );
    expect(standingType(null, [fact(""), fact("Line"), fact("Area")])).toBe(
      "Line",
    );
    expect(standingType("", [fact("")])).toBe("");
    expect(standingType(undefined, [])).toBe("");
  });
});

describe("unsupportedType", () => {
  it("names the type Excel gave", () => {
    expect(unsupportedType("Radar")).toBe(
      "Radar charts are not drawn as shapes",
    );
  });

  it("says what the user sees for a type nobody gave, never null, undefined or a blank name", () => {
    for (const type of ["", null, undefined]) {
      const sentence = unsupportedType(type);
      expect(sentence).toBe("this chart type is not drawn as shapes");
      expect(sentence).not.toMatch(/null|undefined/);
    }
  });
});

describe("overlapsColumns", () => {
  it("is a clustered column chart whose series fully overlap, and nothing else", () => {
    expect(overlapsColumns("ColumnClustered", 100)).toBe(true);
    expect(overlapsColumns("3DColumnClustered", 100)).toBe(true);
    expect(overlapsColumns("ColumnClustered", 0)).toBe(false);
    expect(overlapsColumns("ColumnClustered", null)).toBe(false);
    expect(overlapsColumns("BarClustered", 100)).toBe(false);
    expect(overlapsColumns("", 100)).toBe(false);
  });
});
