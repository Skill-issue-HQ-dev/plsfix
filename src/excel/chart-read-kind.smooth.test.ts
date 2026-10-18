// The smooth flag of a line series: written to the payload for a line series
// Excel says is smooth, and for nothing else (a column has no curve, a host
// that did not answer says nothing). Pure: no Office.js.

import { describe, expect, it } from "vitest";
import {
  seriesFlags,
  UNKNOWN_FACTS,
  type SeriesFacts,
} from "./chart-read-kind";

function fact(extra: Partial<SeriesFacts> = {}): SeriesFacts {
  return {
    ...UNKNOWN_FACTS,
    chartType: "Line",
    axisGroup: "Primary",
    hasDataLabels: true,
    markerStyle: "Automatic",
    ...extra,
  };
}

describe("seriesFlags and smooth", () => {
  it("flags a line series Excel draws as a curve", () => {
    expect(seriesFlags(fact({ smooth: true }), "line")).toEqual({
      smooth: true,
    });
  });

  it("flags nothing for a straight line, or one the host did not answer for", () => {
    expect(seriesFlags(fact({ smooth: false }), "line")).toEqual({});
    expect(seriesFlags(fact({ smooth: null }), "line")).toEqual({});
    expect(seriesFlags(fact(), "line")).toEqual({});
    expect(seriesFlags(UNKNOWN_FACTS, "line")).toEqual({});
  });

  it("flags nothing for a series that is not a line, whatever the host answered", () => {
    expect(seriesFlags(fact({ smooth: true }), "column")).toEqual({});
    expect(seriesFlags(fact({ smooth: true }), "area")).toEqual({});
  });

  it("sits beside the other flags a series carries", () => {
    expect(
      seriesFlags(
        fact({
          smooth: true,
          markerStyle: "None",
          hasDataLabels: false,
          axisGroup: "Secondary",
        }),
        "line",
      ),
    ).toEqual({
      axis: "secondary",
      dataLabels: false,
      markers: false,
      smooth: true,
    });
  });
});
