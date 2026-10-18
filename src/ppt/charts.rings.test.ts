// declineReason for the two kinds C5 draws: a doughnut needs PowerPointApi 1.10
// (arcs have adjustments) and names it below; a combo of columns and lines needs
// 1.8 alone; a combo with a series the slide has no shapes for is a picture on
// every host, with that series named.

import { afterEach, describe, expect, it } from "vitest";
import { kindDecline } from "../chart-kinds";
import {
  needsNewerPowerPoint,
  NO_SHAPE_API,
  pictureNote,
  type ChartData,
} from "../link/chart-model";
import type { PicturePayload } from "../link/model";
import { bootPpt, columnChart } from "../../test/ppt.support";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
} from "../../test/fakeppt";
import { fakePng } from "../../test/fakepng";
import { chartPlan, declineReason, type ChartPlan } from "./charts";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

function planOf(chart: ChartData): ChartPlan {
  const payload: PicturePayload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png: fakePng(800, 400),
    src: {
      workbook: "Model.xlsx",
      sheet: "Model",
      ref: "B4:F12",
      anchor: "PLSFIX_LINK_00000000",
    },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
    chart,
  };
  return chartPlan(payload)!;
}

// Columns and a margin line on the secondary axis, or a second series of any
// other type beside the columns.
function combo(second: "line" | "scatter" = "line"): ChartData {
  const base = columnChart(6);
  const one = base.series[0]!;
  return {
    ...base,
    kind: "combo",
    series: [
      { ...one, type: "column" },
      { ...one, name: "Margin", type: second, axis: "secondary" },
    ],
  };
}

describe("declineReason for a doughnut and a combo", () => {
  it("draws a doughnut where the host has 1.10, and names what it needs below", async () => {
    const { helpers } = await bootPpt();
    const doughnut: ChartData = { ...columnChart(6), kind: "doughnut" };
    expect(kindDecline("doughnut")).toBeNull();
    const plan = planOf(doughnut);
    expect(plan.primitives.length).toBeGreaterThan(0);
    expect(declineReason(plan)).toBeNull();
    helpers.setSupported((_set, version) => version !== "1.10");
    expect(declineReason(plan)).toBe(needsNewerPowerPoint("doughnut"));
  });

  it("draws a combo of columns and lines on PowerPointApi 1.8 alone, which needs no adjustment", async () => {
    const { helpers } = await bootPpt();
    const plan = planOf(combo());
    expect(plan.primitives.length).toBeGreaterThan(0);
    helpers.setSupported((_set, version) => version !== "1.10");
    expect(declineReason(plan)).toBeNull();
    helpers.setSupported((_set, version) => version !== "1.8");
    expect(declineReason(plan)).toBe(NO_SHAPE_API);
  });

  it("names a combo with a series the slide has no shapes for in that series' own words, whatever the host", async () => {
    const { helpers } = await bootPpt();
    const plan = planOf(combo("scatter"));
    const sentence = pictureNote(
      "scatter series in a combo chart are not drawn as shapes",
    );
    expect(plan.primitives).toEqual([]);
    expect(declineReason(plan)).toBe(sentence);
    helpers.setSupported((_set, version) => version !== "1.8");
    expect(declineReason(plan)).toBe(sentence);
  });
});
