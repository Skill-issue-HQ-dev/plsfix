// A chart plan never throws, and the box a rebuild draws into: chartPlan gives
// a chart no layout can handle no primitives and a reason, so it is a picture
// with that reason and not an error out of Update all (where it used to escape
// the upgrade check and reject the whole deck); chartBoxAt is the picture's own
// box, or for a group the box whose drawing stands where the group stands.

import { afterEach, describe, expect, it } from "vitest";
import { frameFor } from "../chart-frame";
import type { Box } from "../layout";
import { pictureNote, type ChartData } from "../link/chart-model";
import type { PicturePayload } from "../link/model";
import { chartOf } from "../../test/hunt/ppt.chart-of.hunt3.support";
import { fakePng } from "../../test/fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
} from "../../test/fakeppt";
import { bootPpt } from "../../test/ppt.support";
import { boxAtFound, chartBoxAt } from "./chart-box";
import { chartPlan, declineReason } from "./charts";
import type { FoundLink } from "./host";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

function payloadOf(chart: ChartData): PicturePayload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 640,
    height: 360,
    png: fakePng(640, 360),
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
}

describe("chartPlan", () => {
  it("gives a waterfall too short to bridge no primitives and the kind's own reason", async () => {
    await bootPpt();
    const plan = chartPlan(payloadOf(chartOf("waterfall", 2)))!;
    expect(plan.primitives).toEqual([]);
    expect(plan.refused).toBeUndefined();
    expect(declineReason(plan)).toBe(
      pictureNote("2 points; waterfall shapes need at least 3"),
    );
  });

  it("gives a chart whose layout throws no primitives and the reason, never the error", async () => {
    await bootPpt();
    const column = chartOf("column", 6, 1);
    const broken = {
      ...column,
      series: [{ ...column.series[0]!, values: undefined }],
    } as unknown as ChartData;

    const plan = chartPlan(payloadOf(broken))!;

    expect(plan.primitives).toEqual([]);
    expect(plan.refused).toMatch(/^the chart could not be laid out \(.+\)$/);
    expect(declineReason(plan)).toBe(pictureNote(plan.refused!));
  });

  it("is drawn as before for a chart that lays out", async () => {
    await bootPpt();
    const plan = chartPlan(payloadOf(chartOf("column", 6, 2)))!;
    expect(plan.primitives.length).toBeGreaterThan(0);
    expect(plan.refused).toBeUndefined();
    expect(declineReason(plan)).toBeNull();
  });
});

function found(type: string, box: Box): FoundLink {
  return { type, ...box } as unknown as FoundLink;
}

describe("chartBoxAt", () => {
  const data = chartOf("column", 6, 2, { title: null });
  const size = { width: 360, height: 202.5 };
  const at = { left: 300, top: 200, width: 360, height: 180 };

  it("is the picture's own corner and width for a picture, and the height its aspect asks for", () => {
    expect(chartBoxAt(found("GeometricShape", at), data, size)).toEqual({
      left: 300,
      top: 200,
      width: 360,
      height: 203,
    });
    expect(boxAtFound(found("Table", at), size).height).toBe(203);
  });

  it("is derived from the group's reach for a group", () => {
    const frame = chartBoxAt(found("Group", at), data, size);
    expect(frame).toEqual(frameFor(data, at, 202.5 / 360, 360));
    expect(frame.top).toBeCloseTo(200 - 21, 6);
    expect(frame.left).toBeCloseTo(300, 6);
  });
});
