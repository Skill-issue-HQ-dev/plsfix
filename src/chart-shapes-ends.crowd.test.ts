// Twelve series that all end on one value, the worst case for end labels: in
// a line and in an area, in a roomy chart and in a small one, the labels that
// are drawn never overlap, never sit on the chart title and stay in the plot;
// the ones that do not fit are dropped, the first series keep theirs, and the
// legend still names every series.

import { describe, expect, it } from "vitest";
import {
  LABEL_HEIGHT,
  TITLE_BAND,
  bands,
  layoutChart,
  type Text,
} from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData } from "./link/chart-model";

const SERIES = 12;

function crowd(kind: "line" | "area"): ChartData {
  return {
    v: 1,
    kind,
    title: "Twelve series",
    categories: ["A", "B", "C"],
    series: Array.from({ length: SERIES }, (_one, i) => ({
      name: `s${String(i)}`,
      values: [i + 1, i + 2, 5],
      labels: ["", "", "5"],
      colors: ["#14213D", "#14213D", "#14213D"],
      dataLabels: false,
    })),
    font: "Arial",
    ink: "#111111",
    titleColor: "#000000",
  };
}

const BOXES: Box[] = [
  { left: 10, top: 20, width: 480, height: 270 },
  { left: 10, top: 20, width: 200, height: 120 },
];

describe.each(["line", "area"] as const)(
  "%s: twelve series on one end value",
  (kind) => {
    it.each(BOXES)(
      "draws no overlapping end labels and none on the title: %o",
      (box) => {
        const data = crowd(kind);
        const out = layoutChart(data, box);
        const { plot, title } = bands(data, box);
        const ends = out.filter(
          (one): one is Text =>
            one.kind === "text" && /^label \d+\.2$/.test(one.name),
        );
        expect(ends.length).toBeGreaterThanOrEqual(1);
        const capacity = Math.max(1, Math.floor(plot.height / LABEL_HEIGHT));
        expect(ends.length).toBeLessThanOrEqual(capacity);
        const sorted = [...ends].sort((a, b) => a.box.top - b.box.top);
        sorted.slice(1).forEach((one, k) => {
          expect(one.box.top).toBeGreaterThanOrEqual(
            sorted[k]!.box.top + LABEL_HEIGHT - 1e-6,
          );
        });
        for (const one of ends) {
          expect(one.box.top).toBeGreaterThanOrEqual(plot.top - 1e-6);
          expect(one.box.top + LABEL_HEIGHT).toBeLessThanOrEqual(
            plot.top + plot.height + 1e-6,
          );
          expect(one.box.top).toBeGreaterThanOrEqual(
            title.top + TITLE_BAND - 1e-6,
          );
        }
      },
    );

    it.each(BOXES)(
      "keeps the labels of the first series and the whole legend: %o",
      (box) => {
        const out = layoutChart(crowd(kind), box);
        const kept = out
          .filter((one) => /^label \d+\.2$/.test(one.name))
          .map((one) => Number(/^label (\d+)/.exec(one.name)![1]));
        expect(kept).toEqual(
          Array.from({ length: kept.length }, (_one, i) => i),
        );
        // Dropping end labels leaves the legend as it is: the same swatches a
        // chart whose series all show data labels (no end labels) draws.
        const data = crowd(kind);
        const labelled = layoutChart(
          {
            ...data,
            series: data.series.map((one) => ({ ...one, dataLabels: true })),
          },
          box,
        );
        const swatches = (list: typeof out) =>
          list.filter((one) => one.name.startsWith("legend swatch")).length;
        expect(swatches(out)).toBe(swatches(labelled));
      },
    );
  },
);
