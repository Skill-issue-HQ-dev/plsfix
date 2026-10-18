// The doughnut layout, ring by ring: several series as concentric rings with
// the first innermost (as Excel draws them), each cut by its own total; every
// slice's label at its ring's mid-radius in the ink that reads on its colour,
// left out where the slice or the ring is too thin to hold its text; the title
// and the legend of the categories, and everything inside its box. Pure
// geometry in slide points.

import { describe, expect, it } from "vitest";
import {
  LABEL_PAD,
  LABEL_SIZE,
  MIN_SEGMENT,
  textWidth,
  type Text,
} from "./chart-shapes";
import { layoutDoughnut } from "./chart-shapes-doughnut";
import {
  BOX,
  COLORS,
  INK,
  LIGHT,
  WHITE,
  arcs,
  circle,
  donuts,
  doughnut,
  labels,
  radii,
  sweep,
} from "../test/chart-doughnut.support";

describe("layoutDoughnut: several series are concentric rings", () => {
  const rings = [
    [1, 1, 0],
    [1, 3, 0],
    [2, 1, 1],
  ];
  const data = doughnut(rings, { hole: 50 });
  const ring = (out: ReturnType<typeof arcs>, j: number) =>
    out.filter((one) => one.name.startsWith(`slice ${String(j)}.`));

  it("draws the first series innermost, as Excel does, and the rings edge to edge", () => {
    const { radius } = circle(data);
    const out = arcs(layoutDoughnut(data, BOX));
    const [inner, middle, outer] = [0, 1, 2].map((j) =>
      ring(out, j).map(radii),
    );
    // The hole is half the outer diameter; three rings share what is left.
    expect(inner![0]!.inner).toBeCloseTo(radius * 0.5, 9);
    expect(inner![0]!.outer).toBeCloseTo(radius * (0.5 + 0.5 / 3), 9);
    expect(middle![0]!.inner).toBeCloseTo(inner![0]!.outer, 9);
    expect(middle![0]!.outer).toBeCloseTo(outer![0]!.inner, 9);
    expect(outer![0]!.outer).toBeCloseTo(radius, 9);
  });

  it("cuts each ring by its own total, from 12 o'clock", () => {
    const out = arcs(layoutDoughnut(data, BOX));
    expect(ring(out, 0).map((one) => [one.start, one.end])).toEqual([
      [-90, 90],
      [90, -90],
    ]);
    expect(ring(out, 1).map(sweep)).toEqual([90, 270]);
    expect(ring(out, 2).map(sweep)).toEqual([180, 90, 90]);
    for (const j of [0, 1, 2]) {
      const total = ring(out, j).reduce((sum, one) => sum + sweep(one), 0);
      expect(Math.abs(total - 360)).toBeLessThan(0.01);
    }
  });

  it("names a slice by its ring and its category, and wears each category's colour in every ring", () => {
    const out = arcs(layoutDoughnut(data, BOX));
    expect(new Set(out.map((one) => one.name)).size).toBe(out.length);
    expect(out.find((one) => one.name === "slice 1.1")!.color).toBe(LIGHT);
    expect(out.find((one) => one.name === "slice 2.1")!.color).toBe(LIGHT);
  });

  it("draws a ring of one slice as a Donut on its own radii, beside the other rings' arcs", () => {
    const single = doughnut(
      [
        [4, 0, 0],
        [1, 1, 1],
      ],
      { hole: 50 },
    );
    const out = layoutDoughnut(single, BOX);
    expect(donuts(out)).toHaveLength(1);
    expect(donuts(out)[0]!.name).toBe("slice 0.0");
    expect(arcs(out)).toHaveLength(3);
    const { radius } = circle(single);
    expect(radii(donuts(out)[0]!).inner).toBeCloseTo(radius * 0.5, 9);
    expect(radii(donuts(out)[0]!).outer).toBeCloseTo(radius * 0.75, 9);
  });
});

describe("layoutDoughnut: the labels", () => {
  const data = doughnut([[1, 1, 2]], { hole: 50 });

  it("sits each label at its slice's mid-angle on the ring's mid-radius", () => {
    const out = layoutDoughnut(data, BOX);
    const { cx, cy } = circle(data);
    const text = labels(out);
    expect(text.map((one) => one.text)).toEqual(["1", "1", "2"]);
    expect(text.map((one) => one.name)).toEqual([
      "label 0.0",
      "label 0.1",
      "label 0.2",
    ]);
    arcs(out).forEach((arc, i) => {
      const { inner, outer } = radii(arc);
      const mid = ((arc.start + sweep(arc) / 2) * Math.PI) / 180;
      const one = text[i]!;
      const x = one.box.left + one.box.width / 2;
      const y = one.box.top + one.box.height / 2;
      expect(x).toBeCloseTo(cx + ((inner + outer) / 2) * Math.cos(mid), 6);
      expect(y).toBeCloseTo(cy + ((inner + outer) / 2) * Math.sin(mid), 6);
      expect(one.align).toBe("c");
    });
  });

  it("writes white on a dark slice and the chart's ink on a light one", () => {
    const [dark, light] = labels(layoutDoughnut(data, BOX));
    expect(dark!.color).toBe(WHITE);
    expect(light!.color).toBe(INK);
  });

  it("leaves a slice too thin to hold its label without one, and keeps the others'", () => {
    const out = layoutDoughnut(doughnut([[1000, 1, 1000]]), BOX);
    expect(arcs(out)).toHaveLength(3);
    expect(labels(out).map((one) => one.name)).toEqual([
      "label 0.0",
      "label 0.2",
    ]);
  });

  it("leaves rings too thin for a label without any, however big their slices", () => {
    const many = doughnut(
      Array.from({ length: 12 }, () => [1, 1, 1, 1]),
      { hole: 50 },
    );
    const { radius } = circle(many);
    expect((radius * 0.5) / 12).toBeLessThan(MIN_SEGMENT);
    expect(labels(layoutDoughnut(many, BOX))).toEqual([]);
  });

  it("writes a label only where its text lies whole inside its slice: every corner of the text is in the ring segment", () => {
    const wide = doughnut([[120, 480, 4_000_000, 300, 1_000_000_000]], {
      hole: 40,
    });
    const out = layoutDoughnut(wide, BOX);
    const { cx, cy } = circle(wide);
    const text = labels(out);
    expect(text.length).toBeGreaterThan(0);
    for (const one of text) {
      const index = one.name.split(".")[1]!;
      const arc = arcs(out).find((a) => a.name === `slice 0.${index}`)!;
      const { inner, outer } = radii(arc);
      const halfWidth = (textWidth(one.text, LABEL_SIZE) - LABEL_PAD) / 2;
      const halfHeight = MIN_SEGMENT / 2;
      const centreX = one.box.left + one.box.width / 2;
      const centreY = one.box.top + one.box.height / 2;
      for (const sx of [-1, 1]) {
        for (const sy of [-1, 1]) {
          const dx = centreX + sx * halfWidth - cx;
          const dy = centreY + sy * halfHeight - cy;
          const r = Math.hypot(dx, dy);
          expect(r, one.name).toBeGreaterThanOrEqual(inner - 1e-6);
          expect(r, one.name).toBeLessThanOrEqual(outer + 1e-6);
          const angle = (Math.atan2(dy, dx) * 180) / Math.PI;
          const along = (((angle - arc.start) % 360) + 360) % 360;
          expect(along, one.name).toBeLessThanOrEqual(sweep(arc) + 1e-6);
        }
      }
    }
  });

  it("writes no label for a ring whose series has data labels off, or a blank text", () => {
    const off = doughnut([[1, 1, 2]]);
    off.series[0] = { ...off.series[0]!, dataLabels: false };
    expect(labels(layoutDoughnut(off, BOX))).toEqual([]);
    const blank = doughnut([[1, 1, 2]]);
    blank.series[0] = { ...blank.series[0]!, labels: ["", "1", "2"] };
    expect(labels(layoutDoughnut(blank, BOX)).map((one) => one.name)).toEqual([
      "label 0.1",
      "label 0.2",
    ]);
  });
});

describe("layoutDoughnut: the title and the legend", () => {
  it("writes the title first, and a legend of the categories in the pie's way", () => {
    const data = doughnut([[1, 1, 2]], { title: "Mix" });
    const out = layoutDoughnut(data, BOX);
    expect(out[0]).toMatchObject({ kind: "text", name: "title", text: "Mix" });
    const words = out.filter(
      (one): one is Text =>
        one.kind === "text" && one.name.startsWith("legend "),
    );
    expect(words.map((one) => one.text)).toEqual([
      "Slice 0",
      "Slice 1",
      "Slice 2",
    ]);
    const swatches = out.filter((one) => one.name.startsWith("legend swatch"));
    expect(
      swatches.map((one) => (one.kind === "rect" ? one.color : "")),
    ).toEqual(COLORS.slice(0, 3));
  });

  it("keeps every primitive inside its box and gives each its own name, at twelve rings of 24 slices", () => {
    const data = doughnut(
      Array.from({ length: 12 }, (_one, j) =>
        Array.from({ length: 24 }, (_x, i) => 1 + ((i * 7 + j) % 5)),
      ),
      { hole: 30, title: "Twelve rings" },
    );
    const out = layoutDoughnut(data, BOX);
    for (const one of out) {
      const { left, top, width, height } = one.box;
      expect(width, one.name).toBeGreaterThanOrEqual(0);
      expect(height, one.name).toBeGreaterThanOrEqual(0);
      expect(left, one.name).toBeGreaterThanOrEqual(BOX.left - 0.05);
      expect(top, one.name).toBeGreaterThanOrEqual(BOX.top - 0.05);
      expect(left + width, one.name).toBeLessThanOrEqual(
        BOX.left + BOX.width + 0.05,
      );
      expect(top + height, one.name).toBeLessThanOrEqual(
        BOX.top + BOX.height + 0.05,
      );
    }
    expect(new Set(out.map((one) => one.name)).size).toBe(out.length);
  });
});
