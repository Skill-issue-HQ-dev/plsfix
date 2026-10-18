// A chart link whose category axis is on its ticks, against the fake PowerPoint
// and the fake relay: the payload's flag survives the validator, and an area and
// a line are drawn edge to edge across the group's width; without the flag they
// sit at the centres of their slots, as every deck already holds them.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { createWorkspace } from "../src/link/workspace";
import type * as LinksModule from "../src/ppt/links";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, seedChart } from "./ppt.support";

enableStrictLoadSemantics();

// 800 x 400 pixels are 600 x 300 points; six points make a slot of 100.
const PNG = fakePng(800, 400);
const VALUES = [30, 50, 40, 70, 60, 80];

function chart(kind: "area" | "line", extra: Partial<ChartData>): ChartData {
  return {
    v: 1,
    kind,
    title: null,
    categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    series: [
      {
        name: "s",
        values: VALUES,
        labels: VALUES.map(String),
        colors: VALUES.map(() => "#B8902D"),
        dataLabels: false,
      },
    ],
    font: "Aptos Narrow",
    ink: "#282623",
    titleColor: "#14213D",
    ...extra,
  };
}

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

const leaves = (shape: FakePptShape): FakePptShape[] =>
  shape.type === "Group" ? shape.group!.shapes.flatMap(leaves) : [shape];

// How far the fills (or the connectors) reach, from the group's left edge: the
// first point and the last one.
async function reach(data: ChartData, prefix: string): Promise<number[]> {
  const ws = await createWorkspace(memoryStore());
  await links.insertFromInbox(await seedChart(data, PNG), ws, relay);
  const group = presentation.slides[0]!.shapes.find(
    (one) => one.type === "Group",
  )!;
  const drawn = leaves(group).filter((one) => one.name.includes(`: ${prefix}`));
  expect(group.width).toBeCloseTo(600);
  return [
    Math.min(...drawn.map((one) => one.left)) - group.left,
    Math.max(...drawn.map((one) => one.left + one.width)) - group.left,
  ];
}

describe.each([
  ["area", "slope "],
  ["line", "line "],
] as const)("an inserted %s chart", (kind, prefix) => {
  it("is drawn edge to edge across the group when its axis is on the ticks", async () => {
    const [first, last] = await reach(chart(kind, { onTicks: true }), prefix);
    expect(first).toBeCloseTo(0);
    expect(last).toBeCloseTo(600);
  });

  it("keeps the centres of the slots without the flag", async () => {
    const [first, last] = await reach(chart(kind, {}), prefix);
    expect(first).toBeCloseTo(50);
    expect(last).toBeCloseTo(550);
  });
});
