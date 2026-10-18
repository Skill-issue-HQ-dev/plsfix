// A stacked area with a value under zero against the fake host: the slide
// paints each series from the baseline up to its running total, which is only
// approximate once a series is negative, so the Insert row and Update all say
// so (the same two places that carry the smooth-line note). A stack of
// positive values, its 100 % form with none under zero and a plain area with
// a negative value (drawn exactly) say nothing.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { createWorkspace } from "../src/link/workspace";
import type * as LinksModule from "../src/ppt/links";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { bootPpt, memoryStore, pushChart, seedChart } from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
const LABEL = "Model: B4:F12";
const NOTE = "values under zero in a stack are drawn approximately";

function stackOf(
  kind: "area" | "stackedArea",
  rows: number[][],
  extra: Partial<ChartData> = {},
): ChartData {
  return {
    v: 1,
    kind,
    title: null,
    categories: ["A", "B", "C"],
    series: rows.map((values, i) => ({
      name: `s${String(i)}`,
      values,
      labels: values.map(String),
      colors: values.map(() => (i === 0 ? "#14213D" : "#F4A259")),
    })),
    font: "Aptos Narrow",
    ink: "#282623",
    titleColor: "#14213D",
    ...extra,
  };
}

let links: typeof LinksModule;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

async function insert(data: ChartData) {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(data, PNG);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { placed, item };
}

describe("a stacked area with a value under zero", () => {
  const negative = stackOf("stackedArea", [
    [10, 20, 30],
    [5, -5, 5],
  ]);

  it("says it is drawn approximately in the Insert row and under Update all", async () => {
    const { item, placed } = await insert(negative);
    expect(placed.note).toBe(NOTE);
    // One sentence with the label, the fragment in brackets as under Update all.
    expect(links.insertNote(placed)).toBe(` (${NOTE}).`);
    await pushChart(item, negative, PNG);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([`${LABEL} (${NOTE})`]);
  });

  it("says it of the 100 % form as well", async () => {
    const percent = stackOf(
      "stackedArea",
      [
        [10, 20, 30],
        [5, -5, 5],
      ],
      {
        v: 2,
        percent: true,
      },
    );
    const { placed } = await insert(percent);
    expect(placed.note).toBe(NOTE);
  });

  it("says nothing of a stack of values from zero up", async () => {
    const { placed } = await insert(
      stackOf("stackedArea", [
        [10, 20, 30],
        [5, 0, 5],
      ]),
    );
    expect(placed.note).toBeUndefined();
  });

  it("says nothing of a plain area with a negative value: it is drawn exactly", async () => {
    const { placed } = await insert(
      stackOf("area", [
        [10, -20, 30],
        [5, 5, 5],
      ]),
    );
    expect(placed.note).toBeUndefined();
  });
});
