// An area chart link against a fake PowerPoint that refuses part of the draw:
// a host that cannot rotate (the hanging triangle's turn is written after the
// shape's name, so the cleanup sweep by name finds it), and a host that refuses
// a triangle's tip, which the grouping sync opens with, so neither the new
// group nor the old group's delete lands. A failed insert leaves nothing on the
// slide; a failed refresh keeps the group it had.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { TAG_KEY } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type * as LinksModule from "../src/ppt/links";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushChart, seedChart } from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);

// Positive, negative, positive: the stretches cross the baseline, so the draw
// holds triangles that hang.
function dip(values: number[]): ChartData {
  return {
    v: 1,
    kind: "area",
    title: null,
    categories: ["A", "B", "C"],
    series: [
      {
        name: "Cash",
        values,
        labels: values.map(String),
        colors: values.map(() => "#B8902D"),
        dataLabels: false,
      },
    ],
    font: "Aptos Narrow",
    ink: "#282623",
    titleColor: "#14213D",
  };
}

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let helpers: FakePptHelpers;

// The cleanup's catch-up poll for ids that never answer waits 3 s by default;
// the swap lands on the copy of shape-ready.ts the fresh boot just loaded.
async function fresh(): Promise<void> {
  uninstallFakePpt();
  ({ links, presentation, relay, helpers } = await bootPpt());
  const ready = await import("../src/ppt/shape-ready");
  ready.setRegisterPollWait(() => Promise.resolve());
}

beforeEach(fresh);
afterEach(() => {
  uninstallFakePpt();
});

async function seed(data: ChartData) {
  const ws = await createWorkspace(memoryStore());
  return { ws, item: await seedChart(data, PNG) };
}

describe("a host that refuses a rotation", () => {
  it("leaves nothing on the slide: the turned triangle already has its name", async () => {
    const { ws, item } = await seed(dip([30, -20, 40]));
    helpers.refuseRotation();
    await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow();
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });
});

describe("a host that refuses a triangle's tip, which the grouping sync opens with", () => {
  it("leaves nothing on the slide after an insert", async () => {
    const { ws, item } = await seed(dip([30, -20, 40]));
    helpers.refuseTriangleTips();
    await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow();
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("keeps the old group on a refresh", async () => {
    const { ws, item } = await seed(dip([30, -20, 40]));
    await links.insertFromInbox(item, ws, relay);
    await pushChart(item, dip([30, -25, 40]), PNG);
    const rows = await links.listLinks(relay);
    const old = presentation.slides[0]!.shapes[0]!;
    const kept = { id: old.id, parts: old.group!.shapes.length };

    helpers.refuseTriangleTips();
    const summary = await links.updateLinks(rows, relay);

    expect(summary.failed).toBe(1);
    const slide = presentation.slides[0]!;
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]!.type).toBe("Group");
    expect(slide.shapes[0]!.tags.get(TAG_KEY)).toBe(item.token);
    expect(slide.shapes[0]!.id).toBe(kept.id);
    expect(slide.shapes[0]!.group!.shapes).toHaveLength(kept.parts);
  });
});
