// The round trips a chart insert costs by size and host: twelve a chunk on the
// desktop, and on the web S1's chunk that shrinks as the slide fills after one
// read of the slide's population; a host slower than the model halves what is
// left. Counted by helpers.syncCount(), never timed. Split from the perf suite
// at the 400-line cap. Strict load semantics are on.

import { afterEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { createWorkspace } from "../src/link/workspace";
import { chunkFor, HOST_COSTS } from "../src/ppt/chart-cost";
import { SHAPES_PER_SYNC } from "../src/ppt/charts";
import { costedClock, type AddCosts } from "./clock";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, columnChart, memoryStore, seedChart } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";
import type { FakeRelay } from "./fakerelay";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
// The 92-shape chart on a web host 3x slower than the model: 12, 12, 6, 3 and
// then threes and ones as the cap halves again: 2 placement reads, 1 read of
// the slide's population, 45 chunks and the sync that groups.
const SLOW_WEB_SYNCS = 2 + 1 + 30 + 1;

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

afterEach(() => {
  uninstallFakePpt();
});

// The syncs a chart costs are the model's chunks, never one per shape (C1
// took the shape budget away, S1 sized the chunks): twelve at a time on the
// desktop, and on the web a chunk that shrinks as the slide fills (12, 12,
// 12, 11, 9, 8, ...) after one read of the slide's population. A slow host
// halves what is left and so affords more round trips; the progress line and
// the question add none of their own.
describe("a chart's syncs by size and host", () => {
  const PLACEMENT_SYNCS = 2;
  const POPULATION_SYNC = 1;
  const GROUP_SYNC = 1;
  const S1_WEB: Partial<AddCosts> = {
    msPerAdd: HOST_COSTS.web.perAdd * 1000,
    msPerExistingShape: HOST_COSTS.web.perExisting * 1000,
  };

  async function insertSyncsFor(
    chart: ChartData,
    costs: Partial<AddCosts>,
    platform: string,
  ): Promise<{ syncs: number; shapes: number }> {
    const clock = costedClock(costs);
    ({ links, presentation, helpers, relay } = await bootPpt({ clock }));
    helpers.setPlatform(platform);
    const { installDrawUi } = await import("../src/ppt/draw-progress");
    installDrawUi({
      progress: () => undefined,
      ask: () => Promise.resolve("draw"),
    });
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(chart, PNG);
    const from = helpers.syncCount();
    await links.insertFromInbox(item, ws, relay);
    const group = presentation.slides[0]!.shapes[0]!;
    return {
      syncs: helpers.syncCount() - from,
      shapes: group.group!.shapes.length,
    };
  }

  function wide(points: number, series: number): ChartData {
    const base = columnChart(points);
    return {
      ...base,
      series: Array.from({ length: series }, (_, j) => ({
        ...base.series[0]!,
        name: `S${String(j)}`,
      })),
    };
  }

  // The chunks the model plans for `shapes` onto an empty web slide.
  function plannedWebChunks(shapes: number): number {
    let chunks = 0;
    for (let done = 0; done < shapes; chunks += 1) {
      done += Math.min(chunkFor(done, "web"), shapes - done);
    }
    return chunks;
  }

  it("a 92-shape chart on the desktop is 8 chunks of at most twelve", async () => {
    const { syncs, shapes } = await insertSyncsFor(
      wide(12, 3),
      { msPerAdd: 100 },
      "PC",
    );
    expect(shapes).toBe(92);
    expect(
      syncs,
      "92 shapes = ceil(92 / 12) = 8 chunks, plus 2 placement reads and 1 group sync; the desktop reads no population",
    ).toBe(PLACEMENT_SYNCS + 8 + GROUP_SYNC);
  });

  it("the same chart on the web is S1's 11 chunks, after one read of the slide", async () => {
    const { syncs, shapes } = await insertSyncsFor(
      wide(12, 3),
      S1_WEB,
      "OfficeOnline",
    );
    expect(shapes).toBe(92);
    expect(
      syncs,
      "92 shapes on an empty web slide = 12, 12, 12, 11, 9, 8, 7, 6, 6, 5, 4 (11 chunks), plus 2 placement reads, 1 read of the slide's population and 1 group sync",
    ).toBe(PLACEMENT_SYNCS + POPULATION_SYNC + 11 + GROUP_SYNC);
  });

  it("the same chart on a web host three times as slow as the model costs more round trips", async () => {
    const { syncs, shapes } = await insertSyncsFor(
      wide(12, 3),
      {
        msPerAdd: HOST_COSTS.web.perAdd * 3000,
        msPerExistingShape: HOST_COSTS.web.perExisting * 3000,
      },
      "OfficeOnline",
    );
    expect(shapes).toBe(92);
    expect(
      syncs,
      "the cap halves on each slow sync and doubles back on each roomy one: more chunks than the 11 of a host that keeps up, fewer than the 45 of a cap that only fell",
    ).toBe(SLOW_WEB_SYNCS);
    expect(syncs).toBeGreaterThan(
      PLACEMENT_SYNCS + POPULATION_SYNC + 11 + GROUP_SYNC,
    );
  });

  it("a three-series monthly line is ceil(shapes / 12) chunks on the desktop, nowhere near one sync per shape", async () => {
    const line: ChartData = { ...wide(40, 3), kind: "line" };
    const { syncs, shapes } = await insertSyncsFor(
      line,
      { msPerAdd: 100 },
      "PC",
    );
    expect(shapes).toBeGreaterThan(150);
    expect(
      syncs,
      `${String(shapes)} shapes in chunks of twelve, plus 2 placement reads and 1 group sync`,
    ).toBe(PLACEMENT_SYNCS + Math.ceil(shapes / SHAPES_PER_SYNC) + GROUP_SYNC);
  });

  it("the same line on the web follows the model's chunks, far fewer syncs than shapes", async () => {
    const line: ChartData = { ...wide(40, 3), kind: "line" };
    const { syncs, shapes } = await insertSyncsFor(
      line,
      S1_WEB,
      "OfficeOnline",
    );
    const chunks = plannedWebChunks(shapes);
    expect(
      syncs,
      `${String(shapes)} shapes in ${String(chunks)} model chunks, plus 2 placement reads, 1 population read and 1 group sync`,
    ).toBe(PLACEMENT_SYNCS + POPULATION_SYNC + chunks + GROUP_SYNC);
    expect(chunks).toBeGreaterThan(Math.ceil(shapes / SHAPES_PER_SYNC));
    expect(syncs).toBeLessThan(shapes);
  });
});
