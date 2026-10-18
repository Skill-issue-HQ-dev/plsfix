// A chart is drawn in chunks the cost model sizes BEFORE each sync, and says
// how far it is. On the web, priced with S1's coefficients on the fake's costed
// clock, the 92-shape column chart goes 12, 12, 12, 11, 9, 8, 7, 6, 6, 5, 4
// onto an empty slide; a host slower than the model halves the cap for the
// rest of the chart and never lets it grow back; a draw onto a slide that
// already holds shapes - an Update all with the old group still there -
// starts smaller. Every chunk reports through the sink the pane installs.
// Strict load semantics are on.

import { afterEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { createWorkspace } from "../src/link/workspace";
import { chunkFor, HOST_COSTS } from "../src/ppt/chart-cost";
import { costedClock, type AddCosts, type CostedClock } from "./clock";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushChart,
  seedChart,
} from "./ppt.support";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

const PNG = fakePng(800, 400);
const LABEL = "Model: B4:F12";
const TAB = "Keep this tab in front.";

// S1's coefficients, in the fake's milliseconds: a host exactly as slow as the
// model says (the sync's own 0.6 s is not an add, so no chunk is ever "slow").
const S1_WEB: AddCosts = {
  msPerAdd: HOST_COSTS.web.perAdd * 1000,
  msPerExistingShape: HOST_COSTS.web.perExisting * 1000,
};
// The same host three times as slow as the model.
const THREE_TIMES: AddCosts = {
  msPerAdd: S1_WEB.msPerAdd * 3,
  msPerExistingShape: S1_WEB.msPerExistingShape * 3,
};

// Three series over twelve points: 92 shapes, the typical banker chart.
const WIDE: ChartData = {
  ...columnChart(12),
  series: [0, 1, 2].map((j) => {
    const values = Array.from({ length: 12 }, (_, i) => 10 + i + j);
    return {
      name: `S${String(j)}`,
      values,
      labels: values.map(String),
      colors: values.map(() => "#B27E54"),
    };
  }),
};
const WIDE_SHAPES = 92;

// An 8-slice pie with a title: 33 shapes.
const PIE: ChartData = (() => {
  const values = [8, 7, 6, 5, 4, 3, 2, 1];
  return {
    ...columnChart(8),
    kind: "pie",
    categories: values.map((_, i) => `Region ${String(i)}`),
    series: [
      {
        name: "Mix",
        values,
        labels: values.map(String),
        colors: values.map(() => "#14213D"),
      },
    ],
  };
})();
const PIE_SHAPES = 33;

// A 10-bar bridge with a title: 41 shapes.
const WATERFALL: ChartData = (() => {
  const values = [1000, 120, -60, 45, -80, -35, 20, -15, 30, 1025];
  return {
    ...columnChart(10),
    kind: "waterfall",
    categories: [
      "Opening",
      "Volume",
      "Price",
      "Mix",
      "COGS",
      "Opex",
      "FX",
      "Other",
      "One-offs",
      "Closing",
    ],
    series: [
      {
        name: "Bridge",
        values,
        labels: values.map(String),
        colors: values.map(() => "#B27E54"),
      },
    ],
  };
})();
const WATERFALL_SHAPES = 41;

interface Booted {
  clock: CostedClock;
  lines: string[];
  insert: (data?: ChartData) => Promise<FakePptShape>;
  update: (data: ChartData) => Promise<void>;
  syncs: () => number;
  slideShapes: () => FakePptShape[];
}

// A fresh deck whose host charges `costs` per add, the progress sink recording
// every line, and a helper that inserts a chart and hands back what landed.
async function boot(
  costs: Partial<AddCosts>,
  platform = "OfficeOnline",
): Promise<Booted> {
  const clock = costedClock(costs);
  const { links, presentation, helpers, relay } = await bootPpt({ clock });
  helpers.setPlatform(platform);
  const lines: string[] = [];
  const { installDrawUi } = await import("../src/ppt/draw-progress");
  installDrawUi({ progress: (message) => lines.push(message) });
  const ws = await createWorkspace(memoryStore());
  let item: Awaited<ReturnType<typeof seedChart>> | undefined;
  return {
    clock,
    lines,
    syncs: () => helpers.syncCount(),
    slideShapes: () => presentation.slides[0]!.shapes,
    insert: async (data = WIDE) => {
      item = await seedChart(data, PNG);
      await links.insertFromInbox(item, ws, relay);
      return presentation.slides[0]!.shapes[0]!;
    },
    update: async (data) => {
      await pushChart(item!, data, PNG);
      await links.updateLinks(await links.listLinks(relay), relay);
    },
  };
}

function line(done: number, total = WIDE_SHAPES, tab = true): string {
  const base = `Drawing ${LABEL}: ${String(done)} of ${String(total)} shapes.`;
  return tab ? `${base} ${TAB}` : base;
}

// The shapes confirmed after each progress line, and the chunk each one was.
function dones(lines: string[]): number[] {
  return lines.map((one) => Number(/: (\d+) of \d+ shapes/.exec(one)![1]));
}
function chunks(lines: string[]): number[] {
  const confirmed = dones(lines);
  return confirmed.slice(1).map((done, i) => done - confirmed[i]!);
}

describe("a web host that keeps up (S1's numbers)", () => {
  it("draws the 92-shape column chart in S1's eleven chunks and reports every one", async () => {
    const b = await boot(S1_WEB);
    const before = b.syncs();
    const drawn = await b.insert();

    expect(drawn.type).toBe("Group");
    expect(drawn.group!.shapes).toHaveLength(WIDE_SHAPES);
    expect(chunks(b.lines)).toEqual([12, 12, 12, 11, 9, 8, 7, 6, 6, 5, 4]);
    expect(dones(b.lines)).toEqual([
      0, 12, 24, 36, 47, 56, 64, 71, 77, 83, 88, 92,
    ]);
    expect(b.lines).toEqual(dones(b.lines).map((done) => line(done)));
    // Two to place, one to count the slide, the eleven chunks, one to group.
    expect(b.syncs() - before).toBe(2 + 1 + 11 + 1);
  });

  it("draws the 8-slice pie in three chunks: 12, 12 and 9", async () => {
    const b = await boot(S1_WEB);
    const drawn = await b.insert(PIE);

    expect(drawn.group!.shapes).toHaveLength(PIE_SHAPES);
    expect(chunks(b.lines)).toEqual([12, 12, 9]);
  });

  it("draws the 10-bar waterfall in four chunks: 12, 12, 12 and 5", async () => {
    const b = await boot(S1_WEB);
    const drawn = await b.insert(WATERFALL);

    expect(drawn.group!.shapes).toHaveLength(WATERFALL_SHAPES);
    expect(chunks(b.lines)).toEqual([12, 12, 12, 5]);
  });

  it("sends the same chunks whatever the host's own pace, as long as it keeps up", async () => {
    const quick = await boot({ msPerAdd: 1 });
    await quick.insert();
    expect(chunks(quick.lines)).toEqual([12, 12, 12, 11, 9, 8, 7, 6, 6, 5, 4]);
  });

  it("leaves the tab sentence off on the desktop, which sends twelve at a time", async () => {
    const b = await boot({ msPerAdd: 100 }, "PC");
    const before = b.syncs();
    await b.insert();

    expect(chunks(b.lines)).toEqual([12, 12, 12, 12, 12, 12, 12, 8]);
    expect(b.lines[0]).toBe(line(0, WIDE_SHAPES, false));
    expect(b.lines.at(-1)).toBe(line(92, WIDE_SHAPES, false));
    expect(b.lines.some((one) => one.includes(TAB))).toBe(false);
    // The desktop does not pay for the read of the slide: 2 + 8 chunks + 1.
    expect(b.syncs() - before).toBe(2 + 8 + 1);
  });
});

describe("a host that is slower than the model says", () => {
  it("halves the cap after a sync over 1.5 times its prediction and over 12 s, and doubles it back after a roomy one", async () => {
    const b = await boot(THREE_TIMES);
    const drawn = await b.insert();

    expect(drawn.group!.shapes).toHaveLength(WIDE_SHAPES);
    const sizes = chunks(b.lines);
    // Chunk one took 10 s (under 12 s: no halving), chunk two 19 s (over both
    // lines: the cap drops to 6), chunk three 12.8 s (drops to 3).
    expect(sizes.slice(0, 4)).toEqual([12, 12, 6, 3]);
    expect(Math.max(...sizes.slice(3))).toBe(3);
    // The slide fills, the host stays slow: the cap reaches one add a sync.
    expect(sizes.at(-1)).toBe(1);
    expect(sizes.reduce((sum, one) => sum + one, 0)).toBe(WIDE_SHAPES);
    // It is not stuck there: a sync back in under half the 12 s target doubles
    // the cap again (rig 05.10: a cap that only halved drew one shape a sync).
    expect(sizes.slice(3).some((one, i) => one > sizes[i + 2]!)).toBe(true);
  });

  it("grows back: a fast sync after a slow one doubles the halved cap, as far as the model allows", async () => {
    // Slow for the first two chunks, then the host is repriced to nothing.
    const b = await boot(THREE_TIMES);
    let reported = 0;
    const { installDrawUi } = await import("../src/ppt/draw-progress");
    installDrawUi({
      progress: (message) => {
        b.lines.push(message);
        reported += 1;
        if (reported === 3) b.clock.costs.msPerAdd = 0;
        if (reported === 3) b.clock.costs.msPerExistingShape = 0;
      },
    });
    await b.insert();

    const sizes = chunks(b.lines);
    // The host recovers after chunk three: the cap doubles from 6 to 12 and the
    // model (the shapes already on the slide) sizes the chunks from there.
    expect(sizes.slice(0, 5)).toEqual([12, 12, 6, 12, 10]);
    expect(Math.max(...sizes.slice(3))).toBe(12);
    expect(b.lines.at(-1)).toBe(line(WIDE_SHAPES));
  });

  it("still gives every pie wedge its angles across chunks that shrank", async () => {
    const b = await boot(THREE_TIMES);
    const drawn = await b.insert(PIE);

    expect(chunks(b.lines).slice(0, 3)).toEqual([12, 12, 6]);
    const wedges = drawn.group!.shapes.filter(
      (shape) => shape.geometry === "Pie",
    );
    expect(wedges).toHaveLength(8);
    expect(wedges.every((wedge) => wedge.adjustments?.length === 2)).toBe(true);
    // Adjustments only land after the shape's first sync, so a wedge in the
    // last chunk is shaped in the sync that groups: all eight are shaped.
    expect(new Set(wedges.map((wedge) => wedge.adjustments![0])).size).toBe(8);
  });
});

describe("Update all, with the old group still on the slide", () => {
  it("starts the 92-shape chart's redraw at the chunk 93 shapes leave room for, in S1's 28 syncs", async () => {
    const b = await boot(S1_WEB);
    await b.insert();
    b.lines.length = 0;
    const values = WIDE.series.map((one) => ({
      ...one,
      values: one.values.map((value) => (value === null ? null : value + 1)),
    }));
    await b.update({ ...WIDE, series: values });

    const sizes = chunks(b.lines);
    // The old group (92 shapes and itself) is on the slide until the sync that
    // groups, and every new add pays for all 93 of them: S1's 292 s, 28 syncs.
    expect(sizes[0]).toBe(chunkFor(93, "web"));
    expect(sizes[0]).toBe(5);
    expect(sizes).toHaveLength(28);
    expect(sizes.reduce((sum, one) => sum + one, 0)).toBe(WIDE_SHAPES);
    expect(Math.max(...sizes)).toBeLessThan(12);
    // One group stays, with the new shapes.
    expect(b.slideShapes()).toHaveLength(1);
    expect(b.slideShapes()[0]!.group!.shapes).toHaveLength(WIDE_SHAPES);
  });

  it("draws the new chart in chunks with the same line, under the link's label", async () => {
    const b = await boot(S1_WEB);
    await b.insert(columnChart(6));
    const oldSlideShapes = 20 + 1;
    b.lines.length = 0;
    await b.update(WIDE);

    // Each chunk is what the model allows for the slide as it stands, the old
    // 20-shape group and its own shape counted.
    const expected: number[] = [];
    for (let done = 0; done < WIDE_SHAPES;) {
      const size = Math.min(
        chunkFor(oldSlideShapes + done, "web"),
        WIDE_SHAPES - done,
      );
      expected.push(size);
      done += size;
    }
    expect(chunks(b.lines)).toEqual(expected);
    expect(b.lines[0]).toBe(line(0));
    expect(b.lines.at(-1)).toBe(line(WIDE_SHAPES));
    expect(b.slideShapes()).toHaveLength(1);
    expect(b.slideShapes()[0]!.group!.shapes).toHaveLength(WIDE_SHAPES);
  });
});
