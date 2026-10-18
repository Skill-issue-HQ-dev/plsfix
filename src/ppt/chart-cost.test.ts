// What drawing a chart costs on PowerPoint for the web, as S1 measured it
// (docs/superpowers/specs/2026-10-03-s1-spike.md, 1.2 and 7): a sync of n adds
// onto a slide of p shapes takes c + a n + b (n p + n (n - 1) / 2) seconds, the
// chunk is the biggest one that stays under the sync target, and the question's
// line is two minutes. Pure arithmetic, so these tests pin the S1 tables and
// the product intent, and move with the cost table only where S1 says so.

import { describe, expect, it } from "vitest";
import { layoutChart } from "../chart-shapes";
import { DEMO_COMBO, DEMO_DOUGHNUT } from "../../test/chart-demo.support";
import {
  askMinutes,
  ASK_ABOVE_SECONDS,
  canDraw,
  chunkFor,
  estimateSeconds,
  HOST_COSTS,
  SAFE_SYNC_SECONDS,
  SHAPES_PER_SYNC,
  SYNC_TARGET_SECONDS,
  SYNC_TIMEOUT_MS,
  syncSeconds,
} from "./chart-cost";

// Shape counts of what a banker sends (audit A.2): a 12 x 3 column chart with
// labels, a 10-bar waterfall, an 8-slice pie, a 12 x 3 line, and the two
// monthly series the old desktop budget turned into pictures.
const COLUMN_12X3 = 92;
const WATERFALL_10 = 41;
const PIE_8 = 33;
const PIE_21 = 21;
const LINE_12X3 = 125;
const LINE_40X3 = 375;
const COLUMN_40X6 = 264;
// A slide with a title, a text box and a picture on it.
const BUSY_SLIDE = 10;

describe("the web's numbers are S1's", () => {
  it("c 0.6 s a sync, a 0.17 s an add, b 0.0205 s per shape already there, 2 s to group", () => {
    expect(HOST_COSTS.web).toEqual({
      perSync: 0.6,
      perAdd: 0.17,
      perExisting: 0.0205,
      grouping: 2,
    });
  });

  it("a chunk is at most 12 adds and a sync is held to 12 s, 20 s is the line nothing may need", () => {
    expect(SHAPES_PER_SYNC).toBe(12);
    expect(SYNC_TARGET_SECONDS).toBe(12);
    expect(SAFE_SYNC_SECONDS).toBe(20);
  });

  it("the web gives up on a sync at 35 s, the desktop (not measured) at 60 s", () => {
    expect(SYNC_TIMEOUT_MS).toEqual({ web: 35_000, desktop: 60_000 });
  });

  it("the desktop row is cheaper than the web in every field", () => {
    for (const field of Object.keys(
      HOST_COSTS.web,
    ) as (keyof typeof HOST_COSTS.web)[]) {
      expect(HOST_COSTS.desktop[field]).toBeLessThan(HOST_COSTS.web[field]);
    }
  });
});

describe("syncSeconds", () => {
  it("is c + a n + b (n p + n (n - 1) / 2): each add sees the ones queued before it", () => {
    const { perSync, perAdd, perExisting } = HOST_COSTS.web;
    const n = 9;
    const p = 24;
    expect(syncSeconds(n, p, "web")).toBeCloseTo(
      perSync + perAdd * n + perExisting * (n * p + (n * (n - 1)) / 2),
      9,
    );
  });

  it("reproduces S1's measured batches: 12 rectangles at 0 about 4 s, one add at 192 about 4.7 s", () => {
    expect(syncSeconds(12, 0, "web")).toBeCloseTo(4.0, 1);
    expect(syncSeconds(1, 192, "web")).toBeCloseTo(4.7, 1);
  });
});

describe("chunkFor", () => {
  it("is S1's chunk table: 12 up to 20 shapes, 8 at 50, 5 at 100, 3 at 150, 2 at 200, 1 from 300", () => {
    const table: [number, number][] = [
      [0, 12],
      [20, 12],
      [50, 8],
      [100, 5],
      [150, 3],
      [200, 2],
      [300, 1],
      [500, 1],
    ];
    for (const [population, chunk] of table) {
      expect(chunkFor(population, "web"), `p = ${String(population)}`).toBe(
        chunk,
      );
    }
  });

  it("is the largest chunk of at most 12 whose sync stays inside the target, never less than 1", () => {
    for (let p = 0; p <= 1_200; p += 1) {
      const n = chunkFor(p, "web");
      expect(n).toBeGreaterThanOrEqual(1);
      expect(n).toBeLessThanOrEqual(SHAPES_PER_SYNC);
      if (n > 1) {
        expect(syncSeconds(n, p, "web")).toBeLessThanOrEqual(
          SYNC_TARGET_SECONDS,
        );
      }
      if (n < SHAPES_PER_SYNC) {
        expect(syncSeconds(n + 1, p, "web")).toBeGreaterThan(
          SYNC_TARGET_SECONDS,
        );
      }
    }
  });

  it("only ever shrinks as the slide fills", () => {
    let last = SHAPES_PER_SYNC;
    for (let p = 0; p <= 1_200; p += 1) {
      const n = chunkFor(p, "web");
      expect(n).toBeLessThanOrEqual(last);
      last = n;
    }
  });

  it("stays at 12 on the desktop until the slide holds thousands of shapes", () => {
    expect(chunkFor(0, "desktop")).toBe(12);
    expect(chunkFor(4_000, "desktop")).toBe(12);
    expect(chunkFor(100_000, "desktop")).toBeLessThan(12);
  });
});

describe("estimateSeconds", () => {
  it("is the syncs chunkFor plans from the slide as it stands, plus the grouping", () => {
    // 92 shapes onto an empty web slide: 12, 12, 12, 11, 9, 8, 7, 6, 6, 5, 4.
    let seconds = 0;
    let done = 0;
    for (const n of [12, 12, 12, 11, 9, 8, 7, 6, 6, 5, 4]) {
      seconds += syncSeconds(n, done, "web");
      done += n;
    }
    expect(done).toBe(92);
    expect(estimateSeconds(92, 0, "web")).toBeCloseTo(
      seconds + HOST_COSTS.web.grouping,
      9,
    );
  });

  it("is S1's chart table before grouping: 21 / 33 / 41 / 92 / 150 shapes = 9 / 18 / 26 / 108 / 271 s", () => {
    const table: [number, number][] = [
      [PIE_21, 9],
      [PIE_8, 18],
      [WATERFALL_10, 26],
      [COLUMN_12X3, 108],
      [150, 271],
    ];
    for (const [shapes, seconds] of table) {
      const before =
        estimateSeconds(shapes, 0, "web") - HOST_COSTS.web.grouping;
      expect(Math.round(before), `${String(shapes)} shapes`).toBe(seconds);
    }
  });

  it("is S1's table for a slide that already holds 20 shapes: 32 s for the pie, 43 for the waterfall, 148 for the column chart", () => {
    const grouping = HOST_COSTS.web.grouping;
    expect(Math.round(estimateSeconds(PIE_8, 20, "web") - grouping)).toBe(32);
    expect(Math.round(estimateSeconds(41, 20, "web") - grouping)).toBe(43);
    expect(Math.round(estimateSeconds(92, 20, "web") - grouping)).toBe(148);
  });

  it("is S1's refresh table when the old group is still on the slide: 92 shapes onto 93 is about 292 s", () => {
    const grouping = HOST_COSTS.web.grouping;
    const refresh = estimateSeconds(COLUMN_12X3, COLUMN_12X3 + 1, "web");
    expect(Math.round(refresh - grouping)).toBeGreaterThanOrEqual(292);
    expect(Math.round(refresh - grouping)).toBeLessThanOrEqual(294);
  });

  it("is nothing for a chart with no shapes", () => {
    expect(estimateSeconds(0, 0, "web")).toBe(0);
    expect(estimateSeconds(0, 40, "desktop")).toBe(0);
  });

  it("treats a negative count as nothing rather than a refund", () => {
    expect(estimateSeconds(-5, 3, "web")).toBe(0);
    expect(estimateSeconds(10, -3, "web")).toBe(estimateSeconds(10, 0, "web"));
  });

  it("grows with the chart, with a fuller slide, and is slower on the web", () => {
    expect(estimateSeconds(50, 0, "web")).toBeGreaterThan(
      estimateSeconds(40, 0, "web"),
    );
    expect(estimateSeconds(50, 30, "web")).toBeGreaterThan(
      estimateSeconds(50, 0, "web"),
    );
    expect(estimateSeconds(50, 0, "web")).toBeGreaterThan(
      estimateSeconds(50, 0, "desktop"),
    );
  });

  it("grows faster than the shape count: each add waits behind the ones before it", () => {
    const double = estimateSeconds(200, 0, "web");
    const single = estimateSeconds(100, 0, "web");
    expect(double).toBeGreaterThan(2 * single);
  });
});

describe("the question's line", () => {
  it("is two minutes", () => {
    expect(ASK_ABOVE_SECONDS).toBe(120);
  });

  it("lets the pie, the waterfall and the column chart draw on an empty web slide without a question", () => {
    for (const shapes of [COLUMN_12X3, WATERFALL_10, PIE_8]) {
      expect(estimateSeconds(shapes, 0, "web")).toBeLessThanOrEqual(
        ASK_ABOVE_SECONDS,
      );
    }
  });

  it("asks before the column chart once the slide holds ten other shapes: it is about 130 s there", () => {
    expect(estimateSeconds(COLUMN_12X3, BUSY_SLIDE, "web")).toBeGreaterThan(
      ASK_ABOVE_SECONDS,
    );
    expect(estimateSeconds(COLUMN_12X3, BUSY_SLIDE, "web")).toBeLessThan(135);
  });

  it("still lets the pie and the waterfall draw on a busy web slide", () => {
    for (const shapes of [WATERFALL_10, PIE_8]) {
      expect(estimateSeconds(shapes, BUSY_SLIDE, "web")).toBeLessThanOrEqual(
        ASK_ABOVE_SECONDS,
      );
    }
  });

  it("asks before a chart that would take the web minutes", () => {
    for (const shapes of [LINE_12X3, COLUMN_40X6, LINE_40X3]) {
      expect(estimateSeconds(shapes, BUSY_SLIDE, "web")).toBeGreaterThan(
        ASK_ABOVE_SECONDS,
      );
    }
  });

  it("never asks about the biggest chart the caps allow today on the desktop", () => {
    for (const shapes of [COLUMN_40X6, LINE_40X3]) {
      expect(estimateSeconds(shapes, BUSY_SLIDE, "desktop")).toBeLessThan(
        ASK_ABOVE_SECONDS,
      );
    }
  });

  it("asks on the desktop too once a chart is big enough (the C2 caps)", () => {
    expect(estimateSeconds(1500, BUSY_SLIDE, "desktop")).toBeGreaterThan(
      ASK_ABOVE_SECONDS,
    );
  });
});

describe("canDraw", () => {
  it("is false once even one add is predicted past the safe 20 s: about 940 shapes on a web slide", () => {
    const { perSync, perAdd, perExisting } = HOST_COSTS.web;
    const limit = Math.floor(
      (SAFE_SYNC_SECONDS - perSync - perAdd) / perExisting,
    );
    expect(limit).toBe(938);
    expect(canDraw(limit, "web")).toBe(true);
    expect(canDraw(limit + 1, "web")).toBe(false);
  });

  it("is true for any slide a person has built, on both hosts", () => {
    expect(canDraw(0, "web")).toBe(true);
    expect(canDraw(400, "web")).toBe(true);
    expect(canDraw(4_000, "desktop")).toBe(true);
  });

  it("keeps the line under the deadline: a draw it allows is one the host's timeout also allows", () => {
    expect(SAFE_SYNC_SECONDS * 1000).toBeLessThan(SYNC_TIMEOUT_MS.web);
  });
});

// C5: the doughnut and the combo are priced by their primitives, one add each,
// like every other kind (the cost model reads the layout's length, never the
// kind). The demo's two charts at the slide's chart size.
describe("the doughnut and the combo are priced by their shapes", () => {
  const BOX = { left: 0, top: 0, width: 600, height: 300 };

  it("counts every primitive of the demo's doughnut and combo as one add", () => {
    const doughnut = layoutChart(DEMO_DOUGHNUT, BOX).length;
    const combo = layoutChart(DEMO_COMBO, BOX).length;
    // Doughnut: title, 6 arcs, 6 labels, 6 legend swatches and 6 names.
    expect(doughnut).toBe(25);
    // Combo: title, 12 bars, 12 + 12 labels, the zero line, 11 connectors,
    // 12 markers, 12 months and a 2-entry legend.
    expect(combo).toBe(1 + 12 + 24 + 1 + 11 + 12 + 12 + 4);
    expect(estimateSeconds(combo, 0, "web")).toBeGreaterThan(
      estimateSeconds(doughnut, 0, "web"),
    );
  });

  it("lets the demo's doughnut and combo draw on a busy web slide without a question: about 40 s and 96 s", () => {
    const doughnut = layoutChart(DEMO_DOUGHNUT, BOX).length;
    const combo = layoutChart(DEMO_COMBO, BOX).length;
    expect(estimateSeconds(doughnut, BUSY_SLIDE, "web")).toBeLessThan(60);
    const long = estimateSeconds(combo, BUSY_SLIDE, "web");
    expect(long).toBeGreaterThan(90);
    expect(long).toBeLessThan(ASK_ABOVE_SECONDS);
  });
});

describe("askMinutes", () => {
  it("rounds to the nearest minute and never says less than the line it asks above", () => {
    const floor = Math.round(ASK_ABOVE_SECONDS / 60);
    expect(askMinutes(121)).toBe(floor);
    expect(askMinutes(149)).toBe(2);
    expect(askMinutes(150)).toBe(3);
    expect(askMinutes(400)).toBe(7);
    expect(askMinutes(10)).toBe(floor);
  });
});
