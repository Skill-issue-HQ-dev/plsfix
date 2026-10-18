// Redraw from this deck: an object that is still on its slide drawn again from
// the deck's own record, with no Excel at hand - a chart group someone cut
// bars out of, a chart that landed as a picture, a text or a table written
// over. The new object is built before the old one goes, where the old one
// sat; the record is written in the same batch. Strict load semantics on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { encodeDeckRecord } from "../src/link/deck-record";
import { TAG_KEY } from "../src/link/model";
import { settleHungSync } from "./hung-sync";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
} from "./fakeppt";
import {
  bootActions,
  insertedThenDeleted,
  keepIdentityOnly,
  tagOf,
  type ActionsPane,
} from "./ppt.deck-actions.support";
import {
  CELLS,
  PNG,
  WIDTHS,
  partsIn,
  recordsIn,
  removeAllParts,
  seedParts,
} from "./ppt.deck-store.support";
import {
  columnChart,
  scatterChart,
  pushText,
  seedChart,
  seedLink,
  seedTable,
  seedText,
} from "./ppt.support";

import type * as DrawModule from "../src/ppt/draw-progress";

enableStrictLoadSemantics();

let pane: ActionsPane;
// The draw UI of the module graph this pane booted: a static import would be
// the one from before bootPpt reset the modules.
let drawUi: typeof DrawModule;
beforeEach(async () => {
  pane = await bootActions();
  drawUi = await import("../src/ppt/draw-progress");
});
afterEach(() => {
  drawUi.installDrawUi(null);
  uninstallFakePpt();
});

const committed = (): ReturnType<typeof recordsIn> =>
  recordsIn(pane.presentation, true);
const onSlide = (index: number): FakePptShape[] =>
  pane.presentation.slides[index]!.shapes;
const rowsOf = (): ReturnType<typeof pane.links.listLinks> =>
  pane.links.listLinks(pane.relay);

// How many round trips a Put back of a six-bar chart takes on a fresh deck,
// so a test can hang the last one (the sync that groups and tags the chart).
async function syncsOfPutBack(): Promise<number> {
  const probe = await bootActions();
  const item = await seedChart(columnChart(6), PNG);
  await insertedThenDeleted(probe, item);
  const from = probe.helpers.syncCount();
  await probe.actions.putBack(item.id);
  return probe.helpers.syncCount() - from;
}

describe("a chart", () => {
  it("whose group lost bars is drawn again where it sits, whole", async () => {
    const item = await seedChart(columnChart(6), PNG);
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const [group] = onSlide(0);
    const whole = group!.group!.shapes.length;
    const box = { left: group!.left, top: group!.top, width: group!.width };
    pane.presentation.deleteShape(group!.group!.shapes[2]!.id);
    pane.presentation.deleteShape(group!.group!.shapes[3]!.id);
    expect(group!.group!.shapes).toHaveLength(whole - 2);
    const before = committed()[0]!;

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 1, failed: 0 });
    expect(onSlide(0)).toHaveLength(1);
    const [again] = onSlide(0);
    expect(again!.id).not.toBe(group!.id);
    expect(again!.type).toBe("Group");
    expect(again!.group!.shapes).toHaveLength(whole);
    expect(again!.left).toBeCloseTo(box.left, 1);
    expect(again!.top).toBeCloseTo(box.top, 1);
    expect(again!.width).toBeCloseTo(box.width, 1);
    expect(tagOf(again!)).toMatchObject({ id: item.id, rev: before.rev });
    expect(again!.tags.get(TAG_KEY)).toBe(item.token);
    // One part, written again in the batch that tagged the new group.
    expect(partsIn(pane.presentation)).toHaveLength(1);
    expect(committed()[0]!.savedAt >= before.savedAt).toBe(true);
    expect(committed()[0]).toMatchObject({ id: item.id, dataKept: true });
  });

  it("that landed as a picture becomes shapes, with no Excel at hand", async () => {
    // A PowerPoint with no shape API: the chart is inserted as a picture, and
    // the record still keeps the chart's data.
    pane.helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.8"),
    );
    const item = await seedChart(columnChart(6), PNG);
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    expect(onSlide(0)[0]!.type).not.toBe("Group");
    expect(committed()[0]).toMatchObject({ dataKept: true });
    // The user moved to a PowerPoint that can draw it.
    pane.helpers.setSupported(() => true);

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 1, failed: 0 });
    expect(onSlide(0)).toHaveLength(1);
    expect(onSlide(0)[0]!.type).toBe("Group");
    expect(tagOf(onSlide(0)[0]!).id).toBe(item.id);
  });

  it("inside a group the user built is reported, and the rest still goes on", async () => {
    const chart = await seedChart(columnChart(6), PNG);
    const text = await seedText("EUR 3.2m");
    await pane.links.insertFromInbox(chart, pane.ws, pane.relay);
    await pane.links.insertFromInbox(text, pane.ws, pane.relay);
    const group = onSlide(0).find((shape) => shape.type === "Group")!;
    const other = onSlide(0).find((shape) => shape.type !== "Group")!;
    pane.presentation.groupShapes(
      [group.id, other.id],
      pane.presentation.slides[0]!.id,
    );

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary.failed + summary.redrawn).toBe(2);
    expect(summary.failed).toBe(1);
    expect(summary.failures.join("\n")).toMatch(/ungroup the chart/);
  });

  it("put back while the host swallows the draw inserts nothing, says so in the list's words, and keeps the row", async () => {
    const syncs = await syncsOfPutBack();
    pane = await bootActions();
    const item = await seedChart(columnChart(6), PNG);
    await insertedThenDeleted(pane, item);
    pane.helpers.hangNextSync(syncs - 1);
    vi.useFakeTimers();

    const stopped = settleHungSync(pane.actions.putBack(item.id));
    await expect(stopped).rejects.toThrow(
      "PowerPoint stopped answering; nothing was inserted, the row is still listed",
    );
    vi.useRealTimers();

    expect(onSlide(0)).toEqual([]);
    expect(committed()).toHaveLength(1);
    // And it can be asked again once the host answers.
    await pane.actions.putBack(item.id);
    expect(onSlide(0)[0]!.type).toBe("Group");
  });

  it("asks before a very long draw, offers no picture, and cancelling leaves the row listed", async () => {
    pane.helpers.setPlatform("OfficeOnline");
    const big = {
      ...columnChart(40),
      series: Array.from({ length: 6 }, (_n, i) => ({
        ...columnChart(40).series[0]!,
        name: `S${String(i)}`,
      })),
    };
    const item = await seedChart(big, PNG);
    await insertedThenDeleted(pane, item);
    const asked: string[][] = [];
    drawUi.installDrawUi({
      progress: () => undefined,
      ask: async (_question, choices) => {
        asked.push(choices.map((choice) => choice.label));
        return "cancel";
      },
    });

    await expect(pane.actions.putBack(item.id)).rejects.toThrow(
      "Cancelled: nothing was inserted, the row is still listed",
    );

    // A draw from the deck has no picture to fall back on.
    expect(asked).toEqual([["Draw it", "Cancel"]]);
    expect(onSlide(0)).toEqual([]);
    expect(committed()).toHaveLength(1);
  });
});

describe("a text or a table", () => {
  it("has its text written back from the record", async () => {
    const item = await seedText("EUR 3.2m");
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    onSlide(0)[0]!.text = "scribbled over";

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 1, failed: 0 });
    expect(onSlide(0)[0]!.text).toBe("EUR 3.2m");
  });

  it("has its cells repainted from the record", async () => {
    const item = await seedTable(CELLS, WIDTHS);
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    onSlide(0)[0]!.table!.cells[0]![0]!.text = "scribbled over";

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 1, failed: 0 });
    expect(onSlide(0)[0]!.table!.cells[0]![0]!.text).toBe("Revenue");
  });
});

describe("what the deck cannot redraw", () => {
  it("counts a link whose record keeps no data, and touches nothing", async () => {
    const picture = await seedLink(PNG);
    await pane.links.insertFromInbox(picture, pane.ws, pane.relay);
    const before = onSlide(0)[0]!;

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 0, noCopy: 1, failed: 0 });
    expect(onSlide(0)[0]!.id).toBe(before.id);
    expect(pane.actions.redrawDetails(summary)).toMatch(
      /kept no data.*Send it again from Excel/,
    );
  });

  it("fails a chart kind this host draws as a picture only, in the kind's own sentence, and leaves the object", async () => {
    const item = await seedChart(scatterChart(4), PNG);
    // It lands as a picture on every host for now, its data kept in the record.
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const before = onSlide(0)[0]!;

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 0, noCopy: 0, failed: 1 });
    expect(summary.failures).toEqual([
      `${item.label}: scatter charts are not drawn as shapes yet. The object stays as it is.`,
    ]);
    expect(summary.failures[0]).not.toMatch(/this PowerPoint/);
    expect(onSlide(0)[0]!.id).toBe(before.id);
  });

  it("counts a link a record was cut down for", async () => {
    const item = await seedText("EUR 3.2m");
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    await keepIdentityOnly(pane, item);
    const summary = await pane.actions.redrawLinks(await rowsOf());
    expect(summary).toMatchObject({ redrawn: 0, noCopy: 1 });
  });

  it("counts a link the deck holds no record of, as a deck from before the store is", async () => {
    const item = await seedText("EUR 3.2m");
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    await removeAllParts();
    await pane.boot.loadDeckStore();
    const summary = await pane.actions.redrawLinks(await rowsOf());
    expect(summary).toMatchObject({ redrawn: 0, noCopy: 1 });
  });

  it("will not draw an older copy over a newer object", async () => {
    const item = await seedText("EUR 3.2m");
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    await pushText(item, "EUR 3.9m");
    await pane.links.updateLinks(await rowsOf(), pane.relay);
    expect(onSlide(0)[0]!.text).toBe("EUR 3.9m");
    // The deck's record fell behind: a store that was full, a co-author's pane.
    const current = pane.boot.deckRecords().get(item.id)!;
    await removeAllParts();
    await seedParts([
      encodeDeckRecord({
        ...current,
        pushedAt: "2020-01-01T00:00:00.000Z",
        payload: { v: 1, kind: "text", text: "EUR 1.0m", hash: "0".repeat(64) },
      }),
    ]);
    await pane.boot.loadDeckStore();

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 0, newer: 1, failed: 0 });
    expect(onSlide(0)[0]!.text).toBe("EUR 3.9m");
    expect(pane.actions.redrawDetails(summary)).toMatch(
      /newer than.*Update it/,
    );
  });
});

describe("the sentence", () => {
  const empty = {
    redrawn: 0,
    noCopy: 0,
    newer: 0,
    chosen: 0,
    failed: 0,
    failures: [],
    notes: [],
  };

  it("says what was done in the toast's words", () => {
    const said = pane.actions.summarizeRedraw;
    expect(said({ ...empty, redrawn: 3 })).toBe("3 redrawn");
    expect(said({ ...empty, redrawn: 1, noCopy: 2, failed: 1 })).toBe(
      "1 redrawn, 2 with no data in the deck, 1 failed",
    );
    expect(said({ ...empty, newer: 1 })).toBe("1 newer than the copy");
    expect(said(empty)).toBe("Nothing to redraw");
  });

  it("keeps the details to what needs saying", () => {
    expect(pane.actions.redrawDetails(empty)).toBeUndefined();
    expect(
      pane.actions.redrawDetails({
        ...empty,
        failed: 1,
        failures: ["Model!B4:F12: boom"],
        notes: ["Model!B4:F12 The deck's copy of this link was not saved"],
      }),
    ).toBe(
      "Model!B4:F12: boom\nModel!B4:F12 The deck's copy of this link was not saved",
    );
  });
});
