// What a deck's own copy of a link can do for an object that left its slide:
// Put back draws it again with its id and token (so Excel's next push refreshes
// it as before), writing the record through the deck flow in the batch that tags
// the object. Strict load semantics on. Drawing an object again in place is
// ppt.deck-redraw.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { SMOOTH_STRAIGHT } from "../src/link/chart-model";
import { TAG_KEY, TAG_LINK } from "../src/link/model";
import { type InsertTarget } from "../src/ppt/placement";
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
  KINDS,
  PNG,
  WIDTHS,
  partsIn,
  recordsIn,
  removeAllParts,
  withoutCustomXml,
} from "./ppt.deck-store.support";
import {
  columnChart,
  scatterChart,
  pushText,
  seedChart,
  seedTable,
  seedText,
  src,
} from "./ppt.support";

enableStrictLoadSemantics();

let pane: ActionsPane;
beforeEach(async () => {
  pane = await bootActions();
});
afterEach(() => {
  uninstallFakePpt();
});

const committed = (): ReturnType<typeof recordsIn> =>
  recordsIn(pane.presentation, true);
const onSlide = (index: number): FakePptShape[] =>
  pane.presentation.slides[index]!.shapes;

describe("Put back", () => {
  it.each(KINDS.filter((kind) => kind.kept))(
    "draws $name again with its own id, token and revision",
    async ({ seed, kind }) => {
      const item = await seed();
      await insertedThenDeleted(pane, item);
      expect(onSlide(0)).toEqual([]);
      const before = committed()[0]!;

      const said = await pane.actions.putBack(item.id);

      expect(said).toMatch(new RegExp(`^Put ${item.label} back\\.`));
      const [back] = onSlide(0);
      expect(back).toBeDefined();
      expect(tagOf(back!)).toMatchObject({
        id: item.id,
        kind,
        rev: before.rev,
        pushedAt: before.pushedAt,
        src,
      });
      expect(back!.tags.get(TAG_KEY)).toBe(item.token);
      // The record rode the batch that tagged the object: one part, the same.
      expect(partsIn(pane.presentation)).toHaveLength(1);
      expect(committed()[0]).toMatchObject({
        id: item.id,
        token: item.token,
        rev: before.rev,
        label: before.label,
        sentAt: before.sentAt,
        dataKept: true,
      });
    },
  );

  // What a chart could not keep is said in the same sentence as the label, not
  // after a full stop ("Put Model: B4:F12 back. smooth lines are drawn straight").
  // A host that cannot draw the chart does not put it back at all (below).
  it("reads as one sentence when the chart comes back with something it could not keep", async () => {
    const base = columnChart(4);
    const item = await seedChart(
      {
        ...base,
        kind: "line",
        series: [{ ...base.series[0]!, smooth: true }],
      },
      PNG,
    );
    await insertedThenDeleted(pane, item);

    const said = await pane.actions.putBack(item.id);

    expect(said).toBe(`Put ${item.label} back (${SMOOTH_STRAIGHT}).`);
    expect(onSlide(0)[0]!.type).toBe("Group");
  });

  it("refreshes exactly as before: Excel's next push reaches the object put back", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    await pane.actions.putBack(item.id);
    await pushText(item, "EUR 3.9m");

    const rows = await pane.links.listLinks(pane.relay);
    expect(rows.map((row) => row.status)).toEqual(["updateAvailable"]);
    const summary = await pane.links.updateLinks(rows, pane.relay);
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(onSlide(0)[0]!.text).toBe("EUR 3.9m");
    expect(committed()[0]).toMatchObject({
      rev: 2,
      payload: { kind: "text", text: "EUR 3.9m" },
    });
  });

  it("puts a table back with its cells, widths and header", async () => {
    const item = await seedTable(CELLS, WIDTHS, true);
    await insertedThenDeleted(pane, item);
    await pane.actions.putBack(item.id);
    const table = onSlide(0)[0]!.table!;
    expect([table.rowCount, table.columnCount]).toEqual([2, 2]);
    expect(table.cells[1]![1]!.text).toBe("800");
    expect(table.styleSettings.isFirstRowHighlighted).toBe(true);
  });

  it("lands on the slide and in the spot the pickers name", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    const target: InsertTarget = {
      slideId: pane.presentation.slides[2]!.id,
      where: "top-right",
    };

    await pane.actions.putBack(item.id, target);

    expect(onSlide(0)).toEqual([]);
    const [back] = onSlide(2);
    expect(back!.left + back!.width / 2).toBeGreaterThan(480);
    expect(back!.top + back!.height / 2).toBeLessThan(270);
    // The slide the user asked for is on screen afterwards, like an Inbox insert.
    expect(pane.presentation.selectedSlideIds).toEqual([
      pane.presentation.slides[2]!.id,
    ]);
  });

  it("fits into the selected shape, and does not show the slide first: that would drop the selection", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    const slide = pane.presentation.slides[0]!;
    const frame = pane.presentation.addShape(slide, {
      type: "GeometricShape",
      left: 100,
      top: 100,
      width: 400,
      height: 200,
    });
    pane.helpers.selectShapes([frame.id]);
    const host = await import("../src/ppt/host");
    const shown = vi.spyOn(host, "goToSlide");

    await pane.actions.putBack(item.id, {
      slideId: slide.id,
      where: "selected-shape",
    });

    expect(shown).not.toHaveBeenCalled();
    const back = onSlide(0).find((shape) => shape.id !== frame.id)!;
    expect(back.left).toBeGreaterThanOrEqual(frame.left - 0.01);
    expect(back.left + back.width).toBeLessThanOrEqual(
      frame.left + frame.width + 0.01,
    );
    expect(back.top).toBeGreaterThanOrEqual(frame.top - 0.01);
  });

  it("shows a named slide before it reads it, so the shapes already there are in the way", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    const far = pane.presentation.slides[2]!;
    const wall = pane.presentation.addShape(far, {
      type: "GeometricShape",
      left: 36,
      top: 36,
      width: 700,
      height: 468,
    });
    // A slide PowerPoint for the web has not loaded reads as empty until it
    // is shown: with nothing in the way, the text would land on the wall.
    pane.helpers.holdSlideReads(far.id);

    await pane.actions.putBack(item.id, { slideId: far.id, where: "free" });

    const back = onSlide(2).find((shape) => shape.id !== wall.id)!;
    const overlap =
      back.left < wall.left + wall.width &&
      back.left + back.width > wall.left &&
      back.top < wall.top + wall.height &&
      back.top + back.height > wall.top;
    expect(overlap, "the object was placed over what the slide holds").toBe(
      false,
    );
  });

  it("costs the round trips of an insert, the deck's records read first, and no more: the record rides the tagging batch", async () => {
    // The part ids, then the text of the parts: Put back reads the file first.
    const STORE_READ_SYNCS = 2;
    const spent = async (keeps: boolean): Promise<number> => {
      pane = await bootActions();
      if (!keeps) withoutCustomXml(pane.helpers);
      const item = await seedText("EUR 3.2m");
      await insertedThenDeleted(pane, item);
      const from = pane.helpers.syncCount();
      // Without custom XML there is no record to put back from: the plain
      // insert of the same export is the bar.
      if (keeps) await pane.actions.putBack(item.id);
      else {
        // Put back reads the deck first, to be sure nothing carries the link.
        await (await import("../src/ppt/host")).scanLinks();
        await pane.links.insertFromInbox(item, pane.ws, pane.relay);
      }
      return pane.helpers.syncCount() - from;
    };
    expect(await spent(true)).toBe((await spent(false)) + STORE_READ_SYNCS);
  });

  it("brings a chart back as native shapes, with no picture anywhere", async () => {
    const item = await seedChart(columnChart(6), PNG);
    await insertedThenDeleted(pane, item);
    await pane.actions.putBack(item.id);
    const [group] = onSlide(0);
    expect(group!.type).toBe("Group");
    expect(tagOf(group!).id).toBe(item.id);
    expect(committed()[0]).toMatchObject({ dataKept: true });
    expect(JSON.stringify(partsIn(pane.presentation))).not.toContain(PNG);
  });

  it("says so, and inserts nothing, when the deck kept the identity only", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    await keepIdentityOnly(pane, item);

    await expect(pane.actions.putBack(item.id)).rejects.toThrow(
      /kept no data.*Send it again from Excel/,
    );
    expect(onSlide(0)).toEqual([]);
  });

  it("will not draw an object a slide holds again: the list was out of date (an undo, a co-author)", async () => {
    const item = await seedText("EUR 3.2m");
    const placed = await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const tag = onSlide(0)
      .find((shape) => shape.id === placed.shapeId)!
      .tags.get(TAG_LINK)!;
    pane.presentation.deleteShape(placed.shapeId);
    // Ctrl+Z: the object is back, and the list the user clicked in is stale.
    const slide = pane.presentation.slides[1]!;
    const back = pane.presentation.addShape(slide, {
      type: "TextBox",
      text: "EUR 3.2m",
      left: 50,
      top: 50,
      width: 200,
      height: 40,
    });
    back.tags.set(TAG_LINK, tag);
    back.tags.set(TAG_KEY, item.token);

    await expect(pane.actions.putBack(item.id)).rejects.toThrow(
      `${item.label} is already on slide 2.`,
    );
    expect(onSlide(0)).toEqual([]);
    expect(onSlide(1)).toHaveLength(1);
  });

  it("reads the deck's records first: a co-author's Break since the pane read them is refused in words and writes nothing", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    // Their pane broke the link: the part is gone from the file, and this
    // pane's index still holds it.
    await removeAllParts();

    for (const press of [1, 2]) {
      await expect(
        pane.actions.putBack(item.id),
        `press ${String(press)}`,
      ).rejects.toThrow("This deck no longer holds a copy of that link.");
    }

    expect(onSlide(0)).toEqual([]);
    // The record and its link key stay gone: no press wrote them back.
    expect(partsIn(pane.presentation)).toEqual([]);
  });

  it("says so when the deck holds no record of the link", async () => {
    await expect(pane.actions.putBack("0".repeat(32))).rejects.toThrow(
      "This deck no longer holds a copy of that link.",
    );
  });

  it("refuses a chart this PowerPoint has no shape API for: the deck has no picture to put in its place", async () => {
    const item = await seedChart(columnChart(6), PNG);
    await insertedThenDeleted(pane, item);
    pane.helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.8"),
    );

    await expect(pane.actions.putBack(item.id)).rejects.toThrow(
      /no shape API.*Send it again from Excel/,
    );
    expect(onSlide(0)).toEqual([]);
    expect(committed()).toHaveLength(1);
  });

  it("refuses a scatter in the kind's own sentence, never 'this PowerPoint cannot', and draws no picture", async () => {
    const item = await seedChart(scatterChart(4), PNG);
    await insertedThenDeleted(pane, item);
    const before = committed();

    const refused = await pane.actions.putBack(item.id).catch((e: Error) => e);

    expect(refused).toBeInstanceOf(Error);
    const said = (refused as Error).message;
    expect(said).toBe(
      `${item.label}: scatter charts are not drawn as shapes yet. Send it again from Excel: the Inbox inserts it as a picture.`,
    );
    expect(said).not.toMatch(/this PowerPoint/);
    expect(onSlide(0)).toEqual([]);
    expect(committed()).toEqual(before);
  });
});

// A second object of a link is PowerPoint's own copy and paste: the tags travel
// with the shape, so nothing here needs a button of its own.
describe("a copy PowerPoint makes of a linked object", () => {
  it("is updated with the original by one Update all, and one record serves both", async () => {
    const item = await seedText("EUR 3.2m");
    const placed = await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const original = onSlide(0).find((shape) => shape.id === placed.shapeId)!;
    const copy = pane.presentation.addShape(pane.presentation.slides[1]!, {
      type: "TextBox",
      text: "EUR 3.2m",
      left: 50,
      top: 50,
      width: 200,
      height: 40,
    });
    copy.tags.set(TAG_LINK, original.tags.get(TAG_LINK)!);
    copy.tags.set(TAG_KEY, item.token);
    await pushText(item, "EUR 3.9m");

    const summary = await pane.links.updateLinks(
      await pane.links.listLinks(pane.relay),
      pane.relay,
    );

    expect(summary).toMatchObject({ updated: 2, failed: 0 });
    expect(onSlide(0)[0]!.text).toBe("EUR 3.9m");
    expect(onSlide(1)[0]!.text).toBe("EUR 3.9m");
    expect(partsIn(pane.presentation)).toHaveLength(1);
  });
});
