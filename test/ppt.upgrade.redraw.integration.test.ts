// Redraw from this deck, with a picture-era link ticked: a chart that landed as
// a picture becomes shapes from the deck's own copy, and says so; a picture the
// user chose stays a picture and the summary says that. A range picture whose
// record holds a table becomes a table. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { buildDeckRecord, encodeDeckRecord } from "../src/link/deck-record";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { removeAllParts, seedParts } from "./ppt.deck-store.support";
import { seedChart, seedLink } from "./ppt.support";
import {
  bootUpgrade,
  boxOf,
  chosenPicture,
  clearDrawUi,
  COLUMN,
  LABEL,
  moveTo,
  PNG,
  shapesOf,
  tagOf,
  type UpgradePane,
} from "./ppt.upgrade.support";

enableStrictLoadSemantics();

let pane: UpgradePane;
beforeEach(async () => {
  pane = await bootUpgrade();
});
afterEach(async () => {
  await clearDrawUi();
  uninstallFakePpt();
});

const CORNER = { left: 80, top: 320, width: 380, height: 190 };
const rowsOf = () => pane.links.listLinks(pane.relay);

describe("Redraw with a picture-era chart ticked", () => {
  it("draws it as shapes at the picture's corner and width, from the record", async () => {
    // A PowerPoint with no shape API inserts the chart as a picture and the
    // record keeps its data; the user then moves to one that can draw it.
    pane.helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.8"),
    );
    const item = await seedChart(COLUMN, PNG);
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const picture = shapesOf(pane)[0]!;
    moveTo(picture, CORNER);
    pane.helpers.setSupported(() => true);

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 1, chosen: 0, failed: 0 });
    expect(summary.notes).toEqual([`${LABEL} now a native chart`]);
    expect(shapesOf(pane)).toHaveLength(1);
    const drawn = shapesOf(pane)[0]!;
    expect(drawn.type).toBe("Group");
    expect(boxOf(drawn)).toEqual(CORNER);
    expect(tagOf(drawn)).toMatchObject({ id: item.id, kind: "chart" });
  });

  it("leaves a picture the user chose alone and says so", async () => {
    await chosenPicture(pane);
    const picture = shapesOf(pane)[0]!;

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 0, chosen: 1, failed: 0 });
    expect(summary.notes).toEqual([
      `${LABEL}: you chose this picture, so it stays one.`,
    ]);
    expect(pane.actions.summarizeRedraw(summary)).toBe(
      "1 kept as the picture you chose",
    );
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(picture.setImageCalls).toBe(1);
  });
});

describe("Redraw with a range picture ticked", () => {
  it("becomes a table when the deck's copy of the link holds one", async () => {
    const item = await seedLink(PNG);
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const picture = shapesOf(pane)[0]!;
    moveTo(picture, { left: 80, top: 320, width: 200, height: 40 });
    const held = pane.boot.deckRecords().get(item.id)!;
    // The deck's copy after Convert to table pushed a table the update never
    // painted: the record is the table's, the object is still the picture.
    const table = {
      v: 1 as const,
      kind: "table" as const,
      rows: 1,
      cols: 1,
      cells: [[{ t: "Revenue" }]],
      widths: [90],
      src: held.src,
      pushedAt: new Date(Date.parse(held.pushedAt) + 1000).toISOString(),
      hash: "1".repeat(64),
    };
    const record = buildDeckRecord(
      { ...held, kind: "table" },
      table,
      held.rev + 1,
      new Date().toISOString(),
    );
    await removeAllParts();
    await seedParts([encodeDeckRecord(record)]);
    await pane.boot.loadDeckStore();

    const summary = await pane.actions.redrawLinks(await rowsOf());

    expect(summary).toMatchObject({ redrawn: 1, failed: 0 });
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.type).toBe("Table");
    expect(shapesOf(pane)[0]!.left).toBe(80);
    expect(tagOf(shapesOf(pane)[0]!)).toMatchObject({ kind: "table" });
  });
});
