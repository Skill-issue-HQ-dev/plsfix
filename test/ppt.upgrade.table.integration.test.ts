// A range link that landed as a picture becomes a native table when Excel's
// Convert to table has pushed a table payload under the same link: the next
// update builds the table at the picture's corner and width, tags it and deletes
// the picture in the one batch. A picture nobody converted stays a picture, and
// a refusal anywhere before that batch leaves the picture and its link as they
// were. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_KEY, TAG_LINK, TAG_PAINT } from "../src/link/model";
import type { TableCell } from "../src/link/model";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { settleHungSync } from "./hung-sync";
import { recordsIn, removeAllParts } from "./ppt.deck-store.support";
import { pushAgain, pushTable, seedLink } from "./ppt.support";
import {
  bootUpgrade,
  boxOf,
  moveTo,
  PNG,
  shapesOf,
  tagOf,
  updateAll,
  type UpgradePane,
} from "./ppt.upgrade.support";

enableStrictLoadSemantics();

const LABEL = "Model!B4:F12";
const CELLS: TableCell[][] = [
  [{ t: "Revenue", b: true }, { t: "1 200" }],
  [{ t: "Cost" }, { t: "800", f: "#EEEEEE" }],
];
const WIDTHS = [90, 60];
const CORNER = { left: 80, top: 320, width: 200, height: 40 };

let pane: UpgradePane;
beforeEach(async () => {
  pane = await bootUpgrade();
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

// The picture an old Export selection made: a range link, a picture on the slide.
async function rangePicture() {
  const item = await seedLink(PNG);
  await pane.links.insertFromInbox(item, pane.ws, pane.relay);
  const picture = shapesOf(pane)[0]!;
  moveTo(picture, CORNER);
  return { item, picture };
}

describe("a range picture whose link now carries a table", () => {
  it("becomes a table at the picture's corner and width, tagged as a table link", async () => {
    const { item, picture } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([`${LABEL} now a table`]);
    expect(shapesOf(pane)).toHaveLength(1);
    const table = shapesOf(pane)[0]!;
    expect(table.type).toBe("Table");
    expect(pane.presentation.peekShape(picture.id)).toBeNull();
    expect(table.left).toBe(CORNER.left);
    expect(table.top).toBe(CORNER.top);
    expect(table.width).toBe(CORNER.width);
    expect(
      table.table!.cells.map((row) => row.map((cell) => cell.text)),
    ).toEqual([
      ["Revenue", "1 200"],
      ["Cost", "800"],
    ]);
    expect(table.table!.cells[1]![1]!.fill.color).toBe("#EEEEEE");
    expect(table.tags.get(TAG_KEY)).toBe(item.token);
    expect(table.tags.get(TAG_PAINT)).toBeDefined();
    expect(tagOf(table)).toMatchObject({ id: item.id, kind: "table", rev: 2 });
    const [record, ...rest] = recordsIn(pane.presentation, true);
    expect(rest).toEqual([]);
    expect(record).toMatchObject({
      id: item.id,
      kind: "table",
      rev: 2,
      dataKept: true,
      label: `${LABEL} table`,
    });
    // The next update repaints the table in place, as any table link.
    await pushTable(item, CELLS, WIDTHS);
    expect(await updateAll(pane)).toMatchObject({ updated: 1, failed: 0 });
    expect(shapesOf(pane)[0]!.id).toBe(table.id);
  });

  it("cannot be reverted to the picture it was, and says so", async () => {
    const { item } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);
    await updateAll(pane);
    const table = shapesOf(pane)[0]!;
    const { revertLinks } = await import("../src/ppt/revert");

    const summary = await revertLinks(
      await pane.links.listLinks(pane.relay),
      pane.relay,
    );

    expect(summary).toMatchObject({ reverted: 0, failed: 1 });
    expect(summary.failures[0]).toContain("is a table now");
    expect(shapesOf(pane)[0]!.id).toBe(table.id);
    expect(tagOf(table)).toMatchObject({ kind: "table", rev: 2 });
  });

  it("stays a picture while nobody has converted it", async () => {
    const { item, picture } = await rangePicture();
    expect(await updateAll(pane)).toMatchObject({ updated: 0, current: 1 });
    await pushAgain(item, PNG);

    expect(await updateAll(pane)).toMatchObject({ updated: 1, failed: 0 });
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(picture.setImageCalls).toBe(2);
  });

  it("keeps the picture and the link when the host refuses the table", async () => {
    const { item, picture } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);
    const rows = await pane.links.listLinks(pane.relay);
    pane.helpers.failNextSync(new Error("the host refused the table"));

    const summary = await pane.links.updateLinks(rows, pane.relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toContain("the host refused the table");
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(boxOf(picture)).toEqual(CORNER);
    expect(tagOf(picture)).toMatchObject({
      id: item.id,
      kind: "range",
      rev: 1,
    });
    expect(await pane.links.listLinks(pane.relay)).toHaveLength(1);
  });

  it("takes the half-built table down and builds once more when the record's write halts the swap", async () => {
    const { item, picture } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);
    const rows = await pane.links.listLinks(pane.relay);
    // The record's part is gone from under the pane: its write halts the last
    // batch before the tags and the picture's delete, so the picture is still
    // the link when the first try is taken down.
    await removeAllParts();

    const summary = await pane.links.updateLinks(rows, pane.relay);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapesOf(pane)).toHaveLength(1);
    const table = shapesOf(pane)[0]!;
    expect(table.type).toBe("Table");
    expect(pane.presentation.peekShape(picture.id)).toBeNull();
    expect(tagOf(table)).toMatchObject({ id: item.id, kind: "table", rev: 2 });
    expect(recordsIn(pane.presentation, true)).toHaveLength(1);
  });

  it("keeps the table as the link when the swap landed although the host answered with an error", async () => {
    const { item, picture } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);
    const rows = await pane.links.listLinks(pane.relay);
    // The host applies the last batch (the tags, the picture's delete) and then
    // refuses to say so: the picture is gone, so the new table is the link.
    pane.helpers.failNextSync(new Error("the host answered with an error"), 2);

    const summary = await pane.links.updateLinks(rows, pane.relay);

    expect(summary.failed).toBe(1);
    expect(shapesOf(pane)).toHaveLength(1);
    expect(pane.presentation.peekShape(picture.id)).toBeNull();
    expect(shapesOf(pane)[0]!.type).toBe("Table");
    expect(await pane.links.listLinks(pane.relay)).toHaveLength(1);
  });

  it("keeps the picture and says so when the host stops answering", async () => {
    const { item, picture } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);
    const rows = await pane.links.listLinks(pane.relay);
    pane.helpers.hangNextSync();
    vi.useFakeTimers();

    const summary = await settleHungSync(
      pane.links.updateLinks(rows, pane.relay),
    );
    vi.useRealTimers();

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures).toEqual([
      `${LABEL}: PowerPoint stopped answering; the picture stays`,
    ]);
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(tagOf(picture)).toMatchObject({ kind: "range", rev: 1 });
  });

  // The update's syncs, from its first: the add (1), the wait for the table to
  // register (2), the last format batch with the tags and the swap (3). The
  // clean-up after a refusal opens with its probe and then its own first round
  // trip, the one a host that has stopped answering swallows.
  it("settles with the picture and its link when the build fails and the clean-up's first round trip is swallowed", async () => {
    const { item, picture } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);
    const rows = await pane.links.listLinks(pane.relay);
    pane.helpers.failNextSync(new Error("the host refused the table"), 1);
    pane.helpers.hangNextSync(3);
    vi.useFakeTimers();

    const summary = await settleHungSync(
      pane.links.updateLinks(rows, pane.relay),
    );
    vi.useRealTimers();

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toContain("the host refused the table");
    expect(pane.presentation.peekShape(picture.id)).not.toBeNull();
    expect(tagOf(picture)).toMatchObject({
      id: item.id,
      kind: "range",
      rev: 1,
    });
    expect(picture.tags.get(TAG_KEY)).toBe(item.token);
    expect(await pane.links.listLinks(pane.relay)).toHaveLength(1);
  });

  it("takes only its own half-built table down, never another copy of the link", async () => {
    const { item, picture } = await rangePicture();
    // A second picture of the same link: copy and paste carries both tags.
    const slide = pane.presentation.slides[0]!;
    const copy = pane.presentation.addShape(slide, {
      type: picture.type,
      left: 80,
      top: 120,
      width: 200,
      height: 40,
    });
    for (const [key, value] of picture.tags) copy.tags.set(key, value);
    await pushTable(item, CELLS, WIDTHS);
    const rows = await pane.links.listLinks(pane.relay);
    // The first row's add is applied and its sync refused: no id was read.
    pane.helpers.refuseNextSync(new Error("the host refused the add"), 0);

    const summary = await pane.links.updateLinks(rows, pane.relay);

    expect(summary).toMatchObject({ updated: 1, failed: 1 });
    expect(pane.presentation.peekShape(picture.id)).not.toBeNull();
    const kinds = shapesOf(pane).map((one) => one.type);
    expect(kinds.filter((type) => type === "Table")).toHaveLength(1);
    expect(shapesOf(pane)).toHaveLength(2);
    expect(
      shapesOf(pane).every((one) => one.tags.get(TAG_LINK) !== undefined),
    ).toBe(true);
  });

  it("is refused in words on a host with no table API, picture untouched", async () => {
    const { item, picture } = await rangePicture();
    await pushTable(item, CELLS, WIDTHS);
    pane.helpers.setSupported((_set, version) => version !== "1.8");

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toMatch(/Tables need PowerPoint/);
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
  });
});
