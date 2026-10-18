// A chart link that older versions put on a slide as a picture becomes a native
// group on the paint that was coming anyway - Update all with a changed source -
// at the picture's corner and width, and the deck's record follows. A picture
// the user chose (the long-draw question's answer, the explicit picture export)
// is never turned into shapes, changed or not; a host that cannot draw the kind,
// a refusal mid-draw and a host that stops answering all keep the picture. Strict
// load semantics are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  PICTURE_CHOSEN,
  SMOOTH_STRAIGHT,
  type ChartData,
} from "../src/link/chart-model";
import { TAG_KEY } from "../src/link/model";
import { settleHungSync } from "./hung-sync";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { recordsIn } from "./ppt.deck-store.support";
import {
  columnChart,
  pushAgain,
  pushChart,
  seedChart,
  seedLink,
} from "./ppt.support";
import {
  bigChart,
  boxOf,
  bootUpgrade,
  chosenPicture,
  clearDrawUi,
  COLUMN,
  GREW,
  LABEL,
  moveTo,
  PNG,
  pictureEraChart,
  shapesOf,
  tagOf,
  updateAll,
  type UpgradePane,
} from "./ppt.upgrade.support";

enableStrictLoadSemantics();

let pane: UpgradePane;
beforeEach(async () => {
  pane = await bootUpgrade();
});
afterEach(async () => {
  vi.useRealTimers();
  await clearDrawUi();
  uninstallFakePpt();
});

const CORNER = { left: 80, top: 320, width: 380, height: 190 };
// A pie needs PowerPointApi 1.10, which one test takes away.
const PIE: ChartData = { ...columnChart(3), kind: "pie", title: "Mix" };

describe("a picture-era chart whose source changed", () => {
  it("becomes a tagged group at the picture's corner and width, and the record follows", async () => {
    const item = await pictureEraChart(pane);
    const picture = shapesOf(pane)[0]!;
    expect(picture.type).not.toBe("Group");
    moveTo(picture, CORNER);
    await pushChart(item, COLUMN, PNG);

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([`${LABEL} now a native chart`]);
    expect(shapesOf(pane)).toHaveLength(1);
    const drawn = shapesOf(pane)[0]!;
    expect(drawn.type).toBe("Group");
    expect(boxOf(drawn)).toEqual(CORNER);
    expect(pane.presentation.peekShape(picture.id)).toBeNull();
    expect(drawn.tags.get(TAG_KEY)).toBe(item.token);
    expect(tagOf(drawn)).toMatchObject({ id: item.id, kind: "chart", rev: 2 });
    expect(tagOf(drawn).choice).toBeUndefined();
    const [record, ...rest] = recordsIn(pane.presentation, true);
    expect(rest).toEqual([]);
    expect(record).toMatchObject({
      id: item.id,
      kind: "chart",
      rev: 2,
      dataKept: true,
    });
    expect(await pane.links.listLinks(pane.relay)).toHaveLength(1);
  });

  // What a drawn chart could not keep joins the line in brackets, as one
  // sentence with the label ("... now a native chart. smooth lines ..." was two).
  it("says what the chart could not keep in brackets after the upgrade", async () => {
    const item = await pictureEraChart(pane);
    await pushChart(
      item,
      {
        ...COLUMN,
        kind: "line",
        series: [{ ...COLUMN.series[0]!, smooth: true }],
      },
      PNG,
    );

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([
      `${LABEL} now a native chart (${SMOOTH_STRAIGHT})`,
    ]);
    expect(shapesOf(pane)[0]!.type).toBe("Group");
  });

  it("is drawn from the new data, not from the picture's size", async () => {
    const item = await pictureEraChart(pane);
    await pushChart(item, GREW, PNG);

    expect(await updateAll(pane)).toMatchObject({ updated: 1, failed: 0 });

    // Seven bars, not six: the group is the new push's chart.
    const rectangles = shapesOf(pane)[0]!.group!.shapes.filter(
      (one) => one.geometry === "Rectangle",
    );
    expect(rectangles).toHaveLength(7);
  });
});

describe("a picture-era chart that is not redrawn", () => {
  it("stays a picture while its source has not changed", async () => {
    await pictureEraChart(pane);
    const picture = shapesOf(pane)[0]!;

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 0, current: 1, failed: 0 });
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(picture.setImageCalls).toBe(1);
  });

  it("is repainted as a picture when the new push carries no chart data", async () => {
    const item = await pictureEraChart(pane);
    const picture = shapesOf(pane)[0]!;
    await pushChart(item, null, PNG, "still no data to read");

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(picture.setImageCalls).toBe(2);
    expect(summary.notes).toEqual([
      `${LABEL} as a picture: still no data to read`,
    ]);
  });

  it("keeps the picture and says why when this host cannot draw the kind", async () => {
    const item = await pictureEraChart(pane);
    const picture = shapesOf(pane)[0]!;
    // A pie needs PowerPointApi 1.10; this PowerPoint stops at 1.9.
    pane.helpers.setSupported((_set, version) => version !== "1.10");
    await pushChart(item, PIE, PNG);

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(picture.setImageCalls).toBe(2);
    expect(summary.notes).toEqual([
      `${LABEL} as a picture: pie shapes need Microsoft 365 2601 / Mac 16.105 or newer`,
    ]);
  });

  it("inside a group the user built is repainted and says to ungroup it", async () => {
    const item = await pictureEraChart(pane);
    const picture = shapesOf(pane)[0]!;
    const slide = pane.presentation.slides[0]!;
    const caption = pane.presentation.addShape(slide, {
      left: 5,
      top: 5,
      width: 60,
    });
    pane.presentation.groupShapes([picture.id, caption.id], slide.id);
    await pushChart(item, COLUMN, PNG);

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(picture.setImageCalls).toBe(2);
    expect(summary.notes).toHaveLength(1);
    expect(summary.notes[0]).toMatch(/ungroup/i);
    expect(
      shapesOf(pane).some(
        (one) => one.fillImage !== null && one.type !== "Group",
      ),
    ).toBe(false);
  });
});

describe("a picture the user chose", () => {
  it("at the long-draw question is never upgraded, changed or not, and the choice outlives a repaint", async () => {
    const item = await chosenPicture(pane);
    const picture = shapesOf(pane)[0]!;
    expect(picture.type).not.toBe("Group");
    expect(tagOf(picture).choice).toBe("picture");

    expect(await updateAll(pane)).toMatchObject({ updated: 0, current: 1 });
    expect(picture.setImageCalls).toBe(1);

    await pushChart(item, bigChart(), PNG);
    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([]);
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(picture.setImageCalls).toBe(2);
    expect(tagOf(picture)).toMatchObject({ rev: 2, choice: "picture" });
    expect(PICTURE_CHOSEN).toBe("Picture (your choice)");
  });

  it("by the explicit picture export is never upgraded, whatever a later push carries", async () => {
    const item = await seedLink(PNG);
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const picture = shapesOf(pane)[0]!;
    await pushChart(item, COLUMN, PNG);

    const summary = await updateAll(pane);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(shapesOf(pane)[0]!.type).not.toBe("Group");
    expect(picture.setImageCalls).toBe(2);
  });

  it("stays the picture the user chose when it is pointed at another source", async () => {
    await chosenPicture(pane);
    const picture = shapesOf(pane)[0]!;
    const other = await seedChart(bigChart(), PNG);
    const { changeSource } = await import("../src/ppt/change-source");
    const [row] = await pane.links.listLinks(pane.relay);

    await changeSource(row!, other, pane.ws, pane.relay);

    expect(tagOf(picture)).toMatchObject({ id: other.id, choice: "picture" });
  });
});

describe("a revert is not an update", () => {
  it("repaints the picture-era chart as the picture it was and draws nothing", async () => {
    // A host with no shape API holds the chart as a picture at two revisions.
    pane.helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.8"),
    );
    const item = await seedChart(COLUMN, PNG);
    await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    await pushChart(item, COLUMN, PNG);
    await updateAll(pane);
    pane.helpers.setSupported(() => true);
    const picture = shapesOf(pane)[0]!;
    const { revertLinks } = await import("../src/ppt/revert");

    const summary = await revertLinks(
      await pane.links.listLinks(pane.relay),
      pane.relay,
    );

    expect(summary).toMatchObject({ reverted: 1, failed: 0 });
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(shapesOf(pane)[0]!.type).not.toBe("Group");
    expect(tagOf(picture)).toMatchObject({ rev: 1 });
  });
});

describe("an upgrade the host refuses or swallows", () => {
  it("keeps the picture and the link when a chunk of the draw is refused, and says so", async () => {
    const item = await pictureEraChart(pane);
    const picture = shapesOf(pane)[0]!;
    moveTo(picture, CORNER);
    await pushChart(item, COLUMN, PNG);
    const rows = await pane.links.listLinks(pane.relay);
    // The upgrade's first sync draws twelve shapes; the host refuses the second.
    pane.helpers.failNextSync(new Error("the host refused the shapes"), 1);

    const summary = await pane.links.updateLinks(rows, pane.relay);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toHaveLength(1);
    expect(summary.notes[0]).toContain("stays a picture");
    // No half-drawn chart: the twelve shapes the first chunk added are gone.
    expect(shapesOf(pane)).toHaveLength(1);
    expect(shapesOf(pane)[0]!.id).toBe(picture.id);
    expect(boxOf(picture)).toEqual(CORNER);
    expect(picture.setImageCalls).toBe(2);
    expect(picture.tags.get(TAG_KEY)).toBe(item.token);
    expect(tagOf(picture)).toMatchObject({ id: item.id, rev: 2 });
    expect(await pane.links.listLinks(pane.relay)).toHaveLength(1);
  });

  it("keeps the picture untouched and says so when the host stops answering", async () => {
    const item = await pictureEraChart(pane);
    const picture = shapesOf(pane)[0]!;
    await pushChart(item, COLUMN, PNG);
    const rows = await pane.links.listLinks(pane.relay);
    pane.helpers.hangNextSync(1);
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
    expect(picture.setImageCalls).toBe(1);
    expect(tagOf(picture)).toMatchObject({ id: item.id, rev: 1 });
    expect(await pane.links.listLinks(pane.relay)).toHaveLength(1);
  });
});

describe("an Update all that upgrades one picture among others", () => {
  it("still repaints the plain pictures in the one round trip they always took", async () => {
    const plain = await Promise.all([0, 1, 2].map(() => seedLink(PNG)));
    for (const one of plain) {
      await pane.links.insertFromInbox(one, pane.ws, pane.relay);
    }
    const era = await pictureEraChart(pane);
    for (const one of plain) await pushAgain(one, PNG);
    await pushChart(era, COLUMN, PNG);
    const rows = await pane.links.listLinks(pane.relay);
    const before = pane.helpers.syncCount();

    const summary = await pane.links.updateLinks(rows, pane.relay);

    expect(summary).toMatchObject({ updated: 4, failed: 0 });
    // The three pictures in one sync, and the chart's own: two chunks (20
    // shapes) and the group.
    expect(pane.helpers.syncCount() - before).toBe(4);
  });
});
