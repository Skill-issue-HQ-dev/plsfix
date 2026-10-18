// @vitest-environment jsdom
// The Linked objects list built from the deck's own records: an object deleted by
// hand reads "Not on a slide" once the scan has settled and never before; Put back
// returns it with its link and the Inbox leaves its row to the list (the settle
// itself, the timer and Redraw are ppt.deck-rows.settle). Strict load semantics.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_KEY, TAG_LINK } from "../src/link/model";
import { SCAN_SETTLE_MS } from "../src/ppt/scan-settle";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import {
  afterBoot,
  deckButton,
  deletedByHand,
  insertFirstWaiting,
  offSlideRows,
  pushedText,
  reannounced,
  refreshedAfterTheSettle,
  waitingChart,
  waitingText,
} from "./ppt.deck-rows.support";
import { scatterChart } from "./ppt.support";
import {
  bootPane,
  click,
  deck,
  linkRows,
  paneRelay,
  settle,
  tickRow,
  toastText,
  waiting,
} from "./stress.ppt.support";
import type * as RelayModule from "../src/link/relay";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return paneRelay();
    },
  };
});

beforeEach(() => {
  // Only the clocks (Date and performance): the pane's own awaits run on real turns.
  vi.useFakeTimers({ toFake: ["Date", "performance"] });
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
  vi.restoreAllMocks();
});

const textOf = (index: number): string | null =>
  deck().slides[index]!.shapes[0]?.text ?? null;

describe("an object deleted by hand", () => {
  it("reads Not on a slide once the scan has settled, with its Put back", async () => {
    await bootPane();
    await deletedByHand();

    await refreshedAfterTheSettle();

    expect(linkRows()).toHaveLength(0);
    expect(offSlideRows()).toHaveLength(1);
    const row = offSlideRows()[0]!;
    expect(row.querySelector(".badge")!.textContent).toBe("Not on a slide");
    expect(row.querySelector(".link-meta")!.textContent).toBe(
      "Model_v4.xlsx · Text",
    );
    expect(deckButton("deck-put-back")!.textContent).toBe("Put back");
    // A list with only such a row is not an empty list.
    expect(document.getElementById("links-empty")!.hidden).toBe(true);
    expect(toastText()).toBe("0 linked objects, 1 not on a slide.");
  });

  it("is not a row before the scan has settled, however the user refreshes: one late scan is not two", async () => {
    await bootPane();
    await deletedByHand();

    click("refresh-links");
    await settle();
    afterBoot(SCAN_SETTLE_MS - 1_000);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);

    // The first late scan agrees with the one at boot: still no row.
    afterBoot(2_000);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(0);
    expect(toastText()).toBe("No linked objects in this deck.");

    // The second one, a full interval later, settles it.
    afterBoot(SCAN_SETTLE_MS);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toHaveLength(1);
  });

  it("is Put back at the active slide with its own id and token, and Update all then refreshes it", async () => {
    await bootPane();
    const item = await deletedByHand();
    await refreshedAfterTheSettle();

    deckButton("deck-put-back")!.click();
    await settle();

    expect(toastText()).toBe(`Put ${item.label} back.`);
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(1);
    const back = deck().slides[0]!.shapes[0]!;
    expect(JSON.parse(back.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      rev: 1,
    });
    expect(back.tags.get(TAG_KEY)).toBe(item.token);
    expect(textOf(0)).toBe("EUR 3.2m");

    afterBoot(1_000);
    await pushedText(item, "EUR 3.9m");
    click("refresh-links");
    await settle();
    expect(linkRows()[0]!.textContent).toContain("Update available");
    click("update-all");
    await settle();
    expect(toastText()).toBe("1 updated");
    expect(textOf(0)).toBe("EUR 3.9m");
  });

  it("is Put back on the slide and in the spot the Inbox pickers name", async () => {
    await bootPane();
    await deletedByHand();
    await refreshedAfterTheSettle();
    click("refresh-inbox");
    await settle();
    const slide = document.getElementById("insert-slide") as HTMLSelectElement;
    const where = document.getElementById("insert-where") as HTMLSelectElement;
    slide.value = "3";
    where.value = "top-right";

    deckButton("deck-put-back")!.click();
    await settle();

    const [back] = deck().slides[2]!.shapes;
    expect(back).toBeDefined();
    expect(back!.left + back!.width / 2).toBeGreaterThan(480);
    expect(deck().slides[0]!.shapes).toEqual([]);
    expect(deck().selectedSlideIds).toEqual([deck().slides[2]!.id]);
  });

  it("is not drawn twice when an undo brought it back after the row was shown", async () => {
    await bootPane();
    const item = await waitingText("EUR 3.2m");
    await insertFirstWaiting();
    const shape = deck().slides[0]!.shapes[0]!;
    const tag = shape.tags.get(TAG_LINK)!;
    deck().deleteShape(shape.id);
    await refreshedAfterTheSettle();
    expect(offSlideRows()).toHaveLength(1);
    // Ctrl+Z: the object is back; the list the user is looking at is stale.
    const undone = deck().addShape(deck().slides[0]!, {
      type: "TextBox",
      text: "EUR 3.2m",
      left: 50,
      top: 50,
      width: 200,
      height: 40,
    });
    undone.tags.set(TAG_LINK, tag);
    undone.tags.set(TAG_KEY, item.token);

    deckButton("deck-put-back")!.click();
    await settle();

    expect(toastText()).toMatch(/^Model!B4:F12 text is already on slide 1\./);
    expect(deck().slides[0]!.shapes).toHaveLength(1);
    // The list is read again, so the row gives way to the object.
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(1);
  });

  it("is not listed for a link that was broken: Break took the record with it", async () => {
    await bootPane();
    await waitingText("EUR 3.2m");
    await insertFirstWaiting();
    tickRow(0);
    click("break-selected");
    click("break-selected");
    await settle();
    expect(toastText()).toMatch(/^1 link broken\./);

    await refreshedAfterTheSettle();

    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(0);
    expect(toastText()).toBe("No linked objects in this deck.");
    // The object stays on the slide, unlinked.
    expect(deck().slides[0]!.shapes).toHaveLength(1);
  });

  it("is found by the search box under the words Not on a slide, and by no slide or status filter", async () => {
    await bootPane();
    await deletedByHand();
    await refreshedAfterTheSettle();
    const search = document.getElementById("link-search") as HTMLInputElement;

    search.value = "not on a slide";
    search.dispatchEvent(new Event("input"));
    expect(offSlideRows()).toHaveLength(1);
    search.value = "nothing like it";
    search.dispatchEvent(new Event("input"));
    expect(offSlideRows()).toHaveLength(0);
    expect(document.getElementById("links-filtered-empty")!.hidden).toBe(false);
    search.value = "";
    search.dispatchEvent(new Event("input"));

    const status = document.getElementById(
      "link-status-filter",
    ) as HTMLSelectElement;
    status.value = "current";
    status.dispatchEvent(new Event("change"));
    expect(offSlideRows()).toHaveLength(0);
    status.value = "all";
    status.dispatchEvent(new Event("change"));
    expect(offSlideRows()).toHaveLength(1);
  });

  it("is in its project's folder, and the source and project filters can pick it", async () => {
    await bootPane();
    await deletedByHand("EUR 3.2m", "Amasty");
    await refreshedAfterTheSettle();

    const folders = [...document.querySelectorAll(".link-folder")].map(
      (folder) => folder.textContent,
    );
    expect(folders).toEqual(["Amasty · 1"]);
    const optionsOf = (id: string): string[] =>
      [...(document.getElementById(id) as HTMLSelectElement).options].map(
        (option) => option.value,
      );
    // Neither name belongs to any object on a slide: only the record carries it.
    expect(optionsOf("link-project-filter")).toContain("Amasty");
    expect(optionsOf("link-source-filter")).toContain("Model_v4.xlsx");

    const pick = (id: string, value: string): void => {
      const select = document.getElementById(id) as HTMLSelectElement;
      select.value = value;
      select.dispatchEvent(new Event("change"));
    };
    pick("link-project-filter", "");
    expect(offSlideRows()).toHaveLength(0);
    pick("link-project-filter", "Amasty");
    expect(offSlideRows()).toHaveLength(1);
    pick("link-source-filter", "Model_v4.xlsx");
    expect(offSlideRows()).toHaveLength(1);
  });

  it("leaves Update all, Break and the other toolbar buttons to the objects that are on slides", async () => {
    await bootPane();
    await deletedByHand();
    await refreshedAfterTheSettle();

    click("update-all");
    await settle();
    expect(toastText()).toBe("No links found");
    click("break-selected");
    click("break-selected");
    await settle();
    expect(toastText()).toBe("Tick a link in the list first.");
    expect(offSlideRows()).toHaveLength(1);
  });
});

describe("the Inbox", () => {
  it("leaves the row of a deleted text to the list, which can put it back", async () => {
    await bootPane();
    const item = await deletedByHand();
    await reannounced(item);

    click("refresh-inbox");
    await settle();

    expect(document.querySelectorAll(".inbox-row")).toHaveLength(0);
  });

  it("keeps the row of a deleted picture: the deck holds its identity only, and the Inbox is the way back", async () => {
    await bootPane();
    const item = await waiting();
    await insertFirstWaiting();
    deck().deleteShape(deck().slides[0]!.shapes[0]!.id);
    await reannounced(item);

    click("refresh-inbox");
    await settle();

    expect(document.querySelectorAll(".inbox-row")).toHaveLength(1);
  });

  it("keeps the row of a deleted chart this host cannot draw as shapes, and its Put back is dead with the reason", async () => {
    await bootPane();
    const item = await waitingChart(scatterChart(4));
    await insertFirstWaiting();
    deck().deleteShape(deck().slides[0]!.shapes[0]!.id);
    await reannounced(item);
    await refreshedAfterTheSettle();

    click("refresh-inbox");
    await settle();

    // The deck holds the chart's data, but no host draws a scatter: the Inbox
    // row is the way back, and the list says why it has none of its own.
    expect(document.querySelectorAll(".inbox-row")).toHaveLength(1);
    expect(offSlideRows()).toHaveLength(1);
    const put = deckButton("deck-put-back")!;
    expect(put.disabled).toBe(true);
    expect(put.title).toMatch(/^Scatter charts are not drawn as shapes yet\./);
    expect(put.title).not.toMatch(/this PowerPoint/);

    // Sending it again changes something: the Inbox inserts the picture.
    document.querySelector<HTMLButtonElement>(".inbox-insert")!.click();
    await settle();
    expect(linkRows()).toHaveLength(1);
    expect(offSlideRows()).toEqual([]);
  });
});
