// A host that refuses one round trip of a Put back, whichever it is (the file
// read, the deck scan, any sync of the insert): the deck is as it was and the row
// can be asked again. Every sync is swept for a text, a table and a chart.

import { afterEach, describe, expect, it } from "vitest";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import {
  bootActions,
  insertedThenDeleted,
  type ActionsPane,
} from "./ppt.deck-actions.support";
import { CELLS, PNG, WIDTHS, recordsIn } from "./ppt.deck-store.support";
import { columnChart, seedChart, seedTable, seedText } from "./ppt.support";
import type { InboxItem } from "../src/link/model";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

const SWEPT: { name: string; seed: () => Promise<InboxItem> }[] = [
  { name: "a text", seed: () => seedText("EUR 3.2m") },
  { name: "a table", seed: () => seedTable(CELLS, WIDTHS, true) },
  { name: "a chart", seed: () => seedChart(columnChart(6), PNG) },
];

// A fresh pane with the export inserted and its object deleted by hand.
async function deleted(seed: () => Promise<InboxItem>): Promise<{
  pane: ActionsPane;
  item: InboxItem;
}> {
  const pane = await bootActions();
  const item = await seed();
  await insertedThenDeleted(pane, item);
  return { pane, item };
}

// What a Put back costs: the round trips before the object is drawn (the file
// read, the deck scan) and in all.
async function costs(
  seed: () => Promise<InboxItem>,
): Promise<{ before: number; all: number }> {
  const twin = await deleted(seed);
  const host = await import("../src/ppt/host");
  const readFrom = twin.pane.helpers.syncCount();
  await twin.pane.boot.loadDeckStore();
  await host.scanLinks();
  const before = twin.pane.helpers.syncCount() - readFrom;
  const whole = await deleted(seed);
  const from = whole.pane.helpers.syncCount();
  await whole.pane.actions.putBack(whole.item.id);
  return { before, all: whole.pane.helpers.syncCount() - from };
}

describe("a Put back whose round trip the host refuses", () => {
  it.each(SWEPT)(
    "leaves the deck as it was, for every sync of $name, and can be asked again",
    async ({ seed }) => {
      const { before, all } = await costs(seed);
      // The insert is more than one round trip, or there is nothing to sweep.
      expect(all).toBeGreaterThan(before + 1);

      for (let sync = 0; sync < all; sync += 1) {
        const { pane, item } = await deleted(seed);
        const held = recordsIn(pane.presentation, true);
        pane.helpers.failNextSync(undefined, sync);
        const at = `sync ${String(sync + 1)} of ${String(all)}`;

        const refused = await pane.actions
          .putBack(item.id)
          .then(() => false)
          .catch(() => true);

        const shapes = pane.presentation.slides[0]!.shapes;
        if (refused) {
          expect(shapes, `${at}: nothing half drawn`).toEqual([]);
          expect(recordsIn(pane.presentation, true), `${at}: record`).toEqual(
            held,
          );
          // Asked again, with the host answering: exactly one object.
          await pane.actions.putBack(item.id);
        }
        expect(shapes, `${at}: one object`).toHaveLength(1);
        expect(
          recordsIn(pane.presentation, true),
          `${at}: one record`,
        ).toHaveLength(1);
      }
    },
  );
});
