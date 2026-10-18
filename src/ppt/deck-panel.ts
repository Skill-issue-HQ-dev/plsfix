// The deck's own rows in the pane: the scans fed to the settle, the rows for
// records no shape carries, and what their buttons and Redraw do. main.ts owns
// the DOM and the busy latch. Invariant: such a row is shown only once a scan
// has settled, and the one timer goes with the page.

import type { LinkId } from "../link/model";
import { requireSelection, type LinkFilter } from "./actions";
import {
  putBack,
  redrawDetails,
  redrawLinks,
  removeFromDeck,
  summarizeRedraw,
} from "./deck-actions";
import { deckRecords } from "./deck-boot";
import { drawingFrom } from "./deck-drawable";
import { onDeckScan } from "./deck-flow";
import { filterDeckRows, offSlideRows, type DeckRowView } from "./deck-rows";
import type { LinkRow } from "./links";
import type { InsertTarget } from "./placement";
import { createScanSettle, createSettleTimer } from "./scan-settle";
import type { DeckRowsPanel } from "./views";

export interface DeckPanelDeps {
  // The pane's guard: one action at a time, its sentence the toast.
  act: (run: () => Promise<string>, action: string) => void;
  // The slide and spot the Inbox's pickers name.
  target: () => Promise<InsertTarget>;
  // The list read again after an action; the Inbox read again after a removal;
  // and the full reload a rescan runs.
  refresh: () => Promise<void>;
  inbox: () => Promise<void>;
  reload: () => Promise<void>;
  selected: () => LinkRow[];
  busy: () => boolean;
  // The lines an action leaves under its toast.
  details: (text: string | undefined) => void;
  // The pane's own page: the timer goes with it.
  root: HTMLElement;
}

export interface DeckPanel {
  // Told about every list the pane draws, the boot one first.
  scanned(rows: readonly LinkRow[]): void;
  // Arms the one rescan that lets a deck's first scan be believed.
  start(): void;
  // The rows of the records no shape carries, filtered like the list is.
  view(filter: LinkFilter): DeckRowsPanel;
  count(): number;
  sources(): string[];
  projects(): string[];
  redraw(): Promise<string>;
}

type Found = LinkRow["found"];

// One guarded action on a record, the list read again after it - after a
// refusal too, which is usually the list having been out of date. A removal
// also reads the Inbox again: a row the record hid shows once the record is gone.
function onRecord(
  deps: DeckPanelDeps,
  action: string,
  run: (id: LinkId) => Promise<string>,
  inbox = false,
): (id: LinkId) => void {
  return (id) => {
    deps.act(async () => {
      try {
        return await run(id);
      } finally {
        await deps.refresh();
        if (inbox) await deps.inbox();
      }
    }, action);
  };
}

async function redrawTicked(deps: DeckPanelDeps): Promise<string> {
  const subset = requireSelection(deps.selected());
  const summary = await redrawLinks(subset);
  deps.details(redrawDetails(summary));
  await deps.refresh();
  return summarizeRedraw(summary);
}

export function createDeckPanel(deps: DeckPanelDeps): DeckPanel {
  const settle = createScanSettle();
  // Every scan the pane takes, the list's own or an action's, feeds the settle.
  onDeckScan((seen) => settle.observe(seen));
  const timer = createSettleTimer({
    settle,
    rescan: deps.reload,
    busy: deps.busy,
    alive: () => deps.root.isConnected,
  });
  let carried: readonly Found[] = [];
  const offSlide = (): DeckRowView[] =>
    offSlideRows(carried, deckRecords(), settle.isSettled(), drawingFrom);

  return {
    scanned(rows) {
      carried = rows.map((row) => row.found);
    },
    // A deck that holds no records has nothing a rescan could add: a refresh
    // the user makes later is a scan like any other.
    start() {
      if (deckRecords().size === 0) return;
      timer.arm();
      window.addEventListener("pagehide", () => timer.stop(), { once: true });
    },
    view: (filter) => ({
      rows: filterDeckRows(offSlide(), filter),
      onPutBack: onRecord(deps, "deck-put-back", async (id) =>
        putBack(id, await deps.target()),
      ),
      onRemove: onRecord(deps, "deck-remove", removeFromDeck, true),
    }),
    count: () => offSlide().length,
    sources: () => offSlide().map((row) => row.source),
    projects: () =>
      offSlide()
        .map((row) => row.project)
        .filter((name): name is string => name !== undefined),
    redraw: () => redrawTicked(deps),
  };
}
