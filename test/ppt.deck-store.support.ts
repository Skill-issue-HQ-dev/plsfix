// What the deck store suites share: records of every sort by number, parts
// seeded into the fake deck by hand (as a past session or a co-author left
// them), and the deck's parts read back as records.

import {
  buildDeckRecord,
  decodeDeckRecord,
  DECK_NAMESPACE,
  encodeDeckRecord,
  type DeckRecord,
} from "../src/link/deck-record";
import type {
  InboxItem,
  Source,
  TablePayload,
  TextPayload,
} from "../src/link/model";
import { fakePng } from "./fakepng";
import {
  customXmlCommitted,
  customXmlSnapshot,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  columnChart,
  pushAgain,
  pushChart,
  pushTable,
  pushText,
  seedChart,
  seedLink,
  seedTable,
  seedText,
} from "./ppt.support";

export const SRC: Source = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

// A link id by number: 32 lowercase hex, the way newLinkId writes one.
export function idOf(n: number): string {
  return n.toString(16).padStart(32, "0");
}

export function textPayload(text: string, pushedAt: string): TextPayload {
  return {
    v: 1,
    kind: "text",
    text,
    src: SRC,
    pushedAt,
    hash: "0".repeat(64),
  };
}

export function tablePayload(
  rows: number,
  cols: number,
  fill: (row: number, col: number) => string = (r, c) => `${r}-${c}`,
): TablePayload {
  return {
    v: 1,
    kind: "table",
    rows,
    cols,
    cells: Array.from({ length: rows }, (_row, r) =>
      Array.from({ length: cols }, (_col, c) => ({ t: fill(r, c) })),
    ),
    widths: Array.from({ length: cols }, () => 60),
    src: SRC,
    pushedAt: "2026-10-02T14:02:00.000Z",
    hash: "1".repeat(64),
  };
}

interface Made {
  rev?: number;
  text?: string;
  savedAt?: string;
}

// A text link's record by number: its text and revision readable at a glance.
export function recordOf(n: number, made: Made = {}): DeckRecord {
  const rev = made.rev ?? 1;
  const pushedAt = `2026-10-02T14:0${String(rev % 10)}:00.000Z`;
  return buildDeckRecord(
    {
      id: idOf(n),
      kind: "text",
      label: `Model!B${String(n)} text`,
      token: `token-${String(n)}`,
      sentAt: "2026-10-02T09:00:00.000Z",
    },
    textPayload(made.text ?? `value ${String(n)}`, pushedAt),
    rev,
    made.savedAt ?? "2026-10-03T08:00:00.000Z",
  );
}

export function xmlOf(n: number, made: Made = {}): string {
  return encodeDeckRecord(recordOf(n, made));
}

// Parts put into the deck the way a past session left them: written and
// synced before the pane under test ever ran.
export async function seedParts(xmls: string[]): Promise<void> {
  await PowerPoint.run(async (context) => {
    for (const xml of xmls) context.presentation.customXmlParts.add(xml);
    await context.sync();
  });
}

// What the deck holds in the deck-link namespace: the parts as records, the
// ones this build cannot read left out.
export function recordsIn(
  presentation: FakePresentation,
  committed = false,
): DeckRecord[] {
  const parts = committed
    ? customXmlCommitted(presentation)
    : customXmlSnapshot(presentation);
  return parts
    .filter((part) => part.namespaceUri === DECK_NAMESPACE)
    .flatMap((part) => {
      const decoded = decodeDeckRecord(part.xml);
      return decoded.ok ? [decoded.record] : [];
    });
}

// A part of the namespace this pane cannot read, which leaves `room` bytes of
// the 4 MB once it is read the way the boot reads it.
export async function fillTheDeck(room = 400): Promise<void> {
  const store = await import("../src/ppt/deck-store");
  const full = `<pad xmlns="urn:plsfix:deck-link:1" v="1" id="${"0".repeat(32)}">${"A".repeat(4 * 1024 * 1024 - room)}</pad>`;
  await PowerPoint.run(async (context) => {
    context.presentation.customXmlParts.add(full);
    await context.sync();
    await store.readDeckRecords(context);
  });
}

// A co-author's pane removes every part from under this one.
export async function removeAllParts(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const parts = context.presentation.customXmlParts;
    parts.load("items/id");
    await context.sync();
    for (const part of parts.items) part.delete();
    await context.sync();
  });
}

export function partsIn(presentation: FakePresentation): string[] {
  return customXmlSnapshot(presentation).map((part) => part.xml);
}

// The five sorts of link the deck keeps (or, for a picture, only names).
export const PNG = fakePng(800, 400);
export const CELLS = [
  [{ t: "Revenue" }, { t: "1 200" }],
  [{ t: "Cost" }, { t: "800" }],
];
export const NEXT_CELLS = [
  [{ t: "Revenue" }, { t: "1 450" }],
  [{ t: "Cost" }, { t: "910" }],
];
export const WIDTHS = [90, 60];

// A host that predates custom XML parts (PowerPointApi 1.7), everything else
// the same: the deck can keep no records and the flows run as they always did.
export function withoutCustomXml(fake: FakePptHelpers): void {
  fake.setSupported(
    (set, version) => !(set === "PowerPointApi" && version === "1.7"),
  );
}

export interface Kind {
  name: string;
  seed: () => Promise<InboxItem>;
  push: (item: InboxItem) => Promise<void>;
  kind: string;
  kept: boolean;
}

export const KINDS: Kind[] = [
  {
    name: "a text",
    seed: () => seedText("EUR 3.2m"),
    push: (item) => pushText(item, "EUR 3.9m"),
    kind: "text",
    kept: true,
  },
  {
    name: "a table",
    seed: () => seedTable(CELLS, WIDTHS),
    push: (item) => pushTable(item, NEXT_CELLS, WIDTHS),
    kind: "table",
    kept: true,
  },
  {
    name: "a chart drawn as shapes",
    seed: () => seedChart(columnChart(6), PNG),
    push: (item) => pushChart(item, columnChart(6), PNG),
    kind: "chart",
    kept: true,
  },
  {
    name: "a range picture",
    seed: () => seedLink(PNG),
    push: (item) => pushAgain(item, fakePng(1600, 800)),
    kind: "range",
    kept: false,
  },
  {
    name: "a chart Excel could not describe",
    seed: () => seedChart(null, PNG, "chart type not supported"),
    push: (item) => pushChart(item, null, PNG, "chart type not supported"),
    kind: "chart",
    kept: false,
  },
];
