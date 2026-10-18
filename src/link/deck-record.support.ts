// What the deck record suites share: one link's identity, a payload of every
// kind, a record built from them and the helpers that say a part by hand or
// read what a refusal said.

import type { ChartData } from "./chart-model";
import { toBase64Url } from "./crypto";
import {
  buildDeckRecord,
  decodeDeckRecord,
  DECK_NAMESPACE,
  type DeckLinkIdentity,
  type DeckRecord,
} from "./deck-record";
import type {
  PicturePayload,
  Source,
  TableCell,
  TablePayload,
  TextPayload,
} from "./model";

export const ID = "0123456789abcdef0123456789abcdef";
export const SRC: Source = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_01234567",
};
export const PUSHED = "2026-10-02T14:02:00.000Z";
export const SAVED = "2026-10-03T08:15:00.000Z";
export const HASH = "a".repeat(64);

export const LINK: DeckLinkIdentity = {
  id: ID,
  kind: "table",
  label: "Model!B4:F12 table",
  project: "Amasty",
  token: "t".repeat(43),
  sentAt: "2026-10-02T09:00:00.000Z",
};

export function table(
  rows: number,
  cols: number,
  text: (row: number, col: number) => string = (r, c) => `${r}-${c}`,
): TablePayload {
  const cells: TableCell[][] = Array.from({ length: rows }, (_row, r) =>
    Array.from({ length: cols }, (_col, c) => ({ t: text(r, c) })),
  );
  return {
    v: 1,
    kind: "table",
    rows,
    cols,
    cells,
    widths: Array.from({ length: cols }, () => 60),
    src: SRC,
    pushedAt: PUSHED,
    hash: HASH,
    h: true,
  };
}

export const TEXT: TextPayload = {
  v: 1,
  kind: "text",
  text: "EUR 3.2m",
  src: SRC,
  pushedAt: PUSHED,
  hash: HASH,
};

export const CHART: ChartData = {
  v: 1,
  kind: "column",
  title: "Revenue",
  categories: ["2024A", "2025E"],
  series: [
    {
      name: "Revenue",
      values: [100, 120],
      labels: ["100", "120"],
      colors: ["#2EC4B6", "#2EC4B6"],
    },
  ],
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
};

export function picture(extra: Partial<PicturePayload> = {}): PicturePayload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png: "iVBORw0KGgo=",
    src: SRC,
    pushedAt: PUSHED,
    hash: HASH,
    ...extra,
  };
}

export function record(
  payload: Parameters<typeof buildDeckRecord>[1],
  link: DeckLinkIdentity = LINK,
): DeckRecord {
  return buildDeckRecord(link, payload, 7, SAVED);
}

export const ATTRS = `xmlns="${DECK_NAMESPACE}" v="1" id="${ID}"`;

// A part as another writer could have left it: the element as given, its body
// the base64url of whatever JSON the test wants to say.
export function part(json: unknown, attrs = ATTRS): string {
  const body = toBase64Url(new TextEncoder().encode(JSON.stringify(json)));
  return `<link ${attrs}>${body}</link>`;
}

// A record as the plain JSON a part carries.
export function wire(source: DeckRecord): Record<string, unknown> {
  return JSON.parse(JSON.stringify(source)) as Record<string, unknown>;
}

export function refusal(xml: string): { reason: string; message: string } {
  const decoded = decodeDeckRecord(xml);
  if (decoded.ok) throw new Error("expected the part to be refused");
  return { reason: decoded.reason, message: decoded.message };
}
