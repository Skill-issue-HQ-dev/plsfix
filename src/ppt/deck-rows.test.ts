// The Linked objects list as the deck sees it: the tags scan crossed with the
// deck's own records. What a record without a shape becomes, when a scan may
// say so, how such a row is found by the list's filters and how the rows fall
// into project folders. Pure: no host, no DOM.

import { describe, expect, it } from "vitest";
import { identityOnly, type DeckRecord } from "../link/deck-record";
import { LINK, record, SRC, TEXT } from "../link/deck-record.support";
import type { LinkFilter } from "./actions";
import type { DrawFrom } from "./deck-drawable";
import {
  deckRowView,
  filterDeckRows,
  isDeckRow,
  OFF_SLIDE_WORD,
  offSlideRows,
  unionByProject,
} from "./deck-rows";

const idOf = (n: number): string => n.toString(16).padStart(32, "0");

// A record by number: a text link of its own id and label.
function recordOf(
  n: number,
  made: { project?: string; label?: string } = {},
): DeckRecord {
  return record(TEXT, {
    id: idOf(n),
    kind: "text",
    label: made.label ?? `Model!B${String(n)} text`,
    token: LINK.token,
    sentAt: LINK.sentAt,
    ...(made.project === undefined ? {} : { project: made.project }),
  });
}

// The host's answer stood in for: a record that kept its data can be drawn, one
// that kept the identity only cannot (deck-drawable.test.ts holds the real one).
const draw: DrawFrom = (made) =>
  made.dataKept
    ? {
        ok: true,
        payload: {
          v: 1,
          kind: "text",
          text: "",
          hash: "",
          src: made.src,
          pushedAt: made.pushedAt,
        },
      }
    : {
        ok: false,
        kept: false,
        issue: "this deck kept no data for it",
        way: "Send it again from Excel.",
        tag: "Data not kept in this deck",
      };

const held = (...records: DeckRecord[]): Map<string, DeckRecord> =>
  new Map(records.map((one) => [one.id, one]));
const scanned = (...ids: string[]): { tag: { id: string } }[] =>
  ids.map((id) => ({ tag: { id } }));

describe("a record no shape carries", () => {
  it("becomes a row of its own once the scan is settled", () => {
    const rows = offSlideRows(
      scanned(idOf(1)),
      held(recordOf(1), recordOf(2)),
      true,
      draw,
    );
    expect(rows.map((row) => row.id)).toEqual([idOf(2)]);
    expect(OFF_SLIDE_WORD).toBe("Not on a slide");
  });

  it("is not a row on a scan that has not settled, however empty the scan", () => {
    const records = held(recordOf(1), recordOf(2));
    expect(offSlideRows([], records, false, draw)).toEqual([]);
    expect(offSlideRows(scanned(idOf(1)), records, false, draw)).toEqual([]);
  });

  it("adds nothing when a shape carries it, though it sits on many slides", () => {
    const records = held(recordOf(1));
    const twice = scanned(idOf(1), idOf(1));
    expect(offSlideRows(twice, records, true, draw)).toEqual([]);
  });

  it("is every record when the settled scan found nothing", () => {
    const rows = offSlideRows([], held(recordOf(2), recordOf(1)), true, draw);
    expect(rows).toHaveLength(2);
  });

  it("reads the same whatever order the deck listed the records in", () => {
    const a = recordOf(1, { label: "Alpha" });
    const b = recordOf(2, { label: "Beta" });
    const one = offSlideRows([], held(a, b), true, draw).map((row) => row.id);
    const other = offSlideRows([], held(b, a), true, draw).map((row) => row.id);
    expect(one).toEqual(other);
    expect(one).toEqual([a.id, b.id]);
  });
});

describe("the row a record becomes", () => {
  it("names the link the way a shape row does, with the time of its last push", () => {
    const row = deckRowView(recordOf(7, { project: "Amasty" }), draw);
    expect(row).toMatchObject({
      offSlide: true,
      key: `deck/${idOf(7)}`,
      id: idOf(7),
      label: "Model!B7 text",
      source: SRC.workbook,
      project: "Amasty",
      kind: "text",
      pushedAt: TEXT.pushedAt,
    });
    expect(isDeckRow(row)).toBe(true);
    expect(isDeckRow({ key: "a/b" })).toBe(false);
  });

  it("can be put back when the host can draw it again from the record", () => {
    const row = deckRowView(recordOf(1), draw);
    expect(row.canPutBack).toBe(true);
    expect(row.deadWhy).toBeUndefined();
    expect(row.deadTag).toBeUndefined();
  });

  it("cannot be put back when the host cannot, and says why in its own sentence", () => {
    const lean = deckRowView(identityOnly(recordOf(1)), draw);
    expect(lean).toMatchObject({
      canPutBack: false,
      deadWhy: "This deck kept no data for it. Send it again from Excel.",
      deadTag: "Data not kept in this deck",
    });
    const scatter: DrawFrom = () => ({
      ok: false,
      kept: true,
      issue: "scatter charts are not drawn as shapes yet",
      way: "Send it again from Excel: the Inbox inserts it as a picture.",
      tag: "Cannot be drawn here",
    });
    expect(deckRowView(recordOf(2), scatter)).toMatchObject({
      canPutBack: false,
      deadWhy:
        "Scatter charts are not drawn as shapes yet. Send it again from Excel: the Inbox inserts it as a picture.",
      deadTag: "Cannot be drawn here",
    });
  });

  it("leaves the project out when the link has none", () => {
    expect("project" in deckRowView(recordOf(1), draw)).toBe(false);
  });
});

describe("the list's filters over these rows", () => {
  const rows = [
    deckRowView(
      recordOf(1, { project: "Amasty", label: "Revenue bridge" }),
      draw,
    ),
    deckRowView(recordOf(2, { label: "Cash flow" }), draw),
  ];
  const all: LinkFilter = { query: "", status: "all" };
  const ids = (found: ReturnType<typeof filterDeckRows>): string[] =>
    found.map((row) => row.id);

  it("shows every row under no filter", () => {
    expect(ids(filterDeckRows(rows, all))).toEqual([idOf(1), idOf(2)]);
  });

  it("finds a row by its label, workbook, project, kind or the words Not on a slide", () => {
    const find = (query: string): string[] =>
      ids(filterDeckRows(rows, { ...all, query }));
    expect(find("bridge")).toEqual([idOf(1)]);
    expect(find(" CASH ")).toEqual([idOf(2)]);
    expect(find("model_v4")).toEqual([idOf(1), idOf(2)]);
    expect(find("amasty")).toEqual([idOf(1)]);
    expect(find("text")).toEqual([idOf(1), idOf(2)]);
    expect(find("not on a slide")).toEqual([idOf(1), idOf(2)]);
    expect(find("nothing like this")).toEqual([]);
  });

  it("follows the source and project filters", () => {
    expect(ids(filterDeckRows(rows, { ...all, source: "Other.xlsx" }))).toEqual(
      [],
    );
    expect(ids(filterDeckRows(rows, { ...all, source: SRC.workbook }))).toEqual(
      [idOf(1), idOf(2)],
    );
    expect(ids(filterDeckRows(rows, { ...all, project: "Amasty" }))).toEqual([
      idOf(1),
    ]);
    // "No project" is the empty name.
    expect(ids(filterDeckRows(rows, { ...all, project: "" }))).toEqual([
      idOf(2),
    ]);
    expect(ids(filterDeckRows(rows, { ...all, project: "all" }))).toHaveLength(
      2,
    );
  });

  it("hides them under a slide or a relay status, which they have neither of", () => {
    expect(filterDeckRows(rows, { ...all, slide: 2 })).toEqual([]);
    expect(filterDeckRows(rows, { ...all, slide: "all" })).toHaveLength(2);
    expect(filterDeckRows(rows, { ...all, status: "current" })).toEqual([]);
    expect(filterDeckRows(rows, { ...all, status: "notPasted" })).toEqual([]);
  });
});

describe("project folders", () => {
  const shape = (project: string | undefined, key: string) => ({
    key,
    project,
  });

  it("puts a deck row in its project's folder after that folder's shape rows", () => {
    const groups = unionByProject(
      [shape("Beta", "b1"), shape(undefined, "n1"), shape("Alpha", "a1")],
      [
        deckRowView(recordOf(1, { project: "Alpha" }), draw),
        deckRowView(recordOf(2), draw),
      ],
    );
    expect(
      groups.map(([name, rows]) => [
        name,
        rows.map((row) => ("offSlide" in row ? row.id : row.key)),
      ]),
    ).toEqual([
      ["Alpha", ["a1", idOf(1)]],
      ["Beta", ["b1"]],
      ["No project", ["n1", idOf(2)]],
    ]);
  });

  it("opens a folder for a project only the deck's records name", () => {
    const groups = unionByProject(
      [],
      [deckRowView(recordOf(1, { project: "Zeta" }), draw)],
    );
    expect(groups.map(([name]) => name)).toEqual(["Zeta"]);
  });

  it("keeps the shape rows' own order inside a folder", () => {
    const groups = unionByProject(
      [shape("A", "second"), shape("A", "first")],
      [],
    );
    expect(groups[0]![1].map((row) => (row as { key: string }).key)).toEqual([
      "second",
      "first",
    ]);
  });
});
