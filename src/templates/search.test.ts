// The picker's filter: what typing into the Templates search keeps. Pure, so it
// is proven on a small fixed list for the semantics (every word must start a
// word of the name, the description or the group) and on the real registry for
// the one promise that must outlive every new block: a template is always found
// by its own name.
import { describe, expect, it } from "vitest";
import { TEMPLATES } from "./index";
import { filterTemplates } from "./search";

interface Row {
  id: string;
  name: string;
  description: string;
  group: string;
}

const ROWS: Row[] = [
  {
    id: "a",
    name: "DCF valuation",
    description: "Free cash flows to enterprise value.",
    group: "Valuation",
  },
  {
    id: "b",
    name: "Debt schedule",
    description: "Level-payment loan with interest and balance.",
    group: "LBO",
  },
  {
    id: "c",
    name: "EBITDA bridge",
    description: "Opening to closing EBITDA, shaped for the waterfall.",
    group: "Analysis",
  },
  {
    id: "d",
    name: "NPV / IRR",
    description: "An outlay and cash flows into NPV and payback.",
    group: "Valuation",
  },
];

const ids = (rows: Row[]): string[] => rows.map((row) => row.id);

describe("filterTemplates", () => {
  it("keeps every entry, in order, for an empty or blank query", () => {
    expect(ids(filterTemplates(ROWS, ""))).toEqual(["a", "b", "c", "d"]);
    expect(ids(filterTemplates(ROWS, "   "))).toEqual(["a", "b", "c", "d"]);
  });

  it("returns a copy, never the list it was given", () => {
    expect(filterTemplates(ROWS, "")).not.toBe(ROWS);
  });

  it("finds a word of the name", () => {
    expect(ids(filterTemplates(ROWS, "dcf"))).toEqual(["a"]);
    expect(ids(filterTemplates(ROWS, "bridge"))).toEqual(["c"]);
  });

  it("finds a word of the description", () => {
    expect(ids(filterTemplates(ROWS, "waterfall"))).toEqual(["c"]);
    expect(ids(filterTemplates(ROWS, "payback"))).toEqual(["d"]);
  });

  it("finds a group, whole or typed so far", () => {
    expect(ids(filterTemplates(ROWS, "lbo"))).toEqual(["b"]);
    expect(ids(filterTemplates(ROWS, "val"))).toEqual(["a", "d"]);
  });

  it("ignores case and the space around the query", () => {
    expect(ids(filterTemplates(ROWS, "  DcF "))).toEqual(["a"]);
  });

  it("wants every word of the query", () => {
    expect(ids(filterTemplates(ROWS, "cash flows"))).toEqual(["a", "d"]);
    expect(ids(filterTemplates(ROWS, "cash payback"))).toEqual(["d"]);
    expect(ids(filterTemplates(ROWS, "debt dcf"))).toEqual([]);
  });

  it("matches the start of a word, not its middle", () => {
    expect(ids(filterTemplates(ROWS, "ebit"))).toEqual(["c"]);
    expect(ids(filterTemplates(ROWS, "bitda"))).toEqual([]);
  });

  it("reads punctuation as a gap, so a slash or a hyphen does not block a match", () => {
    expect(ids(filterTemplates(ROWS, "npv/irr"))).toEqual(["d"]);
    expect(ids(filterTemplates(ROWS, "level-payment"))).toEqual(["b"]);
  });

  it("keeps the list's order among the matches", () => {
    expect(ids(filterTemplates([...ROWS].reverse(), "val"))).toEqual([
      "d",
      "a",
    ]);
  });

  it("answers an empty list for a stranger", () => {
    expect(filterTemplates(ROWS, "zzz")).toEqual([]);
  });
});

describe("filterTemplates over the registry", () => {
  it("finds every template by its own name", () => {
    for (const entry of TEMPLATES) {
      expect(filterTemplates(TEMPLATES, entry.name), entry.name).toContain(
        entry,
      );
    }
  });

  it("narrows to one template by the start of its name", () => {
    expect(
      filterTemplates(TEMPLATES, "debt sch").map((entry) => entry.id),
    ).toEqual(["template-debt-schedule"]);
  });
});
