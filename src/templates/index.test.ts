// The one test that walks EVERY template block through the registry: ids and
// groups, the grid's shape, every placeholder resolving inside its own block,
// every formula an assignment, labels with text, number formats from the
// allowed set, no custom function, no dependency loop (the static walk in
// test/block-grid.ts, proven there on planted loops) and a help sentence for
// each Insert button.
// A new block earns all of it by being listed, with no test of its own needed
// for the contract (its own file proves what it computes).
import { describe, expect, it } from "vitest";
import { findLoop } from "../../test/block-grid";
import { HELP } from "../help/copy";
import {
  build,
  label,
  placeholderRefs,
  resolveFormula,
  TEMPLATE_FORMATS,
  type TemplateCell,
} from "../template-cells";
import {
  entryById,
  inPickerOrder,
  TEMPLATE_GROUPS,
  type TemplateEntry,
  type TemplateGroup,
  TEMPLATES,
} from "./index";

const TEMPLATES_HELP_KEY = "templates-heading";
// The ids the Tools tab has shipped since v2.3: a rename would orphan the
// ribbon rig's presses, the click-through suites and every video capture.
const SHIPPED_IDS = [
  "template-debt-schedule",
  "template-dcf",
  "template-npv-irr",
  "template-working-capital",
  "template-sensitivity",
  "template-ebitda-bridge",
];
// The valuation blocks the T1 slice added: they carry their own test files for
// what they compute, and are listed here so the registry cannot lose one.
const VALUATION_IDS = [
  "template-wacc",
  "template-trading-comps",
  "template-precedents",
  "template-football-field",
];
// The LBO blocks the T2 slice added, each with its own test file for what it
// computes; listed here so the registry cannot lose one.
const LBO_IDS = [
  "template-sources-uses",
  "template-debt-paydown",
  "template-returns",
  "template-cap-table",
];
// A picker row is a name and a line: past this a description wraps to a third
// line in the narrowest dock (320 px holds about 33 characters a line).
const DESCRIPTION_MAX_CHARS = 64;
// A block is a screenful: past this it is a sheet, and the Undo capture and the
// empty-block check are no longer a block's cost.
const BLOCK_CELL_CEILING = 500;

function cellsOf(
  rows: TemplateCell[][],
): { cell: TemplateCell; row: number; column: number }[] {
  return rows.flatMap((cells, row) =>
    cells.map((cell, column) => ({ cell, row, column })),
  );
}

// An entry that is only an id in a group: inPickerOrder reads nothing else.
const stub = (group: TemplateGroup, id: string): TemplateEntry => ({
  id,
  group,
  name: id,
  description: id,
  block: build(id, id, id, [[label(id)]]),
});

describe("inPickerOrder", () => {
  const appended = [
    stub("Analysis", "a1"),
    stub("Valuation", "v1"),
    stub("Operating", "o1"),
    stub("Analysis", "a2"),
    stub("LBO", "l1"),
    stub("Valuation", "v2"),
    stub("Operating", "o2"),
  ];

  it("brings entries appended out of order into the picker's groups, authoring order kept inside each", () => {
    const ids = inPickerOrder(appended).map((one) => one.id);
    expect(ids).toEqual(["v1", "v2", "l1", "o1", "o2", "a1", "a2"]);
  });

  it("leaves the list it is given as it was", () => {
    const before = appended.map((one) => one.id);
    inPickerOrder(appended);
    expect(appended.map((one) => one.id)).toEqual(before);
  });

  it("is what the registry lists, so a block added anywhere lands in its group", () => {
    expect(TEMPLATES).toEqual(inPickerOrder(TEMPLATES));
    const groups = TEMPLATES.map((one) => one.group);
    expect(groups).toEqual(
      [...groups].sort(
        (a, b) => TEMPLATE_GROUPS.indexOf(a) - TEMPLATE_GROUPS.indexOf(b),
      ),
    );
  });
});

describe("the template registry", () => {
  it("lists every entry under a unique id and a unique block", () => {
    const ids = TEMPLATES.map((entry) => entry.id);
    expect(new Set(ids).size).toBe(ids.length);
    const blocks = TEMPLATES.map((entry) => entry.block.id);
    expect(new Set(blocks).size).toBe(blocks.length);
  });

  it("keeps the six ids the Tools tab has always carried", () => {
    const ids = TEMPLATES.map((entry) => entry.id);
    expect(ids).toEqual(expect.arrayContaining(SHIPPED_IDS));
  });

  it("lists the four valuation blocks under Valuation, after the DCF", () => {
    const ids = TEMPLATES.map((entry) => entry.id);
    for (const id of VALUATION_IDS) {
      expect(ids, id).toContain(id);
      expect(entryById(id)?.group, id).toBe("Valuation");
    }
    expect(ids.indexOf("template-dcf")).toBeLessThan(
      ids.indexOf("template-wacc"),
    );
  });

  it("lists the LBO blocks under LBO, after the debt schedule", () => {
    const ids = TEMPLATES.map((entry) => entry.id);
    for (const id of LBO_IDS) {
      expect(ids, id).toContain(id);
      expect(entryById(id)?.group, id).toBe("LBO");
      expect(ids.indexOf("template-debt-schedule"), id).toBeLessThan(
        ids.indexOf(id),
      );
    }
  });

  it("makes each id the block's own id behind the template- prefix", () => {
    for (const entry of TEMPLATES) {
      expect(entry.id).toBe(`template-${entry.block.id}`);
    }
  });

  it("puts every entry in a group the picker lists", () => {
    for (const entry of TEMPLATES) {
      expect(TEMPLATE_GROUPS, entry.id).toContain(entry.group);
    }
  });

  it("lists the groups in the order the picker shows them", () => {
    expect(TEMPLATE_GROUPS).toEqual([
      "Valuation",
      "LBO",
      "Operating",
      "Analysis",
    ]);
  });

  it("gives every listed group at least one template", () => {
    const used = new Set(TEMPLATES.map((entry) => entry.group));
    for (const group of TEMPLATE_GROUPS) expect(used, group).toContain(group);
  });

  it("names each entry as its block does", () => {
    for (const entry of TEMPLATES) {
      expect(entry.name, entry.id).toBe(entry.block.name);
    }
  });

  it("describes each entry in one short line", () => {
    for (const entry of TEMPLATES) {
      const where = `${entry.id}: ${entry.description}`;
      expect(entry.description.trim(), where).not.toBe("");
      expect(entry.description, where).not.toContain("\n");
      expect(entry.description.length, where).toBeLessThanOrEqual(
        DESCRIPTION_MAX_CHARS,
      );
    }
  });

  it("finds an entry by its id and answers null for a stranger", () => {
    expect(entryById("template-dcf")?.name).toBe("DCF valuation");
    expect(entryById("template-nope")).toBeNull();
  });
});

describe.each(TEMPLATES)("$id", (entry) => {
  const { block } = entry;

  it("is a rectangle of the size it declares", () => {
    expect(block.rows).toBeGreaterThan(0);
    expect(block.cols).toBeGreaterThan(0);
    expect(block.cells).toHaveLength(block.rows);
    for (const row of block.cells) expect(row).toHaveLength(block.cols);
    expect(block.rows * block.cols).toBeLessThanOrEqual(BLOCK_CELL_CEILING);
  });

  it("resolves every placeholder inside its own block", () => {
    let formulas = 0;
    for (const { cell, row, column } of cellsOf(block.cells)) {
      if (cell.f === undefined) continue;
      formulas += 1;
      const where = `${entry.id} r${String(row)}c${String(column)}: ${cell.f}`;
      for (const ref of placeholderRefs(cell.f)) {
        expect(ref.row, where).toBeGreaterThanOrEqual(0);
        expect(ref.row, where).toBeLessThan(block.rows);
        expect(ref.column, where).toBeGreaterThanOrEqual(0);
        expect(ref.column, where).toBeLessThan(block.cols);
      }
      expect(resolveFormula(cell.f, { row: 0, column: 0 }), where).not.toMatch(
        /[{}]/,
      );
    }
    expect(formulas).toBeGreaterThan(0);
  });

  it("starts every formula with = and never calls a custom function", () => {
    for (const { cell, row, column } of cellsOf(block.cells)) {
      if (cell.f === undefined) continue;
      const where = `${entry.id} r${String(row)}c${String(column)}: ${cell.f}`;
      expect(cell.f.startsWith("="), where).toBe(true);
      expect(cell.f.toUpperCase(), where).not.toContain("PLSFIX");
    }
  });

  it("has no circular reference: no formula reaches itself through the block", () => {
    expect(findLoop(block), `${entry.id} has a dependency loop`).toEqual([]);
  });

  it("carries a value or a formula in no cell twice", () => {
    for (const { cell } of cellsOf(block.cells)) {
      expect(cell.v === undefined || cell.f === undefined).toBe(true);
    }
  });

  it("gives every typed text a word, and every number a format from the allowed set", () => {
    for (const { cell, row, column } of cellsOf(block.cells)) {
      const where = `${entry.id} r${String(row)}c${String(column)}`;
      if (typeof cell.v === "string") {
        expect(cell.v.trim(), where).not.toBe("");
        expect(cell.format, where).toBeUndefined();
      }
      const numeric = typeof cell.v === "number" || cell.f !== undefined;
      if (numeric) expect(TEMPLATE_FORMATS, where).toContain(cell.format);
    }
  });

  // The adapter writes typed text through `range.formulas`, which reads it like
  // typing: a label opening with = + - or @ becomes a formula, a bare number
  // becomes a number. "Less: depreciation" is safe where "- depreciation" is not.
  it("types no text a cell would read as a formula or a number", () => {
    for (const { cell, row, column } of cellsOf(block.cells)) {
      if (typeof cell.v !== "string") continue;
      const where = `${entry.id} r${String(row)}c${String(column)}: ${cell.v}`;
      expect(cell.v, where).not.toMatch(/^\s*[=+\-@]/);
      expect(Number.isNaN(Number(cell.v)), where).toBe(true);
    }
  });

  it("opens on a cell with something in it", () => {
    const first = block.cells[0]?.[0];
    expect(first?.v !== undefined || first?.f !== undefined).toBe(true);
  });

  it("never types a 0 into a header: the number formats show a zero as a dash", () => {
    for (const { cell, row, column } of cellsOf(block.cells)) {
      if (cell.kind !== "header" || typeof cell.v !== "number") continue;
      const where = `${entry.id} r${String(row)}c${String(column)}`;
      expect(cell.v, `${where} would read "-"`).not.toBe(0);
    }
  });

  it("opens with its own name when it opens with a title", () => {
    const first = block.cells[0]?.[0];
    if (first?.kind === "title") expect(first.v).toBe(block.name);
  });

  it("has a help sentence for its Insert button", () => {
    const sentence = HELP[TEMPLATES_HELP_KEY]?.buttons[entry.id];
    expect(sentence, `no help sentence for ${entry.id}`).toBeDefined();
    expect(sentence?.trim()).not.toBe("");
  });
});

describe("the Templates help copy", () => {
  it("names no template the registry no longer lists", () => {
    const ids = new Set(TEMPLATES.map((entry) => entry.id));
    const stale = Object.keys(HELP[TEMPLATES_HELP_KEY]?.buttons ?? {}).filter(
      (key) => !ids.has(key),
    );
    expect(stale).toEqual([]);
  });
});
