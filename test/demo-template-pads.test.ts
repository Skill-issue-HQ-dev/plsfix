// The demo workbook's Templates sheet keeps one empty pad per block, sized to it, in the picker's
// order (the PADS table in demo/src/sheets/templates.rs). The picker writes into an EMPTY
// rectangle, so a block that outgrows its pad would be refused in the demo: this holds the pad
// table to the registry, and a new, renamed or resized block fails here until the demo follows.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { TEMPLATES } from "../src/templates/index";

const PAD =
  /Pad \{ id: "([^"]+)", name: "([^"]+)", group: "([^"]+)", rows: (\d+), cols: (\d+),/g;

interface Pad {
  id: string;
  name: string;
  group: string;
  rows: number;
  cols: number;
}

function demoPads(): Pad[] {
  const source = readFileSync("demo/src/sheets/templates.rs", "utf8");
  return [...source.matchAll(PAD)].map(([, id, name, group, rows, cols]) => ({
    id: id ?? "",
    name: name ?? "",
    group: group ?? "",
    rows: Number(rows),
    cols: Number(cols),
  }));
}

describe("demo workbook template pads", () => {
  it("lists every registry block in the picker's order", () => {
    expect(demoPads().map((pad) => pad.id)).toEqual(
      TEMPLATES.map((entry) => entry.block.id),
    );
  });

  it("names and groups each pad as the picker does", () => {
    expect(demoPads().map(({ name, group }) => ({ name, group }))).toEqual(
      TEMPLATES.map((entry) => ({ name: entry.name, group: entry.group })),
    );
  });

  it("sizes each pad to its block, rows by columns", () => {
    expect(
      demoPads().map(({ id, rows, cols }) => ({ id, rows, cols })),
    ).toEqual(
      TEMPLATES.map((entry) => ({
        id: entry.block.id,
        rows: entry.block.rows,
        cols: entry.block.cols,
      })),
    );
  });
});
