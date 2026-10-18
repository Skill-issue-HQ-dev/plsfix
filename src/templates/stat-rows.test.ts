// The six statistic rows the comps blocks share: the product's own Comps stats
// grammar (src/stats.ts) behind a COUNT guard, so a column with no number reads
// empty instead of an error. Held to Comps stats' own formulas, row for row.
import { describe, expect, it } from "vitest";
import { evaluateTemplate, ExcelError } from "../../test/formula-eval";
import { readCompsBlock, STATS_LABELS, statsGrid } from "../stats";
import { build, head, input, label, resolveFormula } from "../template-cells";
import { STAT_ROW, statRows } from "./stat-rows";

const CORNER = { row: 0, column: 0 };
// Six data rows under a header, the statistics in column 6 of a 7-wide block.
const WIDE = { first: 2, last: 7, columns: [6], width: 7 };

function formulaAt(rows: ReturnType<typeof statRows>, r: number, c: number) {
  const cell = rows[r]?.[c];
  return cell?.f === undefined ? undefined : resolveFormula(cell.f, CORNER);
}

// A one-column block: header, six numbers typed 1..6, then the six statistics.
function tiny(overrides: Parameters<typeof evaluateTemplate>[1] = {}) {
  const numbers = [1, 2, 3, 4, 5, 6].map((n) => [
    label(`p${String(n)}`),
    input(n, "multiple"),
  ]);
  const rows = [
    [head("Company"), head("x")],
    ...numbers,
    ...statRows({ first: 1, last: 6, columns: [1], width: 2 }),
  ];
  return evaluateTemplate(build("t", "T", "T", rows), overrides);
}
const STAT_START = 8; // sheet row of the first statistic in the tiny block

describe("the statistic rows", () => {
  it("are the six Comps stats rows in its own order, labelled", () => {
    const rows = statRows(WIDE);
    expect(rows).toHaveLength(6);
    expect(rows.map((row) => row[0]?.v)).toEqual([...STATS_LABELS]);
    expect(rows.every((row) => row.length === 7)).toBe(true);
    expect(rows.every((row) => row[0]?.kind === "label")).toBe(true);
  });

  it("write the same statistic as Comps stats, behind a COUNT guard", () => {
    const grid = [
      ["Company", "Multiple"],
      ...Array.from({ length: 6 }, () => ["peer", 1.5]),
    ];
    const product = readCompsBlock(grid, 1, 5);
    const theirs = statsGrid(product, 2);
    statRows(WIDE).forEach((_row, index) => {
      const ours = formulaAt(statRows(WIDE), index, 6);
      const plain = (theirs[index]?.[1] ?? "").slice(1);
      expect(ours, STATS_LABELS[index]).toBe(`=IF(COUNT(G3:G8)=0,"",${plain})`);
    });
  });

  it("leave every other column of the row empty", () => {
    const rows = statRows(WIDE);
    for (const row of rows) {
      expect(row.slice(1, 6).every((cell) => cell.v === undefined)).toBe(true);
      expect(row.slice(1, 6).every((cell) => cell.f === undefined)).toBe(true);
    }
  });

  it("name where the median and the quartiles sit among the six", () => {
    expect(STATS_LABELS[STAT_ROW.p25]).toBe("25th percentile");
    expect(STATS_LABELS[STAT_ROW.median]).toBe("Median");
    expect(STATS_LABELS[STAT_ROW.p75]).toBe("75th percentile");
  });

  it("compute Excel's min, quartiles, median, mean and max", () => {
    const sheet = tiny();
    const column = (offset: number) =>
      sheet.value(`B${String(STAT_START + offset)}`);
    expect(column(0)).toBe(1);
    expect(column(STAT_ROW.p25)).toBeCloseTo(2.25, 12);
    expect(column(STAT_ROW.median)).toBe(3.5);
    expect(column(3)).toBe(3.5);
    expect(column(STAT_ROW.p75)).toBeCloseTo(4.75, 12);
    expect(column(5)).toBe(6);
  });

  it("skip a typed word, so the statistics stay numeric over the rest", () => {
    const sheet = tiny({ B2: "n.a.", B3: "n.m." });
    // Left: 3, 4, 5, 6.
    expect(sheet.value(`B${String(STAT_START)}`)).toBe(3);
    expect(sheet.value(`B${String(STAT_START + STAT_ROW.median)}`)).toBe(4.5);
  });

  it("read empty, not an error, once no number is left", () => {
    const sheet = tiny({
      B2: "n.a.",
      B3: "",
      B4: null,
      B5: "x",
      B6: "y",
      B7: "z",
    });
    for (let offset = 0; offset < 6; offset += 1) {
      expect(sheet.value(`B${String(STAT_START + offset)}`)).toBe("");
    }
  });

  it("still show a genuine error in a peer's figure", () => {
    const sheet = tiny({ B4: new ExcelError("#N/A") });
    expect(sheet.value(`B${String(STAT_START + STAT_ROW.median)}`)).toEqual(
      new ExcelError("#N/A"),
    );
  });
});
