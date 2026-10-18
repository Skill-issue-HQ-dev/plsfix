// The precedent-transactions block against the finance it names: EV / EBITDA
// per deal, the six Comps stats rows over the deals whose multiple means
// something, and the target's implied EV at the 25th percentile, 75th percentile
// and median multiple. Computed by test/formula-eval.ts and held to the same
// figures worked out independently, on the demo deals and on deals typed over
// them, a typed word and a genuine error included.
import { describe, expect, it } from "vitest";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { STATS_LABELS } from "../stats";
import { excelDate, resolveFormula, type Template } from "../template-cells";
import { precedents } from "./precedents";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// Where things sit with the block's corner at A1: the six deals on rows 3-8,
// their six statistics under them (column F), the target's EBITDA in E16 and
// the implied EVs in D17 to D19.
const DEAL_ROWS = [3, 4, 5, 6, 7, 8];
const STAT = {
  min: 9,
  p25: 10,
  median: 11,
  mean: 12,
  p75: 13,
  max: 14,
} as const;
const TARGET_EBITDA = "E16";
const IMPLIED = { p25: "D17", p75: "D18", median: "D19" } as const;

const block = precedents();

// A block cell by its A1 address, the corner being A1.
function cellAt(template: Template, address: string) {
  const column = address.charCodeAt(0) - 65;
  return template.cells[Number(address.slice(1)) - 1]?.[column];
}

const literal = (address: string): number =>
  cellAt(block, address)?.v as number;

interface Deal {
  ev: number;
  ebitda: number;
}

const demoDeals = (): Deal[] =>
  DEAL_ROWS.map((row) => ({
    ev: literal(`D${String(row)}`),
    ebitda: literal(`E${String(row)}`),
  }));

// A multiple over a zero or negative EBITDA is not meaningful and is left out.
const multiples = (deals: Deal[]): number[] =>
  deals.flatMap((deal) => (deal.ebitda > 0 ? [deal.ev / deal.ebitda] : []));

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 1
    ? (sorted[middle] as number)
    : ((sorted[middle - 1] as number) + (sorted[middle] as number)) / 2;
}

// PERCENTILE.INC worked out independently: the rank k * (n - 1) along the
// sorted values, the two nearest interpolated.
function percentile(values: number[], k: number): number {
  const sorted = [...values].sort((a, b) => a - b);
  const rank = k * (sorted.length - 1);
  const whole = Math.floor(rank);
  const here = sorted[whole] as number;
  return here + (rank - whole) * ((sorted[whole + 1] ?? here) - here);
}

function statistics(values: number[]) {
  return {
    min: Math.min(...values),
    p25: percentile(values, 0.25),
    median: median(values),
    mean: values.reduce((a, b) => a + b, 0) / values.length,
    p75: percentile(values, 0.75),
    max: Math.max(...values),
  };
}

describe("the precedent transactions block", () => {
  it("is a 19 by 6 block named Precedent transactions, titled by its name", () => {
    expect(block.id).toBe("precedents");
    expect(block.name).toBe("Precedent transactions");
    expect([block.rows, block.cols]).toEqual([19, 6]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Precedent transactions",
    });
  });

  it("heads the deal table with what each column holds, units on the money", () => {
    expect(block.cells[1]?.map((cell) => cell.v)).toEqual([
      "Date",
      "Target",
      "Acquirer",
      "EV (EUR m)",
      "EBITDA (EUR m)",
      "EV / EBITDA",
    ]);
    expect(block.cells[1]?.every((cell) => cell.kind === "header")).toBe(true);
  });

  it("types six dated deals in as inputs, newest first, dates as Excel serials", () => {
    const dates = DEAL_ROWS.map((row) => cellAt(block, `A${String(row)}`));
    for (const date of dates) {
      expect(date).toMatchObject({ kind: "input", format: "date" });
    }
    const serials = dates.map((date) => date?.v as number);
    expect(serials).toEqual([...serials].sort((a, b) => b - a));
    expect(new Set(serials).size).toBe(6);
    expect(Math.min(...serials)).toBeGreaterThan(excelDate(2020, 1, 1));
    expect(Math.max(...serials)).toBeLessThan(excelDate(2027, 1, 1));
  });

  it("names a distinct target and acquirer on every deal", () => {
    for (const column of ["B", "C"]) {
      const names = DEAL_ROWS.map(
        (row) => cellAt(block, `${column}${String(row)}`)?.v,
      );
      expect(new Set(names).size).toBe(6);
      for (const row of DEAL_ROWS) {
        expect(cellAt(block, `${column}${String(row)}`)).toMatchObject({
          kind: "input",
        });
      }
    }
    for (const column of ["D", "E"]) {
      for (const row of DEAL_ROWS) {
        expect(cellAt(block, `${column}${String(row)}`)).toMatchObject({
          kind: "input",
          format: "decimal",
        });
      }
    }
  });

  it("computes each deal's EV / EBITDA from its own row", () => {
    const sheet = evaluateTemplate(block);
    demoDeals().forEach((deal, index) => {
      expect(
        sheet.value(`F${String(DEAL_ROWS[index])}`),
        `deal ${String(index + 1)}`,
      ).toBeCloseTo(deal.ev / deal.ebitda, 10);
    });
  });

  it("labels and takes the six Comps stats rows down the multiple column", () => {
    const labels = Object.values(STAT).map(
      (row) => cellAt(block, `A${String(row)}`)?.v,
    );
    expect(labels).toEqual([...STATS_LABELS]);

    const sheet = evaluateTemplate(block);
    const want = statistics(multiples(demoDeals()));
    for (const [stat, row] of Object.entries(STAT)) {
      expect(sheet.value(`F${String(row)}`), stat).toBeCloseTo(
        want[stat as keyof typeof want],
        10,
      );
    }
  });

  it("reads like a precedents page: multiples in the range deals clear at", () => {
    for (const value of multiples(demoDeals())) {
      expect(value).toBeGreaterThan(6);
      expect(value).toBeLessThan(14);
    }
  });

  it("prices the target's EBITDA at the 25th percentile, 75th and median multiple", () => {
    const sheet = evaluateTemplate(block);
    const want = statistics(multiples(demoDeals()));
    const ebitda = literal(TARGET_EBITDA);
    expect(sheet.value(IMPLIED.p25)).toBeCloseTo(ebitda * want.p25, 9);
    expect(sheet.value(IMPLIED.p75)).toBeCloseTo(ebitda * want.p75, 9);
    expect(sheet.value(IMPLIED.median)).toBeCloseTo(ebitda * want.median, 9);
  });

  // Worked out apart from the block: 9.4318x and 10.2751x at the quartiles on a
  // target EBITDA of 52 give 490.45 and 534.31; the median line is 498.43.
  it("lands on the figures a banker checks by hand", () => {
    const sheet = evaluateTemplate(block);
    expect(sheet.value("F10")).toBeCloseTo(9.4318, 4);
    expect(sheet.value("F13")).toBeCloseTo(10.2751, 4);
    expect(sheet.value(IMPLIED.p25)).toBeCloseTo(490.45, 2);
    expect(sheet.value(IMPLIED.median)).toBeCloseTo(498.43, 2);
    expect(sheet.value(IMPLIED.p75)).toBeCloseTo(534.31, 2);
  });

  it("types the target's EBITDA in as an input and labels each implied line", () => {
    expect(cellAt(block, TARGET_EBITDA)).toMatchObject({
      kind: "input",
      format: "decimal",
    });
    expect(
      ["A16", "A17", "A18", "A19"].map((address) => cellAt(block, address)?.v),
    ).toEqual([
      "Target EBITDA",
      "Implied EV at the 25th percentile multiple",
      "Implied EV at the 75th percentile multiple",
      "Implied EV at the median multiple",
    ]);
    expect(cellAt(block, IMPLIED.median)).toMatchObject({ kind: "result" });
  });

  it("moves the implied EVs when the target's EBITDA is typed over", () => {
    const sheet = evaluateTemplate(block, { [TARGET_EBITDA]: 100 });
    const want = statistics(multiples(demoDeals()));
    expect(sheet.value(IMPLIED.median)).toBeCloseTo(100 * want.median, 9);
  });

  it("leaves a deal with no positive EBITDA out of the statistics", () => {
    const deals = demoDeals();
    const sheet = evaluateTemplate(block, { E3: -4 });
    expect(sheet.value("F3")).toBe("");
    const rest = statistics(multiples(deals.slice(1)));
    expect(sheet.value("F11")).toBeCloseTo(rest.median, 10);
    expect(sheet.value("F12")).toBeCloseTo(rest.mean, 10);
  });

  it("leaves a typed word in the EBITDA out and keeps the statistics numeric", () => {
    const sheet = evaluateTemplate(block, { E3: "n.a." });
    expect(sheet.value("F3")).toBe("");
    const rest = statistics(multiples(demoDeals().slice(1)));
    expect(sheet.value("F11")).toBeCloseTo(rest.median, 10);
    expect(sheet.value("F13")).toBeCloseTo(rest.p75, 10);
    expect(typeof sheet.value(IMPLIED.median)).toBe("number");
  });

  it("leaves a deal out when its EV is a word, and shows a genuine error", () => {
    expect(evaluateTemplate(block, { D3: "n.a." }).value("F3")).toBe("");
    const broken = evaluateTemplate(block, { D3: new ExcelError("#N/A") });
    expect(broken.value("F3")).toEqual(new ExcelError("#N/A"));
    expect(broken.value("F11")).toEqual(new ExcelError("#N/A"));
    expect(broken.value(IMPLIED.median)).toEqual(new ExcelError("#N/A"));
  });

  it("lets a modeller clear a deal row without breaking anything", () => {
    const blanks: Record<string, Value> = {};
    for (const column of ["D", "E"]) blanks[`${column}8`] = null;
    const sheet = evaluateTemplate(block, blanks);
    expect(sheet.value("F8")).toBe("");
    expect(sheet.value("F14")).toBeCloseTo(
      statistics(multiples(demoDeals().slice(0, 5))).max,
      10,
    );
  });

  it("leaves the statistics and the implied EVs empty once no deal is left", () => {
    const blanks: Record<string, Value> = {};
    for (const row of DEAL_ROWS) blanks[`E${String(row)}`] = null;
    const sheet = evaluateTemplate(block, blanks);
    for (const row of Object.values(STAT)) {
      expect(sheet.value(`F${String(row)}`), `F${String(row)}`).toBe("");
    }
    for (const address of Object.values(IMPLIED)) {
      expect(sheet.value(address), address).toBe("");
    }
  });

  it("drops every implied line when the target's EBITDA means nothing", () => {
    for (const typed of ["n.a.", 0, -3, null]) {
      const sheet = evaluateTemplate(block, { [TARGET_EBITDA]: typed });
      for (const address of Object.values(IMPLIED)) {
        expect(sheet.value(address), `${address} on ${String(typed)}`).toBe("");
      }
    }
  });

  it("writes its formulas where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const cell = block.cells[row]?.[column];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    // N() reads a word or a blank as zero and lets a genuine error through.
    expect(formula(2, 5)).toBe('=IF(AND(N(F7)>0,N(G7)>0),F7/G7,"")');
    expect(formula(10, 5)).toBe('=IF(COUNT(H7:H12)=0,"",MEDIAN(H7:H12))');
    // A line whose metric or multiple means nothing drops out as "".
    expect(formula(16, 3)).toBe('=IF(OR(N(G20)<=0,ISTEXT(H14)),"",G20*H14)');
    expect(formula(17, 3)).toBe('=IF(OR(N(G20)<=0,ISTEXT(H17)),"",G20*H17)');
    expect(formula(18, 3)).toBe('=IF(OR(N(G20)<=0,ISTEXT(H15)),"",G20*H15)');
  });
});
