// The revenue-build block against the arithmetic it names: volume and price each
// grown by a typed yearly rate, revenue as their product in EUR millions, growth
// and the year 1 to 5 CAGR. Computed by test/formula-eval.ts and held to closed
// forms worked out independently, on the demo case and on cases typed over it.
import { describe, expect, it } from "vitest";
import { cellAt, errorsIn } from "../../test/block-grid";
import { evaluateTemplate, type Value } from "../../test/formula-eval";
import { HELP } from "../help/copy";
import { resolveFormula } from "../template-cells";
import { entryById, TEMPLATES } from "./index";
import { revenueBuild } from "./revenue-build";
import { filterTemplates } from "./search";

// Column index 2, row index 4: the corner the pinned formulas resolve against.
const C5 = { row: 4, column: 2 };

// With the block's corner at A1: years one to five run B to F, the rows follow.
const YEARS = ["B", "C", "D", "E", "F"] as const;
const LATER_YEARS = ["C", "D", "E", "F"] as const;
const ROW = {
  header: 2,
  volume: 3,
  volumeGrowth: 4,
  price: 5,
  priceGrowth: 6,
  revenue: 8,
  revenueGrowth: 9,
  cagr: 11,
} as const;
// The level rows and the format they show: a price in whole cents (25.75 must not
// round to 25.8), a volume in tenths.
const LEVELS: [number, string][] = [
  [ROW.volume, "decimal"],
  [ROW.price, "decimal2"],
];
// A volume in thousand units times a price in EUR is EUR thousand; a thousand
// of those make one EUR million.
const THOUSANDS_PER_MILLION = 1_000;
const ROOT_OF_FOUR_INTERVALS = 1 / 4;

const addr = (column: string, row: number): string => `${column}${String(row)}`;

interface Case {
  volume: number;
  price: number;
  /** Growth into years two to five. */
  volumeGrowth: number[];
  priceGrowth: number[];
}

// The year-one levels and growth rates the block types in.
const DEMO: Case = {
  volume: 1200,
  price: 25,
  volumeGrowth: [0.05, 0.05, 0.04, 0.04],
  priceGrowth: [0.03, 0.03, 0.025, 0.025],
};

// The cases the closed form is held to: the demo and four others typed over it.
const CASES: [string, Case][] = [
  ["the demo case", DEMO],
  [
    "volume up and price down",
    {
      volume: 800,
      price: 40,
      volumeGrowth: [0.12, 0.1, 0.08, 0.06],
      priceGrowth: [-0.03, -0.02, -0.02, -0.01],
    },
  ],
  [
    "a shrinking business",
    {
      volume: 500,
      price: 12.5,
      volumeGrowth: [-0.1, -0.05, -0.05, -0.02],
      priceGrowth: [0.01, 0, 0.01, 0],
    },
  ],
  [
    "no growth at all",
    {
      volume: 1000,
      price: 30,
      volumeGrowth: [0, 0, 0, 0],
      priceGrowth: [0, 0, 0, 0],
    },
  ],
  [
    "a one-year volume jump",
    {
      volume: 100,
      price: 9,
      volumeGrowth: [1, 0, 0, 0],
      priceGrowth: [0, 0, 0, 0],
    },
  ],
];

// A case as the cells a modeller overtypes; a missing rate becomes a blank.
function typed(c: Case): Record<string, Value> {
  const cells: Record<string, Value> = {
    [addr("B", ROW.volume)]: c.volume,
    [addr("B", ROW.price)]: c.price,
  };
  LATER_YEARS.forEach((column, step) => {
    cells[addr(column, ROW.volumeGrowth)] = c.volumeGrowth[step] ?? null;
    cells[addr(column, ROW.priceGrowth)] = c.priceGrowth[step] ?? null;
  });
  return cells;
}

// Revenue in EUR m of year 1 to 5 in closed form: the year-one volume and price
// times every growth factor up to that year, never one year built off the last.
function revenueOf(c: Case, year: number): number {
  let factor = 1;
  for (let step = 0; step < year - 1; step += 1) {
    const volume = 1 + (c.volumeGrowth[step] ?? 0);
    const price = 1 + (c.priceGrowth[step] ?? 0);
    factor *= volume * price;
  }
  return (c.volume * c.price * factor) / THOUSANDS_PER_MILLION;
}

const block = revenueBuild();

describe("the revenue build block", () => {
  it("is an 11 by 6 block named Revenue build, titled by its name", () => {
    expect(block.id).toBe("revenue-build");
    expect(block.name).toBe("Revenue build");
    expect([block.rows, block.cols]).toEqual([11, 6]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Revenue build",
    });
  });

  it("numbers the five years in its header row", () => {
    expect(cellAt(block, "A2")).toMatchObject({ kind: "header", v: "Year" });
    YEARS.forEach((column, index) => {
      expect(cellAt(block, addr(column, ROW.header)), column).toMatchObject({
        kind: "header",
        v: index + 1,
      });
    });
  });

  it("names the unit of every row that has one", () => {
    const name = (row: number): string =>
      String(cellAt(block, addr("A", row))?.v);
    expect(name(ROW.volume)).toMatch(/thousand units/);
    expect(name(ROW.price)).toMatch(/EUR/);
    expect(name(ROW.revenue)).toMatch(/EUR m/);
    for (const row of [ROW.volumeGrowth, ROW.priceGrowth, ROW.revenueGrowth]) {
      expect(name(row), `row ${String(row)}`).toMatch(/%/);
    }
    const names = [ROW.volume, ROW.price, ROW.revenue, ROW.cagr].map(name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("types the year-one levels and every growth rate in, and builds the rest", () => {
    for (const [row, format] of LEVELS) {
      const level = cellAt(block, addr("B", row));
      expect(level, `level on row ${String(row)}`).toMatchObject({
        kind: "input",
        format,
      });
      expect(level?.f).toBeUndefined();
    }
    for (const row of [ROW.volumeGrowth, ROW.priceGrowth]) {
      for (const column of LATER_YEARS) {
        const rate = cellAt(block, addr(column, row));
        expect(rate, addr(column, row)).toMatchObject({
          kind: "input",
          format: "percent",
        });
        expect(rate?.f).toBeUndefined();
      }
    }
    for (const [row, format] of LEVELS) {
      for (const column of LATER_YEARS) {
        expect(
          cellAt(block, addr(column, row)),
          addr(column, row),
        ).toMatchObject({ kind: "formula", format });
      }
    }
    for (const column of YEARS) {
      expect(cellAt(block, addr(column, ROW.revenue))).toMatchObject({
        kind: "result",
        format: "decimal",
      });
    }
    for (const column of LATER_YEARS) {
      expect(cellAt(block, addr(column, ROW.revenueGrowth))).toMatchObject({
        kind: "formula",
        format: "percent",
      });
    }
    expect(cellAt(block, addr("B", ROW.cagr))).toMatchObject({
      kind: "result",
      format: "percent",
    });
  });

  it("leaves year one's growth cells empty: there is no year before it", () => {
    for (const row of [ROW.volumeGrowth, ROW.priceGrowth, ROW.revenueGrowth]) {
      const cell = cellAt(block, addr("B", row));
      expect(cell?.v, `B${String(row)}`).toBeUndefined();
      expect(cell?.f, `B${String(row)}`).toBeUndefined();
    }
  });

  it("grows volume and price from the prior year by the typed rates", () => {
    const sheet = evaluateTemplate(block);
    // by hand: volume 1,200 grows 5, 5, 4 and 4 %: 1,260, then 1,323, then
    // 1,375.92, then 1,430.9568; price 25 grows 3, 3, 2.5 and 2.5 %: 25.75,
    // 26.5225, 27.1855625, 27.8652015625.
    const volumes = [1200, 1260, 1323, 1375.92, 1430.9568];
    const prices = [25, 25.75, 26.5225, 27.1855625, 27.8652015625];
    YEARS.forEach((column, index) => {
      expect(
        sheet.value(addr(column, ROW.volume)),
        `volume ${column}`,
      ).toBeCloseTo(volumes[index] as number, 9);
      expect(
        sheet.value(addr(column, ROW.price)),
        `price ${column}`,
      ).toBeCloseTo(prices[index] as number, 9);
    });
  });

  it("makes revenue volume times price in EUR millions", () => {
    const sheet = evaluateTemplate(block);
    // by hand: 1,200 thousand units x EUR 25 = EUR 30,000 thousand = 30.0 EUR m;
    // year two 1,260 x 25.75 = 32,445 thousand = 32.445; year three 1,323 x
    // 26.5225 = 35,089.2675 thousand; year four 37,405.159155; year five
    // 39,873.899659 thousand.
    expect(sheet.value("B8")).toBe(30);
    expect(sheet.value("C8")).toBeCloseTo(32.445, 9);
    expect(sheet.value("D8")).toBeCloseTo(35.0892675, 9);
    expect(sheet.value("E8")).toBeCloseTo(37.405159155, 9);
    expect(sheet.value("F8")).toBeCloseTo(39.873899659, 9);
  });

  it("shows each year's revenue growth against the year before", () => {
    const sheet = evaluateTemplate(block);
    // by hand: (1.05 x 1.03) - 1 = 8.15 % into years two and three,
    // (1.04 x 1.025) - 1 = 6.6 % into years four and five.
    expect(sheet.value("C9")).toBeCloseTo(0.0815, 12);
    expect(sheet.value("D9")).toBeCloseTo(0.0815, 12);
    expect(sheet.value("E9")).toBeCloseTo(0.066, 12);
    expect(sheet.value("F9")).toBeCloseTo(0.066, 12);
  });

  it("takes the CAGR as the fourth root of last over first, minus one", () => {
    const sheet = evaluateTemplate(block);
    // by hand: the year 1 to 5 ratio is (1.05 x 1.03)^2 x (1.04 x 1.025)^2 =
    // 1.0815^2 x 1.066^2, so its fourth root is sqrt(1.0815 x 1.066) =
    // sqrt(1.152879) = 1.0737220, a CAGR of 7.3722 %.
    expect(sheet.value("B11")).toBeCloseTo(Math.sqrt(1.152879) - 1, 12);
    expect(sheet.value("B11")).toBeCloseTo(0.073722, 7);
  });

  it.each(CASES)("agrees with the closed form on %s", (_name, c) => {
    const sheet = evaluateTemplate(block, typed(c));
    YEARS.forEach((column, index) => {
      expect(
        sheet.value(addr(column, ROW.revenue)),
        `revenue ${column}`,
      ).toBeCloseTo(revenueOf(c, index + 1), 9);
    });
    LATER_YEARS.forEach((column, step) => {
      const grown = revenueOf(c, step + 2) / revenueOf(c, step + 1) - 1;
      expect(
        sheet.value(addr(column, ROW.revenueGrowth)),
        `growth ${column}`,
      ).toBeCloseTo(grown, 9);
    });
    const cagr =
      (revenueOf(c, 5) / revenueOf(c, 1)) ** ROOT_OF_FOUR_INTERVALS - 1;
    expect(sheet.value("B11")).toBeCloseTo(cagr, 9);
  });

  it("holds flat revenue and a zero CAGR when nothing grows", () => {
    const flat: Case = {
      ...DEMO,
      volumeGrowth: [0, 0, 0, 0],
      priceGrowth: [0, 0, 0, 0],
    };
    const sheet = evaluateTemplate(block, typed(flat));
    for (const column of YEARS) {
      expect(sheet.value(addr(column, ROW.revenue))).toBe(30);
    }
    for (const column of LATER_YEARS) {
      expect(sheet.value(addr(column, ROW.revenueGrowth))).toBe(0);
    }
    expect(sheet.value("B11")).toBe(0);
  });

  it("scales every year's revenue with year one and leaves growth and CAGR alone", () => {
    const demo = evaluateTemplate(block);
    const doubled = evaluateTemplate(block, { B3: 2400 });
    for (const column of YEARS) {
      expect(doubled.value(addr(column, ROW.revenue))).toBeCloseTo(
        2 * (demo.value(addr(column, ROW.revenue)) as number),
        9,
      );
    }
    for (const column of LATER_YEARS) {
      expect(doubled.value(addr(column, ROW.revenueGrowth))).toBeCloseTo(
        demo.value(addr(column, ROW.revenueGrowth)) as number,
        12,
      );
    }
    expect(doubled.value("B11")).toBeCloseTo(demo.value("B11") as number, 12);
  });

  it("computes with no error and no circular reference anywhere", () => {
    expect(errorsIn(block)).toEqual([]);
  });

  it("shows n.m. rather than an error when there is no revenue to grow from", () => {
    const none = { B3: 0 };
    expect(errorsIn(block, none)).toEqual([]);
    const sheet = evaluateTemplate(block, none);
    for (const column of LATER_YEARS) {
      expect(sheet.value(addr(column, ROW.revenueGrowth)), column).toBe("n.m.");
    }
    expect(sheet.value("B11")).toBe("n.m.");
  });

  it("shows a CAGR of n.m. when the revenue is gone by the last year", () => {
    // Volume down 100 % in year five: revenue 0, which is -100 % growth, not n.m.
    const sheet = evaluateTemplate(block, { F4: -1 });
    expect(sheet.value("F8")).toBe(0);
    expect(sheet.value("F9")).toBe(-1);
    expect(sheet.value("B11")).toBe("n.m.");
  });

  it("reads a blank growth cell as no growth", () => {
    const blanks: Record<string, Value> = {};
    for (const column of LATER_YEARS) {
      blanks[addr(column, ROW.volumeGrowth)] = null;
      blanks[addr(column, ROW.priceGrowth)] = null;
    }
    const sheet = evaluateTemplate(block, blanks);
    expect(sheet.value("F8")).toBe(30);
    expect(sheet.value("B11")).toBe(0);
  });

  it("writes the unit conversion and the CAGR where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const cell = block.cells[row]?.[column];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    expect(formula(2, 2)).toBe("=D7*(1+E8)");
    expect(formula(4, 2)).toBe("=D9*(1+E10)");
    expect(formula(7, 1)).toBe("=D7*D9/1000");
    expect(formula(8, 2)).toBe('=IF(D12>0,E12/D12-1,"n.m.")');
    expect(formula(10, 1)).toBe(
      '=IF(AND(D12>0,H12>0),(H12/D12)^(1/4)-1,"n.m.")',
    );
  });
});

describe("the revenue build's wording", () => {
  it("names its CAGR by the years it spans, in the description and the help", () => {
    // The rate runs over four intervals from year 1 to year 5, not five years.
    const help = HELP["templates-heading"]?.buttons["template-revenue-build"];
    for (const text of [block.description, help ?? ""]) {
      expect(text).toContain("year 1 to 5 CAGR");
      expect(text).not.toMatch(/five-year/);
    }
  });
});

describe("the revenue build in the picker", () => {
  it("is listed under Operating, carrying this very block", () => {
    const entry = entryById("template-revenue-build");
    expect(entry?.group).toBe("Operating");
    expect(entry?.block).toEqual(block);
  });

  it("is found by the words a modeller types for it", () => {
    for (const query of ["revenue build", "cagr", "volume", "price"]) {
      const ids = filterTemplates(TEMPLATES, query).map((entry) => entry.id);
      expect(ids, query).toContain("template-revenue-build");
    }
  });
});
