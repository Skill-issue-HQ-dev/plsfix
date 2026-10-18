// The cap table block against the finance it names: a priced round with the new
// option pool inside the pre-money, so the new money buys exactly its post-money
// stake. Computed by test/formula-eval.ts and held to hand figures, to the defining
// properties of a priced round and to 100 seeded rounds.
import { describe, expect, it } from "vitest";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { resolveFormula } from "../template-cells";
import { capTable } from "./cap-table";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// With the block's corner at A1: the round in B2:B11, the holder table on rows
// 14 to 19 (Founders, Series A, Series B, Option pool, Other holders, New
// investors) with shares before in B, % before in C, new shares in D, shares
// after in E and % after in F, its total on row 20 and the two checks on 22, 23.
const CELL = {
  pre: "B2",
  money: "B3",
  post: "B4",
  stake: "B5",
  pool: "B6",
  before: "B7",
  after: "B8",
  price: "B9",
  issued: "B10",
  topUp: "B11",
  checkTotals: "B22",
  checkStake: "B23",
} as const;
const HOLDER_ROWS = [14, 15, 16, 17, 18, 19];
const POOL_ROW = 17;
const NEW_ROW = 19;
const TOTAL_ROW = 20;
const BEFORE = (row: number): string => `B${String(row)}`;
const BEFORE_PCT = (row: number): string => `C${String(row)}`;
const NEW = (row: number): string => `D${String(row)}`;
const AFTER = (row: number): string => `E${String(row)}`;
const AFTER_PCT = (row: number): string => `F${String(row)}`;

interface Round {
  pre: number;
  money: number;
  pool: number;
  // Founders, Series A, Series B, the existing option pool, other holders.
  shares: number[];
}

const DEMO: Round = {
  pre: 40,
  money: 10,
  pool: 0.1,
  shares: [5.5, 2, 1, 0.5, 1],
};

function overrides(round: Round): Record<string, Value> {
  const values: Record<string, Value> = {
    [CELL.pre]: round.pre,
    [CELL.money]: round.money,
    [CELL.pool]: round.pool,
  };
  round.shares.forEach((shares, index) => {
    values[BEFORE(HOLDER_ROWS[index] as number)] = shares;
  });
  return values;
}

const block = capTable();

function cellAt(address: string) {
  const column = address.charCodeAt(0) - 65;
  return block.cells[Number(address.slice(1)) - 1]?.[column];
}

describe("the cap table block", () => {
  it("is a 23 by 6 block named Cap table, titled by its name", () => {
    expect(block.id).toBe("cap-table");
    expect(block.name).toBe("Cap table");
    expect([block.rows, block.cols]).toEqual([23, 6]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Cap table",
    });
  });

  it("labels the round and the six holders, with units on the numbers", () => {
    expect(block.cells.slice(1, 11).map((row) => row[0]?.v)).toEqual([
      "Pre-money valuation (EUR m)",
      "New money (EUR m)",
      "Post-money valuation (EUR m)",
      "New money as % of post-money",
      "New option pool (% of post-money)",
      "Shares before the round (m)",
      "Shares after the round (m)",
      "Price per share (EUR)",
      "New shares issued (m)",
      "New option pool shares (m)",
    ]);
    expect(block.cells[12]?.map((row) => row.v)).toEqual([
      "Holder",
      "Shares before (m)",
      "% before",
      "New shares (m)",
      "Shares after (m)",
      "% after",
    ]);
    expect(HOLDER_ROWS.map((row) => cellAt(`A${String(row)}`)?.v)).toEqual([
      "Founders",
      "Series A",
      "Series B",
      "Option pool",
      "Other holders",
      "New investors",
    ]);
    expect(cellAt(`A${String(TOTAL_ROW)}`)).toMatchObject({
      kind: "result",
      v: "Total",
    });
  });

  it("types the round and the five existing holdings in as inputs, and derives everything else", () => {
    const typed: [string, string][] = [
      [CELL.pre, "decimal"],
      [CELL.money, "decimal"],
      [CELL.pool, "percent"],
      ...HOLDER_ROWS.slice(0, 5).map((row): [string, string] => [
        BEFORE(row),
        "decimal",
      ]),
    ];
    for (const [address, format] of typed) {
      expect(cellAt(address), address).toMatchObject({ kind: "input", format });
      expect(cellAt(address)?.f, address).toBeUndefined();
    }
    // The new investors hold nothing before the round: no input to type over.
    expect(cellAt(BEFORE(NEW_ROW))?.v).toBeUndefined();
    expect(cellAt(BEFORE(NEW_ROW))?.f).toBeUndefined();
    for (const address of [CELL.post, CELL.after, CELL.price, CELL.issued]) {
      expect(cellAt(address), address).toMatchObject({ kind: "formula" });
    }
  });

  it("shows the price per share to two decimals, the money and the shares to one", () => {
    // by hand: a price of 3.59 would read 3.6 at one decimal; the shares and the
    // EUR millions around it keep one.
    expect(cellAt(CELL.price)).toMatchObject({
      kind: "formula",
      format: "decimal2",
    });
    for (const address of [
      CELL.post,
      CELL.before,
      CELL.after,
      CELL.issued,
      CELL.topUp,
    ]) {
      expect(cellAt(address), address).toMatchObject({ format: "decimal" });
    }
  });

  it("opens on the demo round the hand figures below are worked on", () => {
    expect(cellAt(CELL.pre)?.v).toBe(DEMO.pre);
    expect(cellAt(CELL.money)?.v).toBe(DEMO.money);
    expect(cellAt(CELL.pool)?.v).toBe(DEMO.pool);
    const held = HOLDER_ROWS.slice(0, 5).map((row) => cellAt(BEFORE(row))?.v);
    expect(held).toEqual(DEMO.shares);
    // Typing the demo over itself changes nothing.
    expect(evaluateTemplate(block, overrides(DEMO)).value(CELL.price)).toBe(
      evaluateTemplate(block).value(CELL.price),
    );
  });

  it("prices the demo round: post-money, stake, the new pool and the price per share", () => {
    // by hand (EUR m, shares in millions): post-money 40.0 + 10.0 = 50.0; the new
    // money is 10 / 50 = 20 % of it; the five holders hold 5.5 + 2.0 + 1.0 + 0.5 +
    // 1.0 = 10.0 shares; a 10 % pool and a 20 % stake leave the old holders
    // 70 %, so 10.0 / 0.7 = 14.2857 shares after the round; the price is
    // 50.0 / 14.2857 = 3.50; the new investors get 20 % x 14.2857 = 2.8571 shares
    // and the pool 10 % x 14.2857 = 1.4286.
    const sheet = evaluateTemplate(block);
    expect(sheet.value(CELL.post)).toBe(50);
    expect(sheet.value(CELL.stake)).toBeCloseTo(0.2, 12);
    expect(sheet.value(CELL.before)).toBe(10);
    expect(sheet.value(CELL.after)).toBeCloseTo(10 / 0.7, 12);
    expect(sheet.value(CELL.price)).toBeCloseTo(3.5, 12);
    expect(sheet.value(CELL.issued)).toBeCloseTo(2 / 0.7, 12);
    expect(sheet.value(CELL.topUp)).toBeCloseTo(1 / 0.7, 12);
  });

  it("tables the six holders before and after, the pool topped up and the new money in", () => {
    // by hand: before 5.5, 2.0, 1.0, 0.5, 1.0 and none for the new investors =
    // 55 %, 20 %, 10 %, 5 %, 10 %, 0 %; after 5.5, 2.0, 1.0, 0.5 + 1.4286 =
    // 1.9286, 1.0, 2.8571 = 14.2857 shares = 38.5 %, 14.0 %, 7.0 %, 13.5 %, 7.0 %, 20.0 %.
    const sheet = evaluateTemplate(block);
    const before = [0.55, 0.2, 0.1, 0.05, 0.1, 0];
    const after = [0.385, 0.14, 0.07, 0.135, 0.07, 0.2];
    const shares = [5.5, 2, 1, 0.5 + 1 / 0.7, 1, 2 / 0.7];
    HOLDER_ROWS.forEach((row, index) => {
      const who = `row ${String(row)}`;
      expect(sheet.value(BEFORE_PCT(row)), who).toBeCloseTo(
        before[index] as number,
        12,
      );
      expect(sheet.value(AFTER(row)), who).toBeCloseTo(
        shares[index] as number,
        12,
      );
      expect(sheet.value(AFTER_PCT(row)), who).toBeCloseTo(
        after[index] as number,
        12,
      );
    });
    expect(sheet.value(NEW(POOL_ROW))).toBe(sheet.value(CELL.topUp));
    expect(sheet.value(NEW(NEW_ROW))).toBe(sheet.value(CELL.issued));
  });

  it("totals each column, and both percent columns to 100 %", () => {
    const sheet = evaluateTemplate(block);
    expect(sheet.value(BEFORE(TOTAL_ROW))).toBe(10);
    expect(sheet.value(BEFORE_PCT(TOTAL_ROW))).toBeCloseTo(1, 12);
    expect(sheet.value(AFTER_PCT(TOTAL_ROW))).toBeCloseTo(1, 12);
    expect(sheet.value(AFTER(TOTAL_ROW))).toBeCloseTo(
      sheet.value(CELL.after) as number,
      12,
    );
    expect(sheet.value(NEW(TOTAL_ROW))).toBeCloseTo(3 / 0.7, 12);
  });

  it("puts the pool inside the pre-money: the price pays for the old shares and the new pool, and the new money buys the new shares", () => {
    // by hand: 3.50 x (10.0 + 1.4286) = 40.0 is the pre-money; 3.50 x 2.8571 =
    // 10.0 is the new money. The old holders carry the pool, not the new investors.
    const sheet = evaluateTemplate(block);
    const price = sheet.value(CELL.price) as number;
    const old =
      (sheet.value(CELL.before) as number) +
      (sheet.value(CELL.topUp) as number);
    expect(price * old).toBeCloseTo(40, 10);
    expect(price * (sheet.value(CELL.issued) as number)).toBeCloseTo(10, 10);
  });

  it("dilutes every old holder by the same factor: 70 % of its stake, 100 % less the new money and the new pool", () => {
    // by hand: Founders 55 % x 0.7 = 38.5 %, Series A 20 % x 0.7 = 14 %.
    const sheet = evaluateTemplate(block);
    for (const row of [14, 15, 16, 18]) {
      const was = sheet.value(BEFORE_PCT(row)) as number;
      expect(
        sheet.value(AFTER_PCT(row)) as number,
        `row ${String(row)}`,
      ).toBeCloseTo(was * 0.7, 12);
    }
  });

  it("prices at the pre-money over the shares outstanding when no pool is added", () => {
    // by hand: pool 0 %: 40.0 / 10.0 = 4.00 a share, the 10.0 buys 2.5 shares,
    // and 2.5 of 12.5 after is the 20 % stake.
    const sheet = evaluateTemplate(block, { [CELL.pool]: 0 });
    expect(sheet.value(CELL.price)).toBeCloseTo(4, 12);
    expect(sheet.value(CELL.issued)).toBeCloseTo(2.5, 12);
    expect(sheet.value(CELL.topUp)).toBe(0);
    expect(sheet.value(AFTER_PCT(NEW_ROW))).toBeCloseTo(0.2, 12);
  });

  it("issues nothing to new investors when there is no new money", () => {
    // by hand: money 0, pool 10 %: after = 10.0 / 0.9 = 11.11 shares, price
    // 40.0 / 11.11 = 3.60, the pool takes 10 % of them.
    const sheet = evaluateTemplate(block, { [CELL.money]: 0 });
    expect(sheet.value(CELL.issued)).toBe(0);
    expect(sheet.value(CELL.price)).toBeCloseTo(3.6, 12);
    expect(sheet.value(AFTER_PCT(NEW_ROW))).toBe(0);
  });

  it("reads OK on both checks for the demo and for any consistent round", () => {
    const sheet = evaluateTemplate(block);
    expect(sheet.value(CELL.checkTotals)).toBe("OK");
    expect(sheet.value(CELL.checkStake)).toBe("OK");
  });

  it("says CHECK when a cell the block derived has been typed over", () => {
    // A typed share count in the "after" column no longer adds up to the round:
    // the new investors' stake is not what the new money bought.
    const typedOver = evaluateTemplate(block, { [AFTER(14)]: 6 });
    expect(typedOver.value(CELL.checkStake)).toBe("CHECK");
    // A typed percent leaves the column short of 100 %.
    const short = evaluateTemplate(block, { [AFTER_PCT(14)]: 0.3 });
    expect(short.value(CELL.checkTotals)).toBe("CHECK");
  });

  it("says CHECK, not nonsense, when the new money and the pool take more than all of the company", () => {
    // by hand: 90 of 100 post-money is 90 %, and a 20 % pool on top is 110 %: the
    // old holders would have to own minus 10 %, so the share count turns negative.
    const sheet = evaluateTemplate(block, {
      [CELL.pre]: 10,
      [CELL.money]: 90,
      [CELL.pool]: 0.2,
    });
    expect(sheet.value(CELL.after) as number).toBeLessThan(0);
    expect(sheet.value(CELL.checkStake)).toBe("CHECK");
  });

  it("holds the defining properties of a priced round over 100 seeded rounds", () => {
    // A fixed-seed linear congruential generator, so a failure reproduces.
    let state = 20_261_003;
    const random = (): number => {
      state = (state * 1_664_525 + 1_013_904_223) % 4_294_967_296;
      return state / 4_294_967_296;
    };
    const between = (low: number, high: number): number =>
      low + (high - low) * random();
    for (let run = 0; run < 100; run += 1) {
      const pre = between(5, 500);
      const round: Round = {
        pre,
        money: between(1, 3 * pre),
        pool: between(0, 0.2),
        shares: [0, 1, 2, 3, 4].map(() => between(0.1, 20)),
      };
      const sheet = evaluateTemplate(block, overrides(round));
      const where = `run ${String(run)}`;
      const price = sheet.value(CELL.price) as number;
      const before = sheet.value(CELL.before) as number;
      const after = sheet.value(CELL.after) as number;
      const issued = sheet.value(CELL.issued) as number;
      const topUp = sheet.value(CELL.topUp) as number;
      // The price pays for the old shares and the pool at the pre-money, and the
      // new money buys the new shares at that price.
      expect(price * (before + topUp), where).toBeCloseTo(round.pre, 6);
      expect(price * issued, where).toBeCloseTo(round.money, 6);
      // The pool is its share of everything after, the shares add up.
      expect(topUp / after, where).toBeCloseTo(round.pool, 10);
      expect(before + topUp + issued, where).toBeCloseTo(after, 8);
      expect(sheet.value(AFTER(TOTAL_ROW)) as number, where).toBeCloseTo(
        after,
        8,
      );
      expect(sheet.value(AFTER_PCT(TOTAL_ROW)) as number, where).toBeCloseTo(
        1,
        10,
      );
      expect(sheet.value(CELL.checkTotals), where).toBe("OK");
      expect(sheet.value(CELL.checkStake), where).toBe("OK");
    }
  });

  it("shows the divide-by-zero of an empty table instead of a price", () => {
    const blanks: Record<string, Value> = {};
    for (const row of HOLDER_ROWS.slice(0, 5)) blanks[BEFORE(row)] = null;
    const sheet = evaluateTemplate(block, blanks);
    expect(sheet.value(CELL.price)).toEqual(new ExcelError("#DIV/0!"));
    expect(sheet.value(BEFORE_PCT(14))).toEqual(new ExcelError("#DIV/0!"));
  });

  it("guards nothing with IFERROR, so no error is ever hidden", () => {
    for (const row of block.cells) {
      for (const found of row) expect(found.f ?? "").not.toContain("IFERROR");
    }
  });

  it("writes its formulas where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const found = block.cells[row]?.[column];
      return found?.f === undefined ? undefined : resolveFormula(found.f, C5);
    };
    expect(formula(3, 1)).toBe("=D6+D7");
    expect(formula(4, 1)).toBe("=D7/D8");
    expect(formula(6, 1)).toBe("=D24");
    expect(formula(7, 1)).toBe("=D11/(1-D9-D10)");
    expect(formula(8, 1)).toBe("=D8/D12");
    expect(formula(9, 1)).toBe("=D9*D12");
    expect(formula(10, 1)).toBe("=D10*D12");
    expect(formula(13, 2)).toBe("=D18/D24");
    expect(formula(13, 4)).toBe("=D18+F18");
    expect(formula(13, 5)).toBe("=G18/G24");
    expect(formula(16, 3)).toBe("=D15");
    expect(formula(18, 3)).toBe("=D14");
    expect(formula(19, 1)).toBe("=SUM(D18:D23)");
    expect(formula(21, 1)).toBe(
      '=IF(AND(ABS(E24-1)<0.0001,ABS(H24-1)<0.0001),"OK","CHECK")',
    );
    expect(formula(22, 1)).toBe(
      '=IF(AND(D12>0,ABS(H23-D9)<0.0001),"OK","CHECK")',
    );
  });
});
