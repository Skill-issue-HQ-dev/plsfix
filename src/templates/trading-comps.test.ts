// The trading-comps block against the finance it names: EV multiples and P / E
// per peer, the six Comps stats rows over the peers whose multiple means
// something, and a target's implied EV off the median multiples. Computed by
// test/formula-eval.ts and held to the same figures worked out independently,
// on the demo peers and on peers typed over them. Typed text and a loss-making
// target are in trading-comps.typed.test.ts.
import { describe, expect, it } from "vitest";
import { evaluateTemplate, type Value } from "../../test/formula-eval";
import { STATS_LABELS } from "../stats";
import { resolveFormula, type Template } from "../template-cells";
import { tradingComps } from "./trading-comps";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// Where things sit with the block's corner at A1: the six peers on rows 3-8,
// the six statistics under them, the target on row 17, the implied table last.
const PEER_ROWS = [3, 4, 5, 6, 7, 8];
const STAT = {
  min: 9,
  p25: 10,
  median: 11,
  mean: 12,
  p75: 13,
  max: 14,
} as const;
const TARGET = 17;
const IMPLIED = {
  evRevenue: 20,
  evEbitda: 21,
  pe: 22,
  low: 23,
  high: 24,
  mid: 25,
};

// The five figures a peer is typed in with, and the three multiples they make.
const FIGURES = ["B", "C", "D", "E", "F"] as const;
const [EV, REVENUE, EBITDA, INCOME, DEBT] = FIGURES;
const MULTIPLES = { evRevenue: "G", evEbitda: "H", pe: "I" } as const;
type Multiple = keyof typeof MULTIPLES;

interface Peer {
  ev: number;
  revenue: number;
  ebitda: number;
  income: number;
  debt: number;
}

const block = tradingComps();
const addr = (column: string, row: number): string => `${column}${String(row)}`;

// A block cell by its A1 address, the corner being A1.
function cellAt(template: Template, address: string) {
  const column = address.charCodeAt(0) - 65;
  return template.cells[Number(address.slice(1)) - 1]?.[column];
}

const literal = (address: string): number =>
  cellAt(block, address)?.v as number;

// The peers exactly as the block types them in.
const demoPeers = (): Peer[] =>
  PEER_ROWS.map((row) => ({
    ev: literal(addr(EV, row)),
    revenue: literal(addr(REVENUE, row)),
    ebitda: literal(addr(EBITDA, row)),
    income: literal(addr(INCOME, row)),
    debt: literal(addr(DEBT, row)),
  }));

// What one peer's multiples should be: a multiple over a loss or a blank is not
// meaningful and is left out, which is what NM means on a comps page; so is a
// P / E whose equity value, EV less net debt, is not positive.
function multiplesOf(peer: Peer): Record<Multiple, number | null> {
  return {
    evRevenue: peer.revenue > 0 ? peer.ev / peer.revenue : null,
    evEbitda: peer.ebitda > 0 ? peer.ev / peer.ebitda : null,
    pe:
      peer.income > 0 && peer.ev > peer.debt
        ? (peer.ev - peer.debt) / peer.income
        : null,
  };
}

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
  const next = sorted[whole + 1] ?? here;
  return here + (rank - whole) * (next - here);
}

// One multiple across the peers that have one.
function across(peers: Peer[], multiple: Multiple): number[] {
  return peers.flatMap((peer) => multiplesOf(peer)[multiple] ?? []);
}

// Overrides that blank every figure of the given peer rows.
function emptied(rows: number[]): Record<string, Value> {
  const blanks: Record<string, Value> = {};
  for (const row of rows)
    for (const column of FIGURES) blanks[addr(column, row)] = null;
  return blanks;
}

describe("the trading comps block", () => {
  it("is a 25 by 9 block named Trading comps, titled by its name", () => {
    expect(block.id).toBe("trading-comps");
    expect(block.name).toBe("Trading comps");
    expect([block.rows, block.cols]).toEqual([25, 9]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Trading comps",
    });
  });

  it("heads the peer table with what each column holds, units in the first", () => {
    expect(block.cells[1]?.map((cell) => cell.v)).toEqual([
      "Company (EUR m)",
      "EV",
      "Revenue",
      "EBITDA",
      "Net income",
      "Net debt",
      "EV / Revenue",
      "EV / EBITDA",
      "P / E",
    ]);
    expect(block.cells[1]?.every((cell) => cell.kind === "header")).toBe(true);
  });

  it("types six distinct peers and their five figures in as inputs", () => {
    for (const row of PEER_ROWS) {
      expect(cellAt(block, addr("A", row))).toMatchObject({ kind: "input" });
      for (const column of FIGURES) {
        expect(cellAt(block, addr(column, row))).toMatchObject({
          kind: "input",
          format: "decimal",
        });
      }
    }
    const names = PEER_ROWS.map((row) => cellAt(block, addr("A", row))?.v);
    expect(new Set(names).size).toBe(6);
  });

  it("computes each peer's three multiples from its own row", () => {
    const sheet = evaluateTemplate(block);
    demoPeers().forEach((peer, index) => {
      const want = multiplesOf(peer);
      for (const [name, column] of Object.entries(MULTIPLES)) {
        expect(
          sheet.value(addr(column, PEER_ROWS[index] as number)),
          `${name} of peer ${String(index + 1)}`,
        ).toBeCloseTo(want[name as Multiple] as number, 10);
      }
    });
  });

  it("labels and takes the six Comps stats rows down each multiple's column", () => {
    const labels = Object.values(STAT).map(
      (row) => cellAt(block, addr("A", row))?.v,
    );
    expect(labels).toEqual([...STATS_LABELS]);

    const sheet = evaluateTemplate(block);
    for (const [name, column] of Object.entries(MULTIPLES)) {
      const values = across(demoPeers(), name as Multiple);
      const want = {
        min: Math.min(...values),
        p25: percentile(values, 0.25),
        median: median(values),
        mean: values.reduce((a, b) => a + b, 0) / values.length,
        p75: percentile(values, 0.75),
        max: Math.max(...values),
      };
      for (const [stat, row] of Object.entries(STAT)) {
        expect(sheet.value(addr(column, row)), `${name} ${stat}`).toBeCloseTo(
          want[stat as keyof typeof want],
          10,
        );
      }
    }
  });

  it("reads like a comps page: multiples in the range a mid-cap trades at", () => {
    const peers = demoPeers();
    const within = (multiple: Multiple, low: number, high: number): void => {
      for (const value of across(peers, multiple)) {
        expect(value, multiple).toBeGreaterThan(low);
        expect(value, multiple).toBeLessThan(high);
      }
    };
    within("evRevenue", 0.5, 5);
    within("evEbitda", 5, 15);
    within("pe", 8, 25);
  });

  it("leaves a loss-making peer's multiples out of the statistics", () => {
    const peers = demoPeers();
    const sheet = evaluateTemplate(block, {
      [addr(EBITDA, 3)]: -5,
      [addr(INCOME, 3)]: -2,
    });
    expect(sheet.value(addr(MULTIPLES.evEbitda, 3))).toBe("");
    expect(sheet.value(addr(MULTIPLES.pe, 3))).toBe("");

    const losses = peers.map((peer, index) =>
      index === 0 ? { ...peer, ebitda: -5, income: -2 } : peer,
    );
    const pe = across(losses, "pe");
    expect(sheet.value(addr(MULTIPLES.evEbitda, STAT.median))).toBeCloseTo(
      median(across(losses, "evEbitda")),
      10,
    );
    expect(sheet.value(addr(MULTIPLES.pe, STAT.mean))).toBeCloseTo(
      pe.reduce((a, b) => a + b, 0) / pe.length,
      10,
    );
    // Its EV / Revenue is still meaningful and stays in.
    expect(sheet.value(addr(MULTIPLES.evRevenue, 3))).toBeCloseTo(
      multiplesOf(peers[0] as Peer).evRevenue as number,
      10,
    );
  });

  it("gives a peer whose net debt exceeds its EV no P / E, and keeps its other multiples", () => {
    const peers = demoPeers();
    const sheet = evaluateTemplate(block, { [addr(DEBT, 3)]: 5000 });
    expect(sheet.value(addr(MULTIPLES.pe, 3))).toBe("");
    expect(sheet.value(addr(MULTIPLES.evEbitda, 3))).toBeCloseTo(
      multiplesOf(peers[0] as Peer).evEbitda as number,
      10,
    );
    const rest = peers.map((peer, index) =>
      index === 0 ? { ...peer, debt: 5000 } : peer,
    );
    expect(sheet.value(addr(MULTIPLES.pe, STAT.median))).toBeCloseTo(
      median(across(rest, "pe")),
      10,
    );
  });

  it("lets a modeller clear a peer row without breaking anything", () => {
    const sheet = evaluateTemplate(block, emptied([8]));
    for (const column of Object.values(MULTIPLES)) {
      expect(sheet.value(addr(column, 8)), column).toBe("");
    }
    expect(sheet.value(addr(MULTIPLES.evEbitda, STAT.median))).toBeCloseTo(
      median(across(demoPeers().slice(0, 5), "evEbitda")),
      10,
    );
  });

  it("leaves every statistic and the implied range empty once no peer is left", () => {
    const sheet = evaluateTemplate(block, emptied(PEER_ROWS));
    for (const row of Object.values(STAT)) {
      for (const column of Object.values(MULTIPLES)) {
        expect(sheet.value(addr(column, row)), `${column}${String(row)}`).toBe(
          "",
        );
      }
    }
    for (const row of Object.values(IMPLIED)) {
      expect(sheet.value(addr("D", row)), `D${String(row)}`).toBe("");
    }
  });

  it("types the target in beside the peers' own columns", () => {
    expect(cellAt(block, addr("A", TARGET))).toMatchObject({ kind: "input" });
    for (const column of [REVENUE, EBITDA, INCOME, DEBT]) {
      expect(cellAt(block, addr(column, TARGET)), column).toMatchObject({
        kind: "input",
        format: "decimal",
      });
    }
    expect(block.cells[TARGET - 2]?.slice(0, 6).map((cell) => cell.v)).toEqual([
      "Target",
      undefined,
      "Revenue",
      "EBITDA",
      "Net income",
      "Net debt",
    ]);
  });

  it("prices the target off the median multiples, P / E through equity plus net debt", () => {
    const peers = demoPeers();
    const target = (column: string): number => literal(addr(column, TARGET));
    const want = {
      evRevenue: median(across(peers, "evRevenue")) * target(REVENUE),
      evEbitda: median(across(peers, "evEbitda")) * target(EBITDA),
      pe: median(across(peers, "pe")) * target(INCOME) + target(DEBT),
    };
    const sheet = evaluateTemplate(block);
    expect(sheet.value(addr("D", IMPLIED.evRevenue))).toBeCloseTo(
      want.evRevenue,
      9,
    );
    expect(sheet.value(addr("D", IMPLIED.evEbitda))).toBeCloseTo(
      want.evEbitda,
      9,
    );
    expect(sheet.value(addr("D", IMPLIED.pe))).toBeCloseTo(want.pe, 9);

    const all = Object.values(want);
    const [low, high] = [Math.min(...all), Math.max(...all)];
    expect(sheet.value(addr("D", IMPLIED.low))).toBeCloseTo(low, 9);
    expect(sheet.value(addr("D", IMPLIED.high))).toBeCloseTo(high, 9);
    expect(sheet.value(addr("D", IMPLIED.mid))).toBeCloseTo(
      (low + high) / 2,
      9,
    );
  });

  it("shows the metric and the median multiple each implied EV came from", () => {
    const sheet = evaluateTemplate(block);
    const peers = demoPeers();
    expect(sheet.value(addr("B", IMPLIED.evEbitda))).toBe(
      literal(addr(EBITDA, TARGET)),
    );
    expect(sheet.value(addr("C", IMPLIED.evEbitda))).toBeCloseTo(
      median(across(peers, "evEbitda")),
      10,
    );
    expect(sheet.value(addr("C", IMPLIED.pe))).toBeCloseTo(
      median(across(peers, "pe")),
      10,
    );
  });

  it("moves the implied range when the target's figures are typed over", () => {
    const sheet = evaluateTemplate(block, {
      [addr(REVENUE, TARGET)]: 1000,
      [addr(INCOME, TARGET)]: 40,
      [addr(DEBT, TARGET)]: 0,
    });
    const peers = demoPeers();
    expect(sheet.value(addr("D", IMPLIED.evRevenue))).toBeCloseTo(
      median(across(peers, "evRevenue")) * 1000,
      9,
    );
    expect(sheet.value(addr("D", IMPLIED.pe))).toBeCloseTo(
      median(across(peers, "pe")) * 40,
      9,
    );
  });

  it("writes its formulas where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const cell = block.cells[row]?.[column];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    // N() reads typed text and a blank as zero and lets a real error through.
    expect(formula(2, 6)).toBe('=IF(AND(N(D7)>0,N(E7)>0),D7/E7,"")');
    expect(formula(2, 7)).toBe('=IF(AND(N(D7)>0,N(F7)>0),D7/F7,"")');
    expect(formula(2, 8)).toBe(
      '=IF(AND(N(D7)>0,N(G7)>0,N(D7)-N(H7)>0),(D7-N(H7))/G7,"")',
    );
    expect(formula(10, 6)).toBe('=IF(COUNT(I7:I12)=0,"",MEDIAN(I7:I12))');
    // A method that means nothing drops out; P / E also needs a readable net debt.
    expect(formula(20, 3)).toBe('=IF(OR(N(D25)<=0,ISTEXT(E25)),"",D25*E25)');
    expect(formula(21, 3)).toBe(
      '=IF(OR(N(D26)<=0,ISTEXT(E26),ISTEXT(H21)),"",D26*E26+H21)',
    );
    expect(formula(22, 3)).toBe('=IF(COUNT(F24:F26)=0,"",MIN(F24:F26))');
    expect(formula(23, 3)).toBe('=IF(COUNT(F24:F26)=0,"",MAX(F24:F26))');
    expect(formula(24, 3)).toBe('=IF(COUNT(F27:F28)=0,"",AVERAGE(F27:F28))');
  });
});
