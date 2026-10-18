// Manual cross-check, not part of `npm run check`: writes every template block as
// a real .xlsx, demo values and typed-over cases, lets LibreOffice recalculate it
// headless and compares each cell with test/formula-eval.ts. Needs soffice (SOFFICE
// names it): `npx tsx test/tools/libreoffice-recalc.ts [block id]`. A mismatch is a finding.

import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { cellAddress } from "../../src/find";
import { resolveFormula, type Template } from "../../src/template-cells";
import { TEMPLATES } from "../../src/templates/index";
import { type Computed, evaluateTemplate, ExcelError } from "../formula-eval";
import { LIBREOFFICE_CASES, type Overrides } from "./libreoffice-cases";
import { type XlsxCell, writeXlsx } from "./xlsx-store";

const CORNER = { row: 0, column: 0 };
const say = (line: string): void => {
  process.stdout.write(`${line}\n`);
};
// Raw CSV (full precision, formulas calculated): comma, quote, UTF-8, first sheet.
const CSV_FILTER =
  "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,false,false,false,0";
const REL_TOLERANCE = 1e-9;
// Functions Excel 2010 added are stored in a file as _xlfn.NAME; read without it
// LibreOffice answers #NAME?.
const XLFN = /\b(PERCENTILE\.INC)\(/g;
// What the evaluator throws for a function it has no rule for: not a mismatch,
// the block is outside what it can compute, and the run says so.
const UNKNOWN = /^formula-eval does not know the function/;

interface Run {
  name: string;
  block: Template;
  overrides: Overrides;
  file: string;
}

function overrideCell(address: string, value: Computed): XlsxCell | null {
  if (value === null) return null;
  if (value instanceof ExcelError) return { address, error: value.code };
  if (typeof value === "number") return { address, number: value };
  return { address, text: String(value) };
}

// The block's cells as the sheet holds them, its corner at A1, overrides typed over.
function cellsOf(block: Template, overrides: Overrides): XlsxCell[] {
  const cells: XlsxCell[] = [];
  block.cells.forEach((row, r) =>
    row.forEach((cell, c) => {
      const address = cellAddress(r, c);
      if (address in overrides) {
        const typed = overrideCell(address, overrides[address] ?? null);
        if (typed !== null) cells.push(typed);
      } else if (cell.f !== undefined) {
        const formula = resolveFormula(cell.f, CORNER).slice(1);
        cells.push({ address, formula: formula.replace(XLFN, "_xlfn.$1(") });
      } else if (typeof cell.v === "number") {
        cells.push({ address, number: cell.v });
      } else if (typeof cell.v === "string") {
        cells.push({ address, text: cell.v });
      }
    }),
  );
  return cells;
}

// A small RFC 4180 reader: a quoted cell may hold commas and doubled quotes.
function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text.charAt(i);
    if (quoted) {
      if (ch === '"' && text.charAt(i + 1) === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') quoted = false;
      else cell += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (ch !== "\r") cell += ch;
  }
  if (cell !== "" || row.length > 0) rows.push([...row, cell]);
  return rows;
}

// LibreOffice prints its invalid-argument error as Err:502; Excel shows #VALUE!.
const excelText = (shown: string): string =>
  shown === "Err:502" ? "#VALUE!" : shown;

// A number as LibreOffice shows it: it gives an IRR result a percent format of
// its own, so 0.2005 can arrive as "20.05%"; the value is the same.
function numberShown(shown: string): number {
  return shown.endsWith("%") ? Number(shown.slice(0, -1)) / 100 : Number(shown);
}

function agrees(mine: Computed, shown: string): boolean {
  const theirs = excelText(shown);
  if (mine instanceof ExcelError) return theirs === mine.code;
  if (typeof mine === "number") {
    const n = numberShown(theirs);
    const slack = REL_TOLERANCE * Math.max(1, Math.abs(mine));
    return theirs !== "" && Number.isFinite(n) && Math.abs(n - mine) <= slack;
  }
  if (mine === null) return theirs === "";
  return theirs === String(mine);
}

// Compares one block's every cell with the grid LibreOffice calculated.
function compare(run: Run, grid: string[][]): { cells: number; bad: number } {
  const sheet = evaluateTemplate(run.block, run.overrides);
  let bad = 0;
  run.block.cells.forEach((row, r) =>
    row.forEach((_cell, c) => {
      const address = cellAddress(r, c);
      const mine = sheet.value(address);
      const theirs = grid[r]?.[c] ?? "";
      if (agrees(mine, theirs)) return;
      bad += 1;
      const seen = JSON.stringify(
        mine instanceof ExcelError ? mine.code : mine,
      );
      say(
        `  MISMATCH ${run.name} ${address}: evaluator ${seen}, LibreOffice ${JSON.stringify(theirs)}`,
      );
    }),
  );
  return { cells: run.block.rows * run.block.cols, bad };
}

function runs(only: string[], dir: string): Run[] {
  const demos = TEMPLATES.map((entry) => ({
    name: `${entry.block.id} demo`,
    block: entry.block,
    overrides: {},
  }));
  const cases = LIBREOFFICE_CASES.flatMap(([id, name, overrides]) => {
    const block = TEMPLATES.find((entry) => entry.block.id === id)?.block;
    if (block === undefined) throw new Error(`no template block ${id}`);
    return [{ name: `${id} ${name}`, block, overrides }];
  });
  const wanted = (run: { block: Template }): boolean =>
    only.length === 0 || only.includes(run.block.id);
  return [...demos, ...cases].filter(wanted).map((run, index) => ({
    ...run,
    file: path.join(dir, `case-${String(index)}.xlsx`),
  }));
}

function main(): void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "plsfix-lo-"));
  const all = runs(process.argv.slice(2), dir);
  for (const run of all) writeXlsx(run.file, cellsOf(run.block, run.overrides));
  const profile = `-env:UserInstallation=file://${path.join(dir, "profile")}`;
  const soffice = process.env["SOFFICE"] ?? "soffice";
  const args = [
    profile,
    "--headless",
    "--convert-to",
    CSV_FILTER,
    "--outdir",
    dir,
  ];
  execFileSync(soffice, [...args, ...all.map((run) => run.file)], {
    stdio: "pipe",
    timeout: 600_000,
  });
  let cells = 0;
  let bad = 0;
  const skipped: string[] = [];
  for (const run of all) {
    const csv = fs.readFileSync(run.file.replace(/\.xlsx$/, ".csv"), "utf8");
    try {
      const result = compare(run, parseCsv(csv));
      cells += result.cells;
      bad += result.bad;
      const verdict = result.bad === 0 ? "ok " : "BAD";
      say(`${verdict} ${run.name}: ${String(result.cells)} cells`);
    } catch (error) {
      if (!(error instanceof Error) || !UNKNOWN.test(error.message))
        throw error;
      skipped.push(run.name);
      say(`skip ${run.name}: ${error.message}`);
    }
  }
  const done = all.length - skipped.length;
  say(
    `compared ${String(cells)} cells in ${String(done)} runs (${String(skipped.length)} skipped), ${String(bad)} mismatches`,
  );
  fs.rmSync(dir, { recursive: true, force: true });
  if (bad > 0) process.exitCode = 1;
}

main();
