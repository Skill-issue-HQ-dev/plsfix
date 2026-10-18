// The vocabulary a template grid is written in, beyond the placeholder grammar
// templates.test.ts already pins: the full set of number formats a cell may
// carry, a typed text input and a typed date, whose serial must be the number
// Excel itself stores for that day.
import { describe, expect, it } from "vitest";
import {
  dateInput,
  excelDate,
  inputText,
  TEMPLATE_FORMATS,
} from "./template-cells";

describe("TEMPLATE_FORMATS", () => {
  it("names every number format a template cell may carry", () => {
    expect([...TEMPLATE_FORMATS].sort()).toEqual([
      "currency",
      "date",
      "decimal",
      "decimal2",
      "multiple",
      "percent",
      "whole",
    ]);
  });
});

describe("inputText", () => {
  it("is a typed text cell in the input colour, with no number format", () => {
    expect(inputText("Peer A")).toEqual({ v: "Peer A", kind: "input" });
  });
});

describe("excelDate", () => {
  it("counts days from Excel's own epoch", () => {
    expect(excelDate(1970, 1, 1)).toBe(25_569);
    expect(excelDate(2025, 1, 1)).toBe(45_658);
    expect(excelDate(2025, 11, 14)).toBe(45_975);
  });

  it("keeps Excel's 1900 leap-year day, so March 1900 lines up", () => {
    expect(excelDate(1900, 3, 1)).toBe(61);
  });

  it("rolls a month past 12 into the next year like Date does", () => {
    expect(excelDate(2024, 13, 1)).toBe(excelDate(2025, 1, 1));
  });
});

describe("dateInput", () => {
  it("is a typed number in the date format", () => {
    expect(dateInput(2025, 11, 14)).toEqual({
      v: 45_975,
      kind: "input",
      format: "date",
    });
  });
});
