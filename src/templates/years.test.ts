// The year strip the multi-year blocks share: year n sits in column n, the
// labels in column 0, so a formula names the year before as column n - 1; a
// block with an opening column gives it one lead column and its years shift by it.
import { describe, expect, it } from "vitest";
import { yearColumns, yearHeader } from "./years";

describe("yearColumns", () => {
  it("puts year n in column n, counting from one", () => {
    expect(yearColumns(5)).toEqual([1, 2, 3, 4, 5]);
    expect(yearColumns(3)).toEqual([1, 2, 3]);
  });

  it("is empty for no years", () => {
    expect(yearColumns(0)).toEqual([]);
  });

  it("shifts every year right by the lead columns before year one", () => {
    expect(yearColumns(3, 1)).toEqual([2, 3, 4]);
    expect(yearColumns(3, 2)).toEqual([3, 4, 5]);
  });
});

describe("yearHeader", () => {
  it("heads the label column Year and numbers each year after it", () => {
    const cells = yearHeader(3);
    expect(cells).toHaveLength(4);
    expect(cells[0]).toEqual({ v: "Year", kind: "header" });
    expect(cells.slice(1)).toEqual([
      { v: 1, kind: "header", format: "whole" },
      { v: 2, kind: "header", format: "whole" },
      { v: 3, kind: "header", format: "whole" },
    ]);
  });

  it("puts the lead columns' words, as text, between Year and year one", () => {
    const cells = yearHeader(3, ["Opening"]);
    expect(cells).toHaveLength(5);
    expect(cells[0]).toEqual({ v: "Year", kind: "header" });
    expect(cells[1]).toEqual({ v: "Opening", kind: "header" });
    expect(cells.slice(2).map((cell) => cell.v)).toEqual([1, 2, 3]);
  });
});
