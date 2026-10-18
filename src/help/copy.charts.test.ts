// What the Links help says about a chart link, held to the facts the chart reader
// and the slide settle: a blank cell in the sheet is always a gap on the slide
// (never a zero), and a chart is shapes only where PowerPoint can draw them -
// the card never promises more than the host and the shape budget allow. The
// general rules for every card (length, closed sentences) are in copy.test.ts.

import { describe, expect, it } from "vitest";
import { HELP, HELP_MAX_CHARS } from "./copy";

function allHelpFor(button: string): string[] {
  return Object.values(HELP).flatMap((section) =>
    section.buttons[button] === undefined ? [] : [section.buttons[button]],
  );
}

describe("the help for exporting a chart", () => {
  const [text] = allHelpFor("export-chart");

  it("exists once, for the Excel pane's button", () => {
    expect(allHelpFor("export-chart")).toHaveLength(1);
    expect(text).toBeDefined();
  });

  it("says a blank cell is a gap, and a zero in an area chart (as Excel plots it)", () => {
    expect(text).toMatch(/Blanks are gaps \(zero in an area\)/);
  });

  it("promises shapes only where PowerPoint can draw them", () => {
    expect(text).toMatch(/where PowerPoint can draw them/);
    expect(text).not.toMatch(/every chart|any chart|all charts/i);
  });

  it("still says what is taken when no chart is selected", () => {
    expect(text).toMatch(/sheet's only one/);
  });

  it("stays a card's length", () => {
    expect(text!.length).toBeLessThanOrEqual(HELP_MAX_CHARS);
  });
});
