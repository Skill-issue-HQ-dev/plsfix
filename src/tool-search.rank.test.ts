// Find a tool's ranking and words, over rows copied from the shipped catalogue
// (their labels, sentences, section headings and kickers as the pane reads
// them): stop words, the alias table, a row matching its section's heading and
// kicker, the two tie-breaks, the clip at a word and the no-match line. The
// older cases are in tool-search.test.ts; the live catalogue is held to the same
// answers in src/pane/tool-search.rank.test.ts.
import { describe, expect, it } from "vitest";
import {
  clipAtWord,
  noToolLine,
  rankTools,
  type ToolEntry,
} from "./tool-search";

function row(
  action: string,
  label: string,
  sentence: string,
  heading = "",
  kicker = "",
): ToolEntry {
  return { action, label, tab: "Tools", sentence, heading, kicker };
}

// Page order, as the pane catalogues it: the templates sit above the cycles.
const CATALOGUE: ToolEntry[] = [
  row(
    "template-trading-comps",
    "Insert Trading comps",
    "Writes a trading comps table: six peers in, three multiples, six statistics and an implied EV range out. Comps stats is for your own table.",
    "Templates",
    "Start from a block",
  ),
  row(
    "template-dcf",
    "Insert DCF valuation",
    "Writes a DCF: free cash flow, discounting and an exit multiple.",
    "Templates",
    "Start from a block",
  ),
  row(
    "style-save",
    "Save style",
    "Saves the active cell's look under the name in the box. Enter in the box does the same.",
    "Custom styles",
    "Your styles",
  ),
  row(
    "styles-starter",
    "Starter styles from palette",
    "Adds eight starter styles in your brand colors, or restyles them after the palette changed.",
    "Custom styles",
    "Your styles",
  ),
  row(
    "cycle-number-multiple",
    "0.0x",
    "Steps the multiple format: 1.5x, then 1.50x.",
    "Format cycles",
    "Repeat to step",
  ),
  row(
    "undo",
    "Undo last pls,fix action",
    "Puts back what your last five pls,fix actions overwrote, newest first. Excel's own Ctrl+Z never sees add-in writes.",
    "Model tools",
    "Speed",
  ),
  row(
    "trace-precedents",
    "Precedents",
    "Lists the cells the active cell reads from. Click an address to jump there and keep tracing.",
    "Audit",
    "Check and trace",
  ),
  row(
    "trace-dependents",
    "Dependents",
    "Lists the cells that read from the active cell, the same way round.",
    "Audit",
    "Check and trace",
  ),
  row(
    "trace-precedents-all",
    "Precedents of selection",
    "Lists what every formula cell in the selection reads from, grouped by cell.",
    "Audit",
    "Check and trace",
  ),
  row(
    "styles-scan",
    "Scan styles",
    "Lists the custom cell styles no cell in this workbook uses.",
    "Styles",
    "Cleanup",
  ),
  row(
    "export-brand",
    "Copy palette JSON",
    "Copies the palette as JSON, for another workbook or a colleague.",
    "Output defaults",
    "Workbook settings",
  ),
  row(
    "export-selection",
    "Export selection",
    "Sends the selected range as a native table that refreshes in place.",
    "Export",
    "Send to PowerPoint",
  ),
  row(
    "export-chart",
    "Export active chart",
    "Sends the selected chart, or the sheet's only one: shapes where PowerPoint can draw them.",
    "Export",
    "Send to PowerPoint",
  ),
  row(
    "export-text",
    "Export as text",
    "Sends one cell's text as a text box that keeps its place when it refreshes.",
    "Export",
    "Send to PowerPoint",
  ),
  row(
    "export-picture",
    "Export as picture",
    "Sends the selected range as a linked picture, for a layout a table cannot keep.",
    "Export",
    "Send to PowerPoint",
  ),
  row(
    "push-all",
    "Push all",
    "Renders every link in the shown project, or every link if All projects is selected.",
    "Linked objects",
    "This workbook",
  ),
  row(
    "tornado",
    "Tornado from selection",
    "Draws a tornado chart from a low, base and high table.",
    "Charts",
    "Picture the model",
  ),
];

const actions = (query: string, limit?: number): string[] =>
  rankTools(query, CATALOGUE, limit).map((tool) => tool.action);

describe("stop words", () => {
  it.each([
    "a",
    "an",
    "the",
    "to",
    "of",
    "for",
    "my",
    "in",
    "on",
    "and",
    "with",
  ])("%s is left out of the query", (word) => {
    expect(actions(`${word} undo`)).toEqual(actions("undo"));
    expect(actions(`undo ${word}`)).toEqual(actions("undo"));
  });

  it("drops them from the middle of a longer query too", () => {
    expect(actions("trace the precedents")).toEqual(
      actions("trace precedents"),
    );
    expect(actions("send to powerpoint")).toEqual(actions("send powerpoint"));
  });

  it("keeps a query that is nothing but stop words, so a word typed halfway still finds its tool", () => {
    // "to" on its way to "tornado": it must not go quiet until the third letter.
    expect(actions("to")).toContain("tornado");
    expect(actions("to").length).toBeGreaterThan(0);
  });
});

describe("the alias table", () => {
  it('"multiple" ranks the 0.0x cycle first, above the templates that only say multiples', () => {
    expect(actions("multiple")[0]).toBe("cycle-number-multiple");
    // The templates still match: a sentence says it, and they come after.
    expect(actions("multiple")).toContain("template-trading-comps");
    expect(actions("multiple")).toContain("template-dcf");
  });

  it('"send to powerpoint" finds the four Links export buttons first', () => {
    expect(actions("send to powerpoint").slice(0, 4)).toEqual([
      "export-selection",
      "export-chart",
      "export-text",
      "export-picture",
    ]);
  });

  it.each(["send", "update", "powerpoint"])(
    '"%s" reaches the buttons named export, copy and push',
    (word) => {
      const found = actions(word);
      expect(found).toContain("export-selection");
      expect(found).toContain("export-picture");
      expect(found).toContain("push-all");
      expect(found).toContain("export-brand");
    },
  );

  it('"trace precedents" finds Precedents, and not Dependents', () => {
    const found = actions("trace precedents");
    expect(found[0]).toBe("trace-precedents");
    expect(found).toContain("trace-precedents-all");
    expect(found).not.toContain("trace-dependents");
  });

  it('"trace" alone finds both ends of a trace', () => {
    const bare = [
      row("a", "Precedents", "Lists the cells the active cell reads from."),
      row("b", "Dependents", "Lists the cells that read from the active cell."),
      row("c", "Waterfall", "Builds a bridge chart."),
    ];
    expect(rankTools("trace", bare).map((tool) => tool.action)).toEqual([
      "a",
      "b",
    ]);
  });

  it('"undo" still finds the undo button first', () => {
    expect(actions("undo")[0]).toBe("undo");
  });

  it("counts for less than the word itself: a label that says multiple wins", () => {
    const rows = [
      row("cycle", "0.0x", "Steps the number format."),
      row("exact", "Multiple expansion", "Walks a multiple up or down."),
    ];
    expect(rankTools("multiple", rows).map((tool) => tool.action)).toEqual([
      "exact",
      "cycle",
    ]);
  });

  it("reaches a label, never a sentence: a row that only says export is not a send", () => {
    const rows = [row("x", "Other tool", "Does an export of the sheet.")];
    expect(rankTools("send", rows)).toEqual([]);
  });
});

describe("a row's section", () => {
  it("matches the section's heading and its kicker as well as its own words", () => {
    expect(actions("custom styles")).toContain("style-save");
    expect(actions("your styles")).toContain("style-save");
    expect(actions("cleanup")).toEqual(["styles-scan"]);
  });

  it('"custom styles" puts the Custom styles section\'s Save style before Scan styles', () => {
    const found = actions("custom styles");
    expect(found).toContain("style-save");
    expect(found).toContain("styles-scan");
    expect(found.indexOf("style-save")).toBeLessThan(
      found.indexOf("styles-scan"),
    );
  });

  it("starts a word of the heading, never the middle of one", () => {
    expect(actions("ustom")).not.toContain("style-save");
    expect(actions("cust")).toContain("style-save");
  });

  it("still ranks a row that has no section at all", () => {
    const bare: ToolEntry = {
      action: "bare",
      label: "Bare tool",
      tab: "Tools",
      sentence: "Does one thing.",
    };
    expect(rankTools("bare", [bare])).toEqual([bare]);
  });
});

describe("equal scores", () => {
  it("put a row whose label hits above one that only the section hits", () => {
    const rows = [
      row("section", "Plain tool", "Mentions beta.", "Alpha group"),
      row("label", "Alpha tool", "Mentions beta."),
    ];
    expect(rankTools("alpha beta", rows).map((tool) => tool.action)).toEqual([
      "label",
      "section",
    ]);
  });

  it("put a template Insert row after every other row, wherever the page lists it", () => {
    const rows = [
      row("template-dcf", "Insert DCF valuation", "Writes a DCF."),
      row("insert-color-key", "Insert color key", "Drops the color legend."),
      row("insert-toc", "Insert contents sheet", "Adds a contents sheet."),
    ];
    expect(rankTools("insert", rows).map((tool) => tool.action)).toEqual([
      "insert-color-key",
      "insert-toc",
      "template-dcf",
    ]);
  });

  it("keep a template that scores higher ahead of the rest", () => {
    const rows = [
      row("other", "Other tool", "Mentions dcf once."),
      row("template-dcf", "Insert DCF valuation", "Writes a DCF."),
    ];
    expect(rankTools("dcf", rows).map((tool) => tool.action)).toEqual([
      "template-dcf",
      "other",
    ]);
  });

  it("otherwise keep the order the catalogue gave", () => {
    const rows = [
      row("b", "Beta tool", "About beta."),
      row("a", "Alpha tool", "About alpha."),
    ];
    expect(rankTools("tool", rows).map((tool) => tool.action)).toEqual([
      "b",
      "a",
    ]);
  });
});

describe("clipAtWord", () => {
  it("leaves a sentence that fits alone", () => {
    expect(clipAtWord("Steps the multiple format.", 90)).toBe(
      "Steps the multiple format.",
    );
    expect(clipAtWord("x".repeat(90), 90)).toBe("x".repeat(90));
  });

  it("cuts at the last whole word and says so with an ellipsis", () => {
    const sentence =
      "Writes a trading comps table: six peers in, three multiples, six statistics and an implied EV range out. Comps stats is for your own table.";
    const clipped = clipAtWord(sentence, 90);
    expect(clipped).toBe(
      "Writes a trading comps table: six peers in, three multiples, six statistics and an…",
    );
    expect(clipped.length).toBeLessThanOrEqual(90);
  });

  it("never ends mid-word, wherever the limit falls", () => {
    const sentence =
      "Sends the selected range as a native table that refreshes in place. One unbroken block, up to 60 rows and 20 columns.";
    for (let max = 12; max < sentence.length; max += 1) {
      const clipped = clipAtWord(sentence, max);
      expect(clipped.length, `max ${String(max)}`).toBeLessThanOrEqual(max);
      const kept = clipped.slice(0, -1);
      expect(clipped.endsWith("…"), `max ${String(max)}`).toBe(true);
      expect(sentence.startsWith(kept), `max ${String(max)}`).toBe(true);
      // The next character of the sentence is a space: a whole word was kept.
      expect(sentence[kept.length], `max ${String(max)}`).toMatch(/[\s,.;:]/);
    }
  });

  it("keeps a word that ends exactly where the room does", () => {
    expect(clipAtWord("one two three four", 10)).toBe("one two…");
    expect(clipAtWord("one two three four", 14)).toBe("one two three…");
  });

  it("drops a comma or colon the cut leaves hanging", () => {
    expect(clipAtWord("Steps the format, then the next one", 18)).toBe(
      "Steps the format…",
    );
  });

  it("cuts one unbroken word hard, because there is no boundary to use", () => {
    expect(clipAtWord("x".repeat(100), 20)).toBe(`${"x".repeat(19)}…`);
  });
});

describe("noToolLine", () => {
  it("names the query and offers four words that find something", () => {
    expect(noToolLine("zzz")).toBe(
      'No tool matches "zzz". Try export, undo, color or dcf.',
    );
  });

  it("shows the query as typed, trimmed", () => {
    expect(noToolLine("  Frobnicate  ")).toBe(
      'No tool matches "Frobnicate". Try export, undo, color or dcf.',
    );
  });

  it("offers words the ranking really answers", () => {
    for (const word of ["export", "undo", "color", "dcf"]) {
      expect(actions(word).length, word).toBeGreaterThan(0);
    }
  });
});
