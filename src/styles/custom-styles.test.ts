// The style name boxes of both panes share custom-styles.css: the hint in an empty
// box has to fit it at 320 px (about 180 px beside Save style), and the box is the
// part of its row that gives. The UX gate measures the real width (clipped-placeholder);
// this is the fast guard.

import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

const root = process.cwd();
const read = (file: string): string =>
  readFileSync(path.join(root, file), "utf8");

// Room at 320 px is about 180 px; a 12 px hint takes roughly 6 px a character.
const HINT_MAX_CHARS = 24;
const FIELDS: [file: string, id: string][] = [
  ["taskpane.html", "style-name"],
  ["pptpane.html", "shape-style-name"],
];

function placeholderOf(file: string, id: string): string {
  const input = new RegExp(`<input[^>]*\\bid="${id}"[^>]*>`).exec(read(file));
  expect(input, `${file} has no #${id}`).not.toBeNull();
  const hint = /\bplaceholder="([^"]*)"/.exec(input![0]);
  expect(hint, `#${id} has no placeholder`).not.toBeNull();
  return hint![1]!;
}

describe("the style name boxes", () => {
  it.each(FIELDS)(
    "%s: the hint in #%s is short enough for a 320 px pane",
    (file, id) => {
      const hint = placeholderOf(file, id);
      expect(hint.length).toBeLessThanOrEqual(HINT_MAX_CHARS);
      expect(hint.trim()).not.toBe("");
    },
  );

  it("is the same hint in both panes", () => {
    const [excel, ppt] = FIELDS.map(([file, id]) => placeholderOf(file, id));
    expect(excel).toBe("Style name");
    expect(ppt).toBe(excel);
  });

  it("leaves the box to shrink and the Save style button its own width", () => {
    const css = read("src/styles/custom-styles.css");
    expect(css).toMatch(
      /\.style-save-row\s*\{[^}]*grid-template-columns:\s*minmax\(0,\s*1fr\)\s+auto/,
    );
    expect(css).toMatch(/\.style-name-field\s*\{[^}]*min-width:\s*0/);
  });
});
