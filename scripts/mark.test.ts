// The mark (public/assets/icon.svg) is drawn by hand in six more places: both
// pane headers, the three public pages and the shortcuts-page generator. This
// holds every copy to the icon's geometry and the icon's two fills to the
// stylesheet's lime and ink, so a re-draw in one file cannot drift from the
// rest. Invariant: a lime tile and two ink paths ("p" with its counter, then
// the comma), no <text>: the mark never depends on a font. The video's own copy
// (video/comp/js/ui-mark.js) is the old mark until the videos are re-captured.

import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

const read = (file: string): string =>
  readFileSync(path.join(process.cwd(), file), "utf8");

const attr = (tag: string, name: string): string =>
  new RegExp(`\\b${name}="([^"]+)"`).exec(tag)?.[1] ?? "";
const tags = (source: string, name: string): string[] =>
  [...source.matchAll(new RegExp(`<${name}\\b[^>]*>`, "g"))].map((m) => m[0]);

const icon = read("public/assets/icon.svg");
const tile = tags(icon, "rect")[0] ?? "";
const [letter = "", comma = ""] = tags(icon, "path");

// Every file that carries the mark inline. The generator holds it as a template
// string, which public/shortcuts.html (kept equal by shortcuts:check) repeats.
const COPIES = [
  "taskpane.html",
  "pptpane.html",
  "public/support.html",
  "public/privacy.html",
  "public/shortcuts.html",
  "scripts/build-shortcuts-page.ts",
];

function tokenValue(name: string): string {
  const css = read("src/styles/tokens.css");
  return new RegExp(`${name}:\\s*(#[0-9a-f]{6})`, "i").exec(css)?.[1] ?? "";
}

const subpaths = (d: string): number => (d.match(/M/g) ?? []).length;

describe("the mark", () => {
  it("is a rounded lime tile and two ink paths, never text", () => {
    expect(attr(tile, "rx")).toBe("14");
    expect(tags(icon, "path")).toHaveLength(2);
    expect(icon).not.toMatch(/<text\b|<circle\b|<use\b|font/);
  });

  it("draws the p with an open counter and the comma as one shape each", () => {
    expect(attr(letter, "fill-rule")).toBe("evenodd");
    expect(subpaths(attr(letter, "d"))).toBe(2);
    expect(subpaths(attr(comma, "d"))).toBe(2);
  });

  it("paints the tile in the stylesheet's lime and the glyph in its ink", () => {
    const lime = tokenValue("--mark").toLowerCase();
    const ink = tokenValue("--ink").toLowerCase();
    expect(lime).not.toBe("");
    expect(ink).not.toBe("");
    expect(attr(tile, "fill").toLowerCase()).toBe(lime);
    expect(attr(letter, "fill").toLowerCase()).toBe(ink);
    expect(attr(comma, "fill").toLowerCase()).toBe(ink);
  });

  it.each(COPIES)("is drawn the same in %s", (file) => {
    const source = read(file);
    expect(source).toContain(`d="${attr(letter, "d")}"`);
    expect(source).toContain(`d="${attr(comma, "d")}"`);
    expect(tags(source, "rect")).toContainEqual(
      expect.stringContaining('rx="14"'),
    );
  });

  it("leaves no trace of the old comma-only mark in any copy", () => {
    for (const file of COPIES) {
      expect(read(file), file).not.toContain("M43.5 24C43.5 37.5");
      expect(read(file), file).not.toContain("M36 34.2C37.4 42");
    }
  });
});
