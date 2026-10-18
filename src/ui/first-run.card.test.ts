// @vitest-environment jsdom
// The shipped "New here?" card of the Excel pane (taskpane.html) points a new user at the two
// design-loop sections that sit low on the Tools tab: it names Brand colors and Custom styles by
// the headings they carry and says where they are, and the card stays a short checklist.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it } from "vitest";

let doc: Document;

beforeEach(() => {
  doc = new DOMParser().parseFromString(
    readFileSync(join(process.cwd(), "taskpane.html"), "utf8"),
    "text/html",
  );
});

const text = (node: Element | null): string =>
  (node?.textContent ?? "").replace(/\s+/g, " ").trim();

const cardLines = (): string[] =>
  [...doc.querySelectorAll("#first-run > p")].map((line) => text(line));

describe("the Excel first-run card", () => {
  it("has one line naming Brand colors and Custom styles", () => {
    const lines = cardLines().filter(
      (line) => line.includes("Brand colors") && line.includes("Custom styles"),
    );
    expect(lines).toHaveLength(1);
  });

  it("names the two sections by the headings they carry", () => {
    expect(text(doc.getElementById("brand-colors-heading"))).toBe(
      "Brand colors",
    );
    expect(text(doc.getElementById("custom-styles-heading"))).toBe(
      "Custom styles",
    );
  });

  it("says where they are: below Model formatting, which is the section right above them", () => {
    const line = cardLines().find((entry) => entry.includes("Custom styles"));
    expect(line).toContain("below Model formatting");

    const sections = [...doc.querySelectorAll("#view-tools > section")];
    const at = (headingId: string): number =>
      sections.findIndex((section) => section.querySelector(`#${headingId}`));
    expect(text(doc.getElementById("format-heading"))).toBe("Model formatting");
    expect(at("brand-colors-heading")).toBe(at("format-heading") + 1);
    expect(at("custom-styles-heading")).toBe(at("brand-colors-heading") + 1);
  });

  it("stays a short checklist: numbered lines in order, the ? tip last", () => {
    const lines = cardLines();
    expect(lines.length).toBeLessThanOrEqual(5);
    lines.forEach((line, index) => {
      expect(line.startsWith(`${String(index + 1)}. `), line).toBe(true);
    });
    expect(lines.at(-1)).toContain("Every section has a ?");
  });

  it("keeps the line short enough for a narrow dock", () => {
    for (const line of cardLines()) expect(line.length).toBeLessThan(120);
  });
});

// The two items under the checklist, in a style of their own (a list under the numbered lines, so
// the checklist above stays four short steps): where Undo is, and that Find a tool takes you to a
// tool. Every claim in them is held to the pane it describes, so the card cannot drift from it.
describe("the Excel first-run card: Undo and Find a tool", () => {
  const tips = (): Element[] => [
    ...doc.querySelectorAll("#first-run .first-run-tips > li"),
  ];

  it("adds two items, Undo then Find a tool, in the card's own words", () => {
    expect(tips().map((tip) => text(tip))).toEqual([
      "Undo: Excel's Ctrl+Z cannot take back a pls,fix button. Press Ctrl+Shift+Z or Undo last pls,fix action (Tools > Model tools).",
      "Find a tool: type what you want to do, Enter takes you there.",
    ]);
  });

  it("sets each item's own word in bold, apart from its sentence", () => {
    expect(tips().map((tip) => text(tip.querySelector("strong")))).toEqual([
      "Undo:",
      "Find a tool:",
    ]);
  });

  it("sits under the numbered checklist and above the card's buttons", () => {
    const card = doc.getElementById("first-run")!;
    const order = [...card.children].map(
      (child) => child.className || child.tagName.toLowerCase(),
    );
    const list = order.indexOf("first-run-tips");
    expect(list).toBeGreaterThan(order.lastIndexOf("p"));
    expect(list).toBeLessThan(order.indexOf("io-row"));
  });

  it("names the key and the place the Undo button really has", () => {
    const button = doc.querySelector<HTMLElement>('[data-action="undo"]')!;
    expect(text(button.querySelector("strong"))).toBe(
      "Undo last pls,fix action",
    );
    expect(text(button.querySelector("kbd"))).toBe("Ctrl+Shift+Z");
    const section = button.closest("section")!;
    expect(text(section.querySelector("h2"))).toBe("Model tools");
    expect(section.closest("#view-tools")).not.toBeNull();

    const bindings = JSON.parse(
      readFileSync(join(process.cwd(), "public/shortcuts.json"), "utf8"),
    ) as { shortcuts: { action: string; key: { default: string } }[] };
    expect(
      bindings.shortcuts.find((entry) => entry.action === "PLSFIX_UNDO")?.key
        .default,
    ).toBe("Ctrl+Shift+Z");
  });

  it("names the search box by the label it carries", () => {
    const box = doc.getElementById("tool-search")!;
    expect(box.getAttribute("aria-label")).toBe("Find a tool");
  });
});
