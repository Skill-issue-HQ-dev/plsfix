// @vitest-environment jsdom
// The coverage gate for the help copy. Parses the shipped taskpane.html and
// pptpane.html, so a button added, renamed or deleted fails here rather than
// leaving a silent hole in the "?" card. Both directions are checked: no
// button without a sentence, and no sentence naming a button that is gone.

import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { drawStyleRows } from "../pane/custom-style-rows";
import { drawShapeStyleRows } from "../ppt/shape-style-rows";
import { renderTemplateList } from "../pane/template-picker";
import type { StylePreview } from "../custom-styles";
import type { ShapeStylePreview } from "../shape-styles";
import type { DeckRowView } from "../ppt/deck-rows";
import { renderLinkRows } from "../ppt/views";
import { TEMPLATES } from "../templates/index";
import { DEFAULT_SETTINGS } from "../settings";
import { renderPaletteEditor } from "../ui/palette-editor";
import {
  GENERATED_BUTTONS,
  HELP,
  HELP_MAX_CHARS,
  helpSections,
  type HelpSection,
} from "./copy";
import { DESIGN_HELP } from "./copy.design";
import { EXCEL_HELP } from "./copy.excel";
import { PPT_HELP } from "./copy.ppt";
import { TEMPLATE_HELP } from "./copy.templates";

const PANES = ["taskpane.html", "pptpane.html"] as const;

const SAMPLE_STYLE: StylePreview = {
  name: "Style name",
  kind: "user",
  fill: "#F2E8DF",
  fontColor: "#1F1D1B",
  bold: true,
};

const SAMPLE_SHAPE_STYLE: ShapeStylePreview = {
  name: "Style name",
  kind: "user",
  fill: "#F2E8DF",
  outline: { color: "#282623", weight: 1, dashed: false },
  textColor: "#1F1D1B",
  bold: true,
};

const SAMPLE_DECK_ROW: DeckRowView = {
  offSlide: true,
  key: "deck/0",
  id: "0".repeat(32),
  label: "Model!B4:F12 table",
  source: "Model.xlsx",
  kind: "table",
  pushedAt: "2026-10-03T08:00:00.000Z",
  canPutBack: true,
};

// The buttons a pane builds at boot rather than ships in its HTML (the
// Templates picker's Insert buttons, one per registry entry, and the palette
// editor's Add color) are built into the
// parsed copy first, so a generated button is held to the very coverage rules a
// hand-written one is: it needs a sentence, and a sentence needs its button.
function withGeneratedButtons(doc: Document): Document {
  const templates = doc.getElementById("template-list");
  if (templates !== null) renderTemplateList(templates, TEMPLATES);
  // The same editor in each pane: the Excel Brand tab's, the PowerPoint Settings tab's.
  for (const id of ["brand-colors-editor", "ppt-brand-editor"]) {
    const palette = doc.getElementById(id);
    if (palette === null) continue;
    renderPaletteEditor(palette, {
      settings: DEFAULT_SETTINGS,
      onPreview: () => undefined,
      onCommit: () => undefined,
    });
  }
  const styleList = doc.getElementById("custom-style-list");
  const styleNote = doc.getElementById("custom-style-empty");
  if (styleList !== null && styleNote !== null) {
    // One style row: its apply, Update and Delete buttons are the section's
    // generated ones (the apply button reads as the style's own name).
    drawStyleRows(styleList, styleNote, { styles: [SAMPLE_STYLE] });
  }
  const shapeList = doc.getElementById("shape-style-list");
  const shapeNote = doc.getElementById("shape-style-empty");
  if (shapeList !== null && shapeNote !== null) {
    // One shape style row: its apply, Update and Delete buttons are the
    // section's generated ones, the way the Excel list's are.
    drawShapeStyleRows(shapeList, shapeNote, { styles: [SAMPLE_SHAPE_STYLE] });
  }
  const deckRows = doc.getElementById("link-rows");
  if (deckRows !== null) {
    // One Not on a slide row: its Put back and Remove from this deck buttons
    // are the deck list's generated ones.
    renderLinkRows(deckRows as HTMLTableSectionElement, [], () => undefined, {
      rows: [SAMPLE_DECK_ROW],
      onPutBack: () => undefined,
      onRemove: () => undefined,
    });
  }
  return doc;
}

// What a generated button reads as: an icon button, whose face is a glyph, by its
// aria-label (the way the card names one), every other button by its own text.
function faceOf(button: HTMLElement): string {
  return button.classList.contains("icon-button")
    ? (button.getAttribute("aria-label") ?? "")
    : (button.textContent ?? "");
}

// The jsdom environment turns import.meta.url into an http URL, so the panes
// are read from the vitest root instead - which is this repo.
function sectionsOf(pane: string): HelpSection[] {
  const html = readFileSync(path.join(process.cwd(), pane), "utf8");
  const parsed = new DOMParser().parseFromString(html, "text/html");
  return helpSections(withGeneratedButtons(parsed));
}

const PANE_SECTIONS = new Map(PANES.map((pane) => [pane, sectionsOf(pane)]));

function everySection(): { pane: string; section: HelpSection }[] {
  const all: { pane: string; section: HelpSection }[] = [];
  for (const [pane, sections] of PANE_SECTIONS) {
    for (const section of sections) all.push({ pane, section });
  }
  return all;
}

describe("help copy", () => {
  it("finds the sections of both panes", () => {
    expect(PANE_SECTIONS.get("taskpane.html")!.length).toBeGreaterThan(10);
    expect(PANE_SECTIONS.get("pptpane.html")!.length).toBeGreaterThan(2);
  });

  // One key means one section: the two panes both call a section "Linked
  // objects", and the deck's one carries data-help to stay apart from Excel's.
  it("keys every section apart across the two panes", () => {
    const seen = new Map<string, string>();
    const clashes: string[] = [];
    for (const { pane, section } of everySection()) {
      const first = seen.get(section.key);
      if (first !== undefined)
        clashes.push(`${section.key}: ${first}, ${pane}`);
      else seen.set(section.key, pane);
    }
    expect(clashes).toEqual([]);
  });

  it("has a line on what every section is for", () => {
    const missing = everySection()
      .filter(({ section }) => HELP[section.key] === undefined)
      .map(({ pane, section }) => `${pane} > ${section.key}`);
    expect(missing).toEqual([]);
  });

  it("has a sentence for every button in every section", () => {
    const missing: string[] = [];
    for (const { pane, section } of everySection()) {
      const copy = HELP[section.key];
      if (copy === undefined) continue;
      for (const button of section.buttons) {
        if (copy.buttons[button.key] === undefined) {
          missing.push(`${pane} > ${section.key} > ${button.key}`);
        }
      }
    }
    expect(missing).toEqual([]);
  });

  it("names no button that the panes no longer ship", () => {
    const stale: string[] = [];
    const covered = new Set<string>();
    for (const { section } of everySection()) {
      covered.add(section.key);
      const copy = HELP[section.key];
      if (copy === undefined) continue;
      const keys = new Set(section.buttons.map((button) => button.key));
      for (const key of Object.keys(copy.buttons)) {
        if (!keys.has(key)) stale.push(`${section.key} > ${key}`);
      }
    }
    for (const key of Object.keys(HELP)) {
      if (!covered.has(key)) stale.push(`${key} (no such section)`);
    }
    expect(stale).toEqual([]);
  });

  it("keeps every sentence short enough to read in a card", () => {
    const long: string[] = [];
    for (const [key, copy] of Object.entries(HELP)) {
      if (copy.about.length > HELP_MAX_CHARS) {
        long.push(`${key} (about, ${String(copy.about.length)})`);
      }
      if ((copy.note?.length ?? 0) > HELP_MAX_CHARS) {
        long.push(`${key} (note, ${String(copy.note?.length)})`);
      }
      for (const [button, text] of Object.entries(copy.buttons)) {
        if (text.length > HELP_MAX_CHARS) {
          long.push(`${key} > ${button} (${String(text.length)})`);
        }
      }
    }
    expect(long).toEqual([]);
  });

  it("writes plain sentences: no em dash, every one closed", () => {
    const bad: string[] = [];
    for (const [key, copy] of Object.entries(HELP)) {
      for (const [button, text] of [
        ["about", copy.about] as const,
        ...(copy.note === undefined ? [] : [["note", copy.note] as const]),
        ...Object.entries(copy.buttons),
      ]) {
        if (text.includes("—")) bad.push(`${key} > ${button} (em dash)`);
        if (!text.endsWith(".")) bad.push(`${key} > ${button} (no full stop)`);
      }
    }
    expect(bad).toEqual([]);
  });

  // PowerPoint's own Undo has not been shown to take back a swatch or a change of
  // the deck theme on any host: the cards promise it of neither. Restore theme
  // colors is the way back for the theme, and its sentence says so.
  it("promises no PowerPoint Undo for a brand swatch or the deck theme", () => {
    for (const key of ["ppt-swatches", "ppt-brand"]) {
      const copy = HELP[key]!;
      for (const text of [
        copy.about,
        copy.note ?? "",
        ...Object.values(copy.buttons),
      ]) {
        expect(text, key).not.toMatch(/undo/i);
      }
    }
    expect(HELP["ppt-brand"]!.buttons["restore-theme-colors"]).toMatch(
      /puts back the theme colors/i,
    );
  });

  // The cards speak US English, like the pane.
  it("spells color the US way in every card", () => {
    for (const [key, copy] of Object.entries(HELP)) {
      for (const text of [
        copy.about,
        copy.note ?? "",
        ...Object.values(copy.buttons),
      ]) {
        expect(text, key).not.toMatch(/colour/i);
      }
    }
  });
});

// The card of a section is built before the buttons a pane makes later exist, so
// those are listed by fixed labels (GENERATED_BUTTONS). Held to the real thing:
// a label that differs from the button's own, or a key with no sentence, fails.
describe("the generated buttons the card lists ahead of their rows", () => {
  it("carry the label the real button shows, and each has a sentence", () => {
    for (const [sectionKey, generated] of Object.entries(GENERATED_BUTTONS)) {
      const section = everySection().find(
        (one) => one.section.key === sectionKey,
      )!.section;
      for (const [key, label] of Object.entries(generated)) {
        expect(HELP[sectionKey]!.buttons[key], key).toBeDefined();
        const real = section.buttons.find((button) => button.key === key);
        expect(real && faceOf(real.element), key).toBe(label);
      }
    }
  });
});

// The Links tab's export section, button by button: Export selection sends a
// native table, a picture only comes from its own button, and the card names
// nothing that is gone. The generic rules above hold the sentences to the
// markup; this holds the markup's order and what the two range cards promise.
// The status floor (L3) added a word to the Linked objects list; the card
// that explains the buttons has to explain it too, in the button that meets it.
describe("the deck list's In this deck state", () => {
  it("is explained where Update all meets it, and the filter lists it", () => {
    const sentence = PPT_HELP["deck-links"]!.buttons["update-all"]!;
    expect(sentence).toContain("In this deck");
    const html = readFileSync(
      path.resolve(__dirname, "../../pptpane.html"),
      "utf8",
    );
    expect(html).toContain('<option value="inDeck">In this deck</option>');
  });
});

describe("the Links tab's export help", () => {
  const section = (): HelpSection =>
    PANE_SECTIONS.get("taskpane.html")!.find(
      (one) => one.key === "export-heading",
    )!;
  const sentence = (button: string): string =>
    HELP["export-heading"]!.buttons[button] ?? "";

  it("explains the four export buttons in the order the pane shows them", () => {
    // The section also holds the Copy for PowerPoint bar, local mode's fallback.
    expect(section().buttons.map((button) => button.key)).toEqual([
      "export-selection",
      "export-chart",
      "export-text",
      "export-picture",
      "copy-for-powerpoint",
    ]);
    expect(Object.keys(HELP["export-heading"]!.buttons)).not.toContain(
      "export-table",
    );
  });

  it("says Export selection sends a table that refreshes in place", () => {
    expect(sentence("export-selection")).toMatch(/native table/);
    expect(sentence("export-selection")).toMatch(/60 rows and 20 columns/);
    expect(sentence("export-selection")).not.toMatch(/picture/);
  });

  it("says Export as picture is a linked picture, for a layout a table cannot keep", () => {
    expect(sentence("export-picture")).toMatch(/linked picture/);
    expect(sentence("export-picture")).toMatch(/table cannot keep/);
  });
});

// The real panes never hit these branches (every section has a heading and a
// key, every button carries a data-action or a real id), so they are
// exercised over hand-built markup instead.
describe("helpSections: defensive branches hand-built markup exercises", () => {
  function parse(html: string): Document {
    return new DOMParser().parseFromString(html, "text/html");
  }

  it("skips a section with no heading to hang the toggle on", () => {
    const doc = parse(
      '<section data-help="x"><p>No heading here.</p></section>',
    );
    expect(helpSections(doc)).toEqual([]);
  });

  it("skips a section with a heading but no data-help or aria-labelledby", () => {
    const doc = parse(
      '<section><div class="section-heading"><h2>Untitled</h2></div></section>',
    );
    expect(helpSections(doc)).toEqual([]);
  });

  it("leaves out a button with neither data-action nor a real id", () => {
    const doc = parse(
      '<section data-help="x"><div class="section-heading"><h2>X</h2></div>' +
        '<button id="named"></button><button></button></section>',
    );
    expect(helpSections(doc)[0]!.buttons.map((b) => b.key)).toEqual(["named"]);
  });
});

// HELP is the four files spread into one object: a section key in two of them
// would silently overwrite the first, so each key is owned by exactly one file.
describe("the copy files", () => {
  const FILES = {
    "copy.excel.ts": EXCEL_HELP,
    "copy.templates.ts": TEMPLATE_HELP,
    "copy.design.ts": DESIGN_HELP,
    "copy.ppt.ts": PPT_HELP,
  };

  it("never repeat a section key between them", () => {
    const owner = new Map<string, string>();
    const repeats: string[] = [];
    for (const [file, copy] of Object.entries(FILES)) {
      for (const key of Object.keys(copy)) {
        const first = owner.get(key);
        if (first !== undefined) repeats.push(`${key}: ${first} and ${file}`);
        owner.set(key, file);
      }
    }
    expect(repeats).toEqual([]);
  });

  it("lose nothing in the merge: HELP holds every section of every file", () => {
    const keys = Object.values(FILES).flatMap((copy) => Object.keys(copy));
    expect(Object.keys(HELP).sort()).toEqual(keys.sort());
  });

  it("keep the whole Templates section in copy.templates.ts", () => {
    expect(Object.keys(TEMPLATE_HELP)).toEqual(["templates-heading"]);
    expect(Object.keys(EXCEL_HELP)).not.toContain("templates-heading");
    const ids = TEMPLATES.map((entry) => entry.id).sort();
    const buttons = TEMPLATE_HELP["templates-heading"]?.buttons ?? {};
    expect(Object.keys(buttons).sort()).toEqual(ids);
  });
});
