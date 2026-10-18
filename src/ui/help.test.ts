// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from "vitest";
import { HELP } from "../help/copy";
import { installHelp } from "./help";

// Real section and button keys, so the test also proves the wiring reaches the
// shipped copy rather than a fixture of its own.
const TOC = `
  <section aria-labelledby="toc-heading">
    <div class="section-heading"><div><h2 id="toc-heading">Table of contents</h2></div></div>
    <button data-action="insert-toc" type="button">
      <span class="action-icon">x</span><span><strong>Insert contents sheet</strong><small>note</small></span>
    </button>
  </section>`;
const SHEETS = `
  <section aria-labelledby="sheets-heading">
    <div class="section-heading"><div><h2 id="sheets-heading">Sheet explorer</h2></div>
      <button id="refresh-sheets" class="icon-button" type="button" aria-label="Refresh the sheet list">R</button>
    </div>
  </section>`;

function toggles(): HTMLButtonElement[] {
  return Array.from(document.querySelectorAll(".help-toggle"));
}

function cardOf(toggle: HTMLElement): HTMLElement {
  return document.getElementById(toggle.getAttribute("aria-controls")!)!;
}

describe("installHelp", () => {
  beforeEach(() => {
    document.body.innerHTML = TOC + SHEETS;
  });

  it("puts a keyboard-reachable ? in every section heading", () => {
    installHelp(document);
    expect(toggles()).toHaveLength(2);
    for (const toggle of toggles()) {
      expect(toggle.tagName).toBe("BUTTON");
      expect(toggle.type).toBe("button");
      expect(toggle.textContent).toBe("?");
      expect(toggle.getAttribute("aria-expanded")).toBe("false");
      expect(toggle.closest(".section-heading")).not.toBeNull();
    }
    expect(toggles()[0]!.getAttribute("aria-label")).toContain(
      "Table of contents",
    );
  });

  it("starts with every card closed, under its own heading", () => {
    installHelp(document);
    for (const toggle of toggles()) {
      const card = cardOf(toggle);
      expect(card.hidden).toBe(true);
      expect(card.previousElementSibling).toBe(
        toggle.closest(".section-heading"),
      );
    }
  });

  it("opens a card listing each button's label and its sentence", () => {
    installHelp(document);
    const toggle = toggles()[0]!;
    toggle.click();
    const card = cardOf(toggle);
    expect(card.hidden).toBe(false);
    expect(toggle.getAttribute("aria-expanded")).toBe("true");
    expect(card.textContent).toContain(HELP["toc-heading"]!.about);
    expect(card.querySelector("dt")!.textContent).toBe("Insert contents sheet");
    expect(card.querySelector("dd")!.textContent).toBe(
      HELP["toc-heading"]!.buttons["insert-toc"],
    );
  });

  it("labels an icon button by its aria-label, not its glyph", () => {
    installHelp(document);
    const toggle = toggles()[1]!;
    toggle.click();
    expect(cardOf(toggle).querySelector("dt")!.textContent).toBe(
      "Refresh the sheet list",
    );
  });

  it("closes the card when its own ? is pressed again", () => {
    installHelp(document);
    const toggle = toggles()[0]!;
    toggle.click();
    toggle.click();
    expect(cardOf(toggle).hidden).toBe(true);
    expect(toggle.getAttribute("aria-expanded")).toBe("false");
  });

  it("keeps one card open at a time", () => {
    installHelp(document);
    const [first, second] = toggles();
    first!.click();
    second!.click();
    expect(cardOf(first!).hidden).toBe(true);
    expect(first!.getAttribute("aria-expanded")).toBe("false");
    expect(cardOf(second!).hidden).toBe(false);
  });

  it("closes on Escape and puts focus back on the ?", () => {
    installHelp(document);
    const toggle = toggles()[0]!;
    toggle.click();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(cardOf(toggle).hidden).toBe(true);
    expect(document.activeElement).toBe(toggle);
  });

  it("leaves a section the copy does not cover alone", () => {
    document.body.innerHTML = `
      <section aria-labelledby="nowhere-heading">
        <div class="section-heading"><div><h2 id="nowhere-heading">Nowhere</h2></div></div>
        <button data-action="apply-nowhere" type="button">Apply</button>
      </section>`;
    installHelp(document);
    expect(toggles()).toHaveLength(0);
  });

  it("remembers nothing: a fresh install opens closed", () => {
    installHelp(document);
    toggles()[0]!.click();
    document.body.innerHTML = TOC;
    installHelp(document);
    expect(cardOf(toggles()[0]!).hidden).toBe(true);
  });
});

// The Linked objects card of the PowerPoint pane is built before any row exists:
// the buttons of a Not on a slide row are made later, so their sentences reach
// the card by a list of generated keys with fixed labels (src/help/copy.ts).
describe("a card with buttons the pane builds later", () => {
  const DECK = `
    <section data-help="deck-links">
      <div class="section-heading"><div><h2>Linked objects</h2></div></div>
      <button id="refresh-links" type="button">Refresh</button>
      <table><tbody id="link-rows"></tbody></table>
    </section>`;

  function deckCard(): HTMLElement {
    installHelp(document);
    const toggle = toggles()[0]!;
    toggle.click();
    return cardOf(toggle);
  }

  it("lists the generated buttons by their fixed labels, though no row exists yet", () => {
    document.body.innerHTML = DECK;
    const card = deckCard();
    const rows = new Map(
      [...card.querySelectorAll(".help-row")].map((row) => [
        row.querySelector("dt")!.textContent,
        row.querySelector("dd")!.textContent,
      ]),
    );
    const copy = HELP["deck-links"]!.buttons;
    expect(rows.get("Put back")).toBe(copy["deck-put-back"]);
    expect(rows.get("Remove from this deck")).toBe(copy["deck-remove"]);
    // The static button keeps its own entry, read off the page as ever.
    expect(rows.get("Refresh")).toBe(copy["refresh-links"]);
  });

  it("lists a generated button once: a real one on the page keeps its own entry", () => {
    document.body.innerHTML = DECK.replace(
      '<tbody id="link-rows"></tbody>',
      '<tbody id="link-rows"><tr><td><button data-action="deck-put-back" aria-label="Put back Model!B4">Put back</button></td></tr></tbody>',
    );
    const terms = [...deckCard().querySelectorAll("dt")].map(
      (term) => term.textContent,
    );
    expect(terms.filter((term) => term?.startsWith("Put back"))).toHaveLength(
      1,
    );
  });

  it("says when a deleted object shows up, under the section's line", () => {
    document.body.innerHTML = DECK;
    const card = deckCard();
    const note = card.querySelector(".help-note")!;
    expect(note.textContent).toBe(HELP["deck-links"]!.note);
    expect(note.textContent).toMatch(/half a minute/);
    expect(note.textContent).toMatch(/never from the first scan/);
    expect(note.previousElementSibling).toBe(card.querySelector(".help-about"));
  });
});
