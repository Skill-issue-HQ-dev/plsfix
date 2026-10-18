// The "?" on every section heading and the card it opens: what the section is
// for, then one line per button. Owns the toggle, the card and the rule that
// one card is open at a time; the sentences are data in src/help/copy.ts and
// the labels are read off the buttons themselves, so they can never drift.
// Nothing is remembered - closing the pane closes the help with it.

import {
  GENERATED_BUTTONS,
  HELP,
  helpSections,
  type HelpSection,
  type SectionHelp,
} from "../help/copy";

interface OpenCard {
  toggle: HTMLElement;
  card: HTMLElement;
}

// What the button reads as: its aria-label when the face is a glyph, else the
// bold line of an action row, else whatever text it carries.
function labelOf(button: HTMLElement): string {
  const face =
    button.getAttribute("aria-label") ??
    button.querySelector("strong")?.textContent ??
    button.textContent ??
    "";
  return face.replace(/\s+/g, " ").trim() || button.id;
}

function titleOf(section: HelpSection): string {
  const heading = section.heading.querySelector("h2")?.textContent ?? "";
  return heading.replace(/\s+/g, " ").trim() || section.key;
}

function entry(label: string, sentence: string): HTMLElement {
  const row = document.createElement("div");
  row.className = "help-row";
  const term = document.createElement("dt");
  term.textContent = label;
  const detail = document.createElement("dd");
  detail.textContent = sentence;
  row.append(term, detail);
  return row;
}

function paragraph(text: string, className: string): HTMLElement {
  const line = document.createElement("p");
  line.className = className;
  line.textContent = text;
  return line;
}

function buildCard(
  id: string,
  section: HelpSection,
  copy: SectionHelp,
): HTMLElement {
  const card = document.createElement("div");
  card.id = id;
  card.className = "help-card";
  card.hidden = true;
  card.append(paragraph(copy.about, "help-about"));
  if (copy.note !== undefined) {
    card.append(paragraph(copy.note, "help-about help-note"));
  }

  const list = document.createElement("dl");
  list.className = "help-list";
  for (const button of section.buttons) {
    const sentence = copy.buttons[button.key];
    if (sentence !== undefined) {
      list.append(entry(labelOf(button.element), sentence));
    }
  }
  // The buttons a pane builds later (a row's) are not on the page yet: listed
  // from the copy by their fixed labels, unless a real one is already there.
  const present = new Set(section.buttons.map((button) => button.key));
  const later = Object.entries(GENERATED_BUTTONS[section.key] ?? {});
  for (const [key, label] of later) {
    const sentence = copy.buttons[key];
    if (sentence !== undefined && !present.has(key)) {
      list.append(entry(label, sentence));
    }
  }
  if (list.childElementCount > 0) card.append(list);
  return card;
}

function buildToggle(id: string, cardId: string, title: string): HTMLElement {
  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.id = id;
  toggle.className = "help-toggle";
  toggle.setAttribute("aria-expanded", "false");
  toggle.setAttribute("aria-controls", cardId);
  toggle.setAttribute("aria-label", `About ${title}`);
  const mark = document.createElement("span");
  mark.textContent = "?";
  toggle.append(mark);
  return toggle;
}

/**
 * Adds the "?" toggle and its card to every section of `root` the copy covers.
 * A section with no entry is left untouched - the copy gate is what reports it,
 * not the pane at runtime.
 */
export function installHelp(root: ParentNode): void {
  let open: OpenCard | null = null;

  function close(): void {
    if (open === null) return;
    open.card.hidden = true;
    open.toggle.setAttribute("aria-expanded", "false");
    open = null;
  }

  for (const section of helpSections(root)) {
    const copy = HELP[section.key];
    if (copy === undefined) continue;
    const cardId = `help-${section.key}`;
    const card = buildCard(cardId, section, copy);
    const toggle = buildToggle(`${cardId}-toggle`, cardId, titleOf(section));
    toggle.addEventListener("click", () => {
      const wasOpen = open?.toggle === toggle;
      close();
      if (wasOpen) return;
      card.hidden = false;
      toggle.setAttribute("aria-expanded", "true");
      open = { toggle, card };
    });
    section.heading.append(toggle);
    section.heading.insertAdjacentElement("afterend", card);
  }

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape" || open === null) return;
    const toggle = open.toggle;
    close();
    toggle.focus();
  });
}
