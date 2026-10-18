// The Brand tab's seven color slots (Primary, Accent, the five modelling colors), wired apart from the
// tab's other settings: a picker previews while it is dragged and saves when it is let go, a hex field
// saves on change. Primary and Accent are brand colors, so a pick that repeats another brand color is
// refused on commit, in the palette editor's status line, and the slot keeps the color it had.

import { heldElsewhere, repeatSentence } from "../palette";
import {
  type BrandSettings,
  getActiveSettings,
  normalizeHex,
} from "../settings";
import { getElement } from "../ui/dom";
import { sayInEditor } from "../ui/palette-editor";
import { toast } from "./shared";

export const PALETTE_SLOTS = [
  "primary",
  "accent",
  "input",
  "formula",
  "link",
  "external",
  "partial",
] as const;
export type PaletteSlot = (typeof PALETTE_SLOTS)[number];

/** What the Brand tab does with the settings a slot produces. */
export interface SlotHooks {
  /** Applied in memory and drawn, not saved: an edit still being made. */
  preview: (next: BrandSettings) => void;
  /** Applied and saved, here and in the workbook. */
  commit: (patch: Partial<BrandSettings>) => void;
  /** The tab drawn again from the active settings. */
  render: () => void;
}

// Where a brand slot sits in the brand list; the modelling colors are not brand colors.
const BRAND_INDEX: Partial<Record<PaletteSlot, number>> = {
  primary: 0,
  accent: 1,
};

// A brand slot's color as it stood before the first preview of a drag that no commit has
// settled yet: the color a refused pick goes back to, not one the drag passed on its way.
const before = new Map<PaletteSlot, string>();

function drag(slot: PaletteSlot, value: string, hooks: SlotHooks): void {
  const current = getActiveSettings();
  if (BRAND_INDEX[slot] !== undefined && !before.has(slot)) {
    before.set(slot, current[slot]);
  }
  hooks.preview({ ...current, [slot]: value });
}

/** A refused repeat, said where every refusal is said: the palette editor's status line. */
export function sayRepeat(hex: string): void {
  sayInEditor(getElement("brand-colors-editor"), repeatSentence(hex));
}

// A Primary or Accent that repeats another brand color is refused: the settings go back to
// what they were before the drag and the editor's status line says why.
function refused(slot: PaletteSlot, value: string, hooks: SlotHooks): boolean {
  const index = BRAND_INDEX[slot];
  const was = before.get(slot);
  before.delete(slot);
  if (index === undefined) return false;
  const current = getActiveSettings();
  const back: BrandSettings = { ...current, [slot]: was ?? current[slot] };
  if (!heldElsewhere(back, index, value)) return false;
  hooks.preview(back);
  sayRepeat(value);
  return true;
}

function commit(slot: PaletteSlot, value: string, hooks: SlotHooks): void {
  if (!refused(slot, value, hooks)) hooks.commit({ [slot]: value });
}

/** Wires every slot's picker and hex field. Called once at boot, after the markup exists. */
export function wireSlots(hooks: SlotHooks): void {
  for (const slot of PALETTE_SLOTS) {
    const picker = document.querySelector<HTMLInputElement>(
      `[data-slot-color="${slot}"]`,
    );
    const hex = document.querySelector<HTMLInputElement>(
      `[data-slot-hex="${slot}"]`,
    );
    picker?.addEventListener("input", () => {
      const value = normalizeHex(picker.value);
      if (value) drag(slot, value, hooks);
    });
    picker?.addEventListener("change", () => {
      const value = normalizeHex(picker.value);
      if (value) commit(slot, value, hooks);
    });
    hex?.addEventListener("change", () => {
      const value = normalizeHex(hex.value);
      if (!value) {
        toast.show(`"${hex.value}" is not a hex color like #6D7E12.`, "error");
        hooks.render();
        return;
      }
      commit(slot, value, hooks);
    });
  }
}
