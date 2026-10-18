// Human names for what a format cycle just landed on, in the exact step
// order src/cycles.ts's own tables build - never Excel's own format code or
// a hex colour, which reads as line noise to the modeller a toast is for.
// Split out of cycles.ts (already at its own 400-line cap) rather than
// pushed further over it. Pure: no Office.js, no DOM.

import { fillLadder, fontLadder, type LadderSlot } from "./brand-ladders";
import { currencyFormat } from "./numbers";
import type { BrandSettings } from "./settings";
import type { NumberCycleFamily, RowStyleKind } from "./cycles";

// buildNumberCycles' own step order for the three families a currency never
// touches. A currency symbol places before or after the digits by language,
// so its own labels are built fresh from the live settings below instead.
const STATIC_NUMBER_LABELS: Record<
  Exclude<NumberCycleFamily, "currency">,
  string[]
> = {
  general: ["1,234", "1,234.0", "1,234.00"],
  percent: ["12.3%", "12%", "12.34%"],
  multiple: ["0.0x", "0.00x"],
  date: ["dd.mm.yyyy", "mmm-yy", "yyyy"],
};

// The trailing comma in the cycle's third currency format divides the
// displayed value by a thousand - Excel's own scaling trick for a model kept
// in whole currency but shown in thousands.
function currencyLabels(settings: BrandSettings): string[] {
  const { currency, language } = settings;
  return [
    currencyFormat(currency, language, "1,234"),
    currencyFormat(currency, language, "1,234.0"),
    `${currencyFormat(currency, language, "1,234")} (thousands)`,
  ];
}

/** Same shape and step order as buildNumberCycles(settings). */
export function numberCycleLabels(
  settings: BrandSettings,
): Record<NumberCycleFamily, string[]> {
  return { ...STATIC_NUMBER_LABELS, currency: currencyLabels(settings) };
}

/** buildRowStyleCycles' own step order. */
export const ROW_STYLE_LABELS: Record<RowStyleKind, string[]> = {
  title: ["Filled", "Header rule", "Accent rule"],
  result: ["Filled", "Double rule", "Accent block"],
  item: ["Plain", "Header tint", "Result tint"],
};

// buildFillCycle's own step order for a two-color palette (fillCycleLabels
// below answers any palette). headerFill/resultFill are theme tints of
// primary/accent (deriveTheme), not palette slots of their own, so they get
// a descriptive name rather than the settings key nobody edits directly.
export const FILL_CYCLE_LABELS = [
  "Header tint",
  "Result tint",
  "Accent",
  "Primary",
  "none",
];

// buildFontCycle's own step order for a two-color palette, named the way the
// Brand tab's own palette rows already label these slots (taskpane.html).
export const FONT_CYCLE_LABELS = [
  "Formulas",
  "Inputs",
  "Cross-sheet links",
  "Accent",
  "Primary",
];

// The same two ladders for a palette with any number of brand colors: one
// label per step the cycle really has, repeats skipped on both sides alike, the
// extras named by their place in the brand list ("Brand 3" is the third color).
const SLOT_LABELS: Record<Exclude<LadderSlot, number>, string> = {
  headerTint: "Header tint",
  resultTint: "Result tint",
  accent: "Accent",
  primary: "Primary",
  formula: "Formulas",
  input: "Inputs",
  link: "Cross-sheet links",
};

function slotLabel(slot: LadderSlot): string {
  return typeof slot === "number"
    ? `Brand ${String(slot + 1)}`
    : SLOT_LABELS[slot];
}

/** buildFillCycle(settings)'s step names, the clear step last. */
export function fillCycleLabels(settings: BrandSettings): string[] {
  return [...fillLadder(settings).map((step) => slotLabel(step.slot)), "none"];
}

/** buildFontCycle(settings)'s step names. */
export function fontCycleLabels(settings: BrandSettings): string[] {
  return fontLadder(settings).map((step) => slotLabel(step.slot));
}

// buildBorderCycle's own step order: nothing, a rule under the row, a
// heavier one, the double rule a total takes, a box, then the full grid.
export const BORDER_CYCLE_LABELS = [
  "none",
  "underline",
  "heavy underline",
  "double underline",
  "box",
  "grid",
];

const ALIGN_LABELS: Record<string, string> = {
  Left: "left",
  Center: "centre",
  Right: "right",
  General: "general",
};

const UNDERLINE_LABELS: Record<string, string> = {
  Single: "single",
  Double: "double",
  None: "none",
};

export function alignLabel(value: string): string {
  return ALIGN_LABELS[value] ?? value.toLowerCase();
}

export function underlineLabel(value: string): string {
  return UNDERLINE_LABELS[value] ?? value.toLowerCase();
}
