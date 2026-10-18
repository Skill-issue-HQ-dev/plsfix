// The brand's fill and font ladders, as the steps Ctrl+Shift+F / G walk: today's
// steps exactly for a two-color palette, brand colors 3 and up appended, and a
// color equal to an earlier step skipped so a cycle can never stall on a
// repeat. Pure; each step names its slot, cycle-labels.ts turns that into words.
import { brandColors } from "./palette";
import { deriveTheme, type BrandSettings } from "./settings";

/** A named palette slot, or a position in the brand list (2 is its third color). */
export type LadderSlot =
  | "headerTint"
  | "resultTint"
  | "accent"
  | "primary"
  | "formula"
  | "input"
  | "link"
  | number;

export interface LadderStep {
  color: string;
  slot: LadderSlot;
}

// The first step of each color stays; later ones with the same color go.
function distinct(steps: LadderStep[]): LadderStep[] {
  const seen = new Set<string>();
  return steps.filter(({ color }) => {
    const key = color.toUpperCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function extraSteps(settings: BrandSettings): LadderStep[] {
  return brandColors(settings)
    .slice(2)
    .map((color, at) => ({ color, slot: at + 2 }));
}

/** The fill ladder, before the clear step the cycle ends on. */
export function fillLadder(settings: BrandSettings): LadderStep[] {
  const theme = deriveTheme(settings);
  return distinct([
    { color: theme.headerFill, slot: "headerTint" },
    { color: theme.resultFill, slot: "resultTint" },
    { color: settings.accent, slot: "accent" },
    { color: settings.primary, slot: "primary" },
    ...extraSteps(settings),
  ]);
}

/** The font color ladder: the model's own colors, then the brand's. */
export function fontLadder(settings: BrandSettings): LadderStep[] {
  const theme = deriveTheme(settings);
  return distinct([
    { color: theme.formulaFont, slot: "formula" },
    { color: theme.inputFont, slot: "input" },
    { color: theme.linkFont, slot: "link" },
    { color: settings.accent, slot: "accent" },
    { color: settings.primary, slot: "primary" },
    ...extraSteps(settings),
  ]);
}
