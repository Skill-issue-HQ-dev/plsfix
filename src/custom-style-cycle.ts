// The style-cycle key's decision as plain data: which style follows the one the active cell wears,
// the workbook's own spelling of the listed names, and the receipt. No Office.js: the adapter
// src/excel/custom-style-cycle.ts reads and writes. Invariant: the cycle is a ring over the list
// plus Excel's Normal (first ... last, Normal, first ...), entered from any other style at the first.

import { sameName, type RegistryEntry } from "./custom-styles";

/** Excel's built-in default style, where the ring passes after the last of the list. */
export const NORMAL_STYLE = "Normal";
export const NO_STYLES_YET = "Save a style first: Tools > Custom styles.";

export interface CycleStep {
  /** The style to put on the selection: one of the list, or Normal after the last. */
  name: string;
  /** Its 1-based place in the list; null for Normal, which is not one of it. */
  place: number | null;
  /** How many styles the list holds. */
  total: number;
}

/**
 * The style that follows `worn` in the ring, or null when the list is empty
 * (nothing to step through). A style that is not in the list, Normal included,
 * is entered at the first.
 */
export function nextStyle(
  names: readonly string[],
  worn: string | null | undefined,
): CycleStep | null {
  const total = names.length;
  const first = names[0];
  if (first === undefined) return null;
  const at = worn ? names.findIndex((name) => sameName(name, worn)) : -1;
  if (at < 0) return { name: first, place: 1, total };
  const next = names[at + 1];
  if (next === undefined) return { name: NORMAL_STYLE, place: null, total };
  return { name: next, place: at + 2, total };
}

/** "Style: Margin % (2 of 5)", or "Style: Normal" for the step past the last. */
export function cycleReceipt(step: CycleStep): string {
  return step.place === null
    ? `Style: ${step.name}`
    : `Style: ${step.name} (${String(step.place)} of ${String(step.total)})`;
}

/**
 * The listed styles the workbook still has, in list order, each in the
 * workbook's own spelling (the name a style is put on a cell by).
 */
export function listedNames(
  entries: readonly RegistryEntry[],
  existing: readonly string[],
): string[] {
  return entries.flatMap((entry) => {
    const own = existing.find((name) => sameName(name, entry.name));
    return own === undefined ? [] : [own];
  });
}
