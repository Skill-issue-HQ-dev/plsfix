// The brand colors on the Excel pane: the Tools tab's swatch strips, the Brand
// tab's palette editor, Read from selection, Use these colors and Edit colors.
// Wiring only: lists and maths are src/palette.ts, widgets src/ui/, sheet writes
// src/excel/. Invariant: saving stays the Brand tab's (brand-tab.ts hands it in).

import { paintSwatch, readSelectionFills } from "../excel";
import { brandColors, MAX_BRAND_COLORS, withBrandList } from "../palette";
import { getActiveSettings, type BrandSettings } from "../settings";
import { getElement } from "../ui/dom";
import { applyBrandColors, renderPaletteEditor } from "../ui/palette-editor";
import { renderSwatches, type SwatchStrip } from "../ui/swatches";
import { guard, isExcelReady, toast } from "./shared";

/** What the Brand tab does with settings this module produces. */
export interface BrandColorsHooks {
  /** Applied in memory only: an edit still being made. */
  preview: (settings: BrandSettings) => void;
  /** Applied and saved, here and in the workbook. */
  commit: (settings: BrandSettings) => void;
}

// Excel writes a fill, and a font color; it cannot write an "Automatic" font
// color, so only the Fill strip has a none.
export const EXCEL_SWATCH_STRIPS: readonly SwatchStrip[] = [
  { mode: "fill", label: "Fill", none: "No fill" },
  { mode: "font", label: "Font" },
];

// dispatch.ts's sentence for any action with no connected Excel; the swatches
// run outside its switch, so they refuse the same way.
const NOT_CONNECTED = "Excel is not connected.";
const NO_FILLS = "No fills in the selection.";
const LOGO_NEEDS_TWO = "Upload a logo with at least two colors first.";

const noop = (): void => undefined;
let hooks: BrandColorsHooks = { preview: noop, commit: noop };
let logoColors: string[] = [];

async function pickSwatch(mode: string, hex: string | null): Promise<string> {
  if (!isExcelReady()) throw new Error(NOT_CONNECTED);
  if (mode !== "fill" && mode !== "font") {
    throw new Error(`Brand color: unknown swatch ${mode}.`);
  }
  return paintSwatch(mode, hex);
}

/**
 * Draws the swatch strips and the palette editor from the active settings.
 * The Brand tab calls it from every render, so it also runs when a palette is
 * adopted from the workbook at boot.
 */
export function renderBrandColors(): void {
  const settings = getActiveSettings();
  renderSwatches(getElement("brand-swatches"), {
    colors: brandColors(settings),
    strips: EXCEL_SWATCH_STRIPS,
    onPick: (mode, hex) => {
      void guard(() => pickSwatch(mode, hex), `swatch-${mode}`);
    },
  });
  renderPaletteEditor(getElement("brand-colors-editor"), {
    settings,
    onPreview: (next) => hooks.preview(next),
    onCommit: (next) => hooks.commit(next),
  });
}

// Edit colors: the Brand tab, scrolled to the palette, the first picker in focus.
function openPalette(): void {
  getElement<HTMLButtonElement>("tab-brand").click();
  getElement("palette-heading").scrollIntoView?.({ block: "start" });
  document.querySelector<HTMLElement>('[data-slot-color="primary"]')?.focus();
}

// Under the same rule as a paste: two or more replace the list, one is added.
async function readFromSelection(): Promise<string> {
  if (!isExcelReady()) throw new Error(NOT_CONNECTED);
  const colors = await readSelectionFills();
  return applyBrandColors(
    getElement("brand-colors-editor"),
    colors,
    0,
    NO_FILLS,
  );
}

/**
 * The colors a logo was read as, for Use these colors: the button shows for two
 * or more and hides for none (a reset, a logo with nothing in it).
 */
export function showLogoColors(colors: readonly string[]): void {
  logoColors = colors.slice(0, MAX_BRAND_COLORS);
  getElement("use-logo-colors").hidden = logoColors.length < 2;
}

function useLogoColors(): void {
  const next = withBrandList(getActiveSettings(), logoColors);
  if (!next) {
    toast.show(LOGO_NEEDS_TWO, "error");
    return;
  }
  hooks.commit(next);
  toast.show(`${String(next.brand.length)} brand colors set from your logo`);
}

/** Wires the three buttons and takes the Brand tab's save path. */
export function wireBrandColors(deps: BrandColorsHooks): void {
  hooks = deps;
  getElement("edit-brand-colors").addEventListener("click", openPalette);
  getElement("read-selection-colors").addEventListener("click", () => {
    void guard(readFromSelection, "read-selection-colors");
  });
  getElement("use-logo-colors").addEventListener("click", useLogoColors);
}
