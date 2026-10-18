// One-click brand swatches: a brand color as the fill or the font color of every
// selected area, or no fill. The cycles' write path minus the active-cell read.
// Invariant: never paints over an overlay's tint (the Paintbrush apply's own
// refusal); one write sync.

import { selectedAreas } from "./areas";
import { requireNoOverlayOwner } from "./fill-store";
import { protectedNote, syncWrite } from "./protection";
import { captureUndoAreas } from "./undo";
import { normalizeHex } from "../settings";

export type SwatchPaintMode = "fill" | "font";

const STAGE = "Brand color";

function receipt(mode: SwatchPaintMode, color: string | null): string {
  if (color === null) return "No fill";
  return `${mode === "fill" ? "Fill" : "Font"} ${color}`;
}

/**
 * Paints `hex` onto the selection as its fill or its font color; a null color
 * on `fill` clears the fill. Answers the receipt the pane toasts: "Fill
 * #B27E54", "Font #B27E54" or "No fill".
 */
export async function paintSwatch(
  mode: SwatchPaintMode,
  hex: string | null,
): Promise<string> {
  const color = hex === null ? null : normalizeHex(hex);
  if (hex !== null && color === null) {
    throw new Error(`${STAGE}: ${hex} is not a color.`);
  }
  if (mode === "font" && color === null) {
    throw new Error(`${STAGE}: a font color needs a color.`);
  }
  // Painted over an overlay's tint, Undo would capture the tint as the
  // modeller's own fill and hand it back later in the overlay's place.
  requireNoOverlayOwner(STAGE);

  return Excel.run(async (context) => {
    const areas = await selectedAreas(context, STAGE);
    await captureUndoAreas(context, areas);
    for (const area of areas) {
      if (mode === "font") area.format.font.color = color!;
      else if (color === null) area.format.fill.clear();
      else area.format.fill.color = color;
    }
    await syncWrite(context, STAGE, protectedNote, areas.length);
    return receipt(mode, color);
  });
}
