// Reads brand colors off the sheet: the distinct fills of the selected cells in
// reading order, for Read from selection. Fills go through fill-store's readFill.
// Invariant: a read only (nothing written, no Undo slot), under the cell cap.

import { cappedAreas } from "./areas";
import { fillGrid, fillKeyColor, requestFills } from "./fill-store";
import { normalizeHex } from "../settings";

const STAGE = "Read from selection";

/**
 * The fill colors of the selection as #RRGGBB, each once, row by row across
 * every area (the first area first). Cells with no fill add nothing.
 */
export async function readSelectionFills(): Promise<string[]> {
  return Excel.run(async (context) => {
    const areas = await cappedAreas(context, STAGE);
    const grids = areas.map((area) => requestFills(area));
    await context.sync();

    const colors: string[] = [];
    for (const grid of grids) {
      for (const row of fillGrid(grid)) {
        for (const key of row) {
          const raw = fillKeyColor(key);
          const color = raw === null ? null : normalizeHex(raw);
          if (color && !colors.includes(color)) colors.push(color);
        }
      }
    }
    return colors;
  });
}
