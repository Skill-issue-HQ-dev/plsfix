// The box a chart link is drawn into again: a picture's own box, or for a group
// the layout box derived from the union of its shapes (chart-frame.ts).
// Invariant: a rebuild never moves the corner or changes the width the user can
// see; the height is the one geometry a refresh decides for itself.

import { frameFor } from "../chart-frame";
import type { Box, Size } from "../layout";
import type { ChartData } from "../link/chart-model";
import type { FoundLink } from "./host";
import { GROUP_TYPE } from "./shapes";

// The corner and width the user chose, and the only geometry a refresh decides
// for itself - the height the chart's own aspect asks for at that width.
export function boxAtFound(found: FoundLink, size: Size): Box {
  return {
    left: found.left,
    top: found.top,
    width: found.width,
    height: Math.round(found.width * (size.height / size.width)),
  };
}

// The layout box of a chart drawn again over `found`: a picture's own box, or
// for a group the box whose drawing stands where the group stands.
export function chartBoxAt(found: FoundLink, data: ChartData, size: Size): Box {
  if (found.type !== GROUP_TYPE) return boxAtFound(found, size);
  return frameFor(data, found, size.height / size.width, size.width);
}
