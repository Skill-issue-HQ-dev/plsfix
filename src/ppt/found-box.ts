// The box a scanned shape occupies: its own, or - for a group the host reports
// with no area (PowerPoint for Mac) - the union of the shapes the scan found in
// it. Pure over the scan's flat list. Invariant: a link found inside such a
// group is rebuilt where the group stands, not at the slide's corner in no box.

import { hasArea, reportedBox, type Box } from "../layout";
import { boxOf } from "./shape-bounds";
import { GROUP_TYPE, type PlacedShape } from "./shapes";

// The scan lists a group's children with the path of ids that leads to them:
// the group's own path plus its own id.
function isChildOf(entry: PlacedShape, child: PlacedShape): boolean {
  const path = [...entry.groupPath, entry.shape.id];
  return (
    child.slideId === entry.slideId &&
    child.groupPath.length === path.length &&
    child.groupPath.every((id, index) => id === path[index])
  );
}

// `all` is every shape the scan listed, groups' children included (a group is
// opened three levels deep, so one nested deeper keeps whatever the host said).
export function scannedBox(
  entry: PlacedShape,
  all: readonly PlacedShape[],
): Box {
  const own = boxOf(entry.shape);
  if (hasArea(own) || entry.shape.type !== GROUP_TYPE) return own;
  const inside = all.filter((child) => isChildOf(entry, child));
  return reportedBox(
    own,
    inside.map((child) => scannedBox(child, all)),
  );
}
