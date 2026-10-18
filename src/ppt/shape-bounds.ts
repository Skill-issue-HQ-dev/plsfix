// A shape's box as the host reports it, and a group's read off its children
// when the host reports none (PowerPoint for Mac, 13.09). Needs the shape
// loaded with SHAPE_PROPERTIES. Invariant: a group is never taken as a box of
// no size while its children have one.

import { hasArea, reportedBox, type Box } from "../layout";
import { withSyncDeadline } from "./chart-draw";
import {
  GROUP_API,
  GROUP_TYPE,
  hasPowerPointApi,
  SHAPE_PROPERTIES,
} from "./shapes";

export function boxOf(shape: PowerPoint.Shape): Box {
  return {
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

export async function boundsOf(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
): Promise<Box> {
  const own = boxOf(shape);
  if (
    hasArea(own) ||
    shape.type !== GROUP_TYPE ||
    !hasPowerPointApi(GROUP_API)
  ) {
    return own;
  }
  const inner = shape.group.shapes;
  inner.load(SHAPE_PROPERTIES);
  await withSyncDeadline(context.sync(), "reading a group's shapes");
  const children: Box[] = [];
  for (const child of inner.items) {
    children.push(await boundsOf(context, child));
  }
  return reportedBox(own, children);
}
