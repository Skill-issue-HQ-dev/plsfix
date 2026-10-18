// The linked picture of a chart link, in the only two places a picture is
// still the right shape for one: the user's own "Insert as picture" at the
// long-draw question, and the group a host or a source can no longer draw
// becoming one on Update all. One tagged rectangle, image filled. A budget, a
// size or a stalled host never lands one (the timeout paths in charts.ts
// throw instead). Split out of charts.ts at its 400-line cap.

import type { Box } from "../layout";
import {
  encodeTag,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
} from "../link/model";
import { cleanupByToken } from "./chart-cleanup";
import { withSyncDeadline } from "./chart-draw";
import { rideDeckCopy } from "./deck-flow";
import type { ResolvedTarget } from "./placement";

// The linked picture as one shape: the rectangle every picture lands, image
// filled and carrying both tags, so it is still a link the deck updates.
export function addPicture(
  shapes: PowerPoint.ShapeCollection,
  box: Box,
  label: string,
  png: string,
  identity: { tag: LinkTag; token: string },
): PowerPoint.Shape {
  const shape = shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    box,
  );
  shape.name = `pls,fix link ${label}`;
  shape.lineFormat.visible = false;
  shape.fill.setImage(png);
  shape.tags.add(TAG_LINK, encodeTag(identity.tag));
  shape.tags.add(TAG_KEY, identity.token);
  return shape;
}

// The picture added, its id read back and the batch committed under the
// deadline in the caller's run: the user's choice at the question lands it
// here, at the box the chart was placed in on the slide the placement chose.
export async function pictureSynced(
  context: PowerPoint.RequestContext,
  at: ResolvedTarget,
  item: InboxItem,
  png: string,
  tag: LinkTag,
): Promise<string> {
  const { slideId, before } = at;
  const shapes = context.presentation.slides.getItem(slideId).shapes;
  // The tag remembers the choice, so no update ever turns this picture into
  // shapes (upgrade.ts).
  const chosen: LinkTag = { ...tag, choice: "picture" };
  const shape = addPicture(shapes, at.placement.box, item.label, png, {
    tag: chosen,
    token: item.token,
  });
  rideDeckCopy(context, tag.id);
  shape.load("id");
  // A host that applies the add and both tags before answering the sync
  // itself with an error leaves a fully tagged, orphaned picture behind:
  // its id was never read back, so a cleanup by id has nothing to go on.
  // cleanupByToken finds it by the token instead, sparing every id already
  // in `before` - the token is the link's, and an older copy shares it.
  try {
    await withSyncDeadline(context.sync());
  } catch (error) {
    await withSyncDeadline(cleanupByToken(slideId, item.token, before)).catch(
      () => undefined,
    );
    throw error;
  }
  return shape.id;
}
