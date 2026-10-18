// Every repaint of a link already in the deck: the tags a repaint writes, the
// one-round-trip batch route, the reinsertion fallback below PowerPointApi 1.8
// and the single host error a stale row earns. Split out of host.ts, which
// owns the scan, the insert and the break.
// Invariant: the row a button acts on is the last scan's, so the shape it
// names may already be gone - that is the one error reworded, never any other.
// Every repaint also keeps the deck's own copy of the link (deck-flow.ts).

import type { Size } from "../layout";
import {
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type LinkTag,
  type Payload,
  type PicturePayload,
} from "../link/model";
import { base64ToBytes, pngSize } from "../link/png";
import { aspectChanged } from "../link/status";
import { withSyncDeadline } from "./chart-draw";
import { refreshChartGroup } from "./charts";
import { joinNotes, refreshRecord } from "./deck-build";
import { rideDeckCopy, withDeckRecords } from "./deck-flow";
import type { FoundLink } from "./host";
import { missingShapeError } from "./missing-shape";
import { insertPictureBySelection } from "./picture";
import {
  GROUP_API,
  GROUP_TYPE,
  hasPowerPointApi,
  isGrouped,
  shapeAt,
} from "./shapes";
import { refreshTable } from "./tables";
import { refreshText } from "./texts";
import {
  declinedNote,
  refusedUpgrade,
  upgradeChart,
  upgradeOf,
  upgradeTable,
  wantsUpgrade,
  type Upgrade,
} from "./upgrade";

// What the host calls a native table (PowerPoint.ShapeType has no member typed here).
const TABLE_TYPE = "Table";

// fill.setImage arrived in PowerPointApi 1.8. An older host repaints by
// deleting the shape and inserting the picture again at the same box.
export function supportsInPlaceRefresh(): boolean {
  return hasPowerPointApi(GROUP_API);
}

// The tag a shape carries after this revision: the link's own identity plus
// where the payload came from. Shared with the insert in host.ts.
export function tagFor(
  link: {
    id: string;
    kind: LinkTag["kind"];
    project?: string;
    choice?: LinkTag["choice"];
  },
  payload: Payload,
  rev: number,
): LinkTag {
  return {
    v: 1,
    id: link.id,
    kind: link.kind,
    rev,
    src: payload.src,
    pushedAt: payload.pushedAt,
    ...(link.project ? { project: link.project } : {}),
    ...(link.choice ? { choice: link.choice } : {}),
  };
}

export async function writeTags(
  slideId: string,
  shapeId: string,
  tag: LinkTag,
  token: string,
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(slideId)
      .shapes.getItem(shapeId);
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, token);
    rideDeckCopy(context, tag.id);
    await withSyncDeadline(context.sync(), "tagging the picture");
  });
}

// The one geometry a refresh is allowed to write: a picture whose aspect ratio
// moved gets a height under the width the user chose. Left, top and width are
// the user's, and stay the user's.
function refreshedHeight(found: FoundLink, size: Size): number {
  return aspectChanged(found.width, found.height, size.width, size.height)
    ? Math.round(found.width * (size.height / size.width))
    : found.height;
}

function refreshStage(found: FoundLink): string {
  return `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
}

export interface RefreshRequest {
  found: FoundLink;
  payload: Payload;
  rev: number;
  // Set by an update (not a revert or a re-point): a picture link whose source
  // changed may become a native object on this paint (upgrade.ts).
  upgrade?: boolean;
}

interface PictureRequest extends RefreshRequest {
  payload: PicturePayload;
}

// A chart group repaints shape by shape, so it leaves the batch like a table,
// and so does a picture that becomes one, and a table asked for a picture (the
// row path says what is wrong).
function isPicture(entry: RefreshRequest): entry is PictureRequest {
  return (
    entry.payload.kind === "picture" &&
    entry.found.type !== GROUP_TYPE &&
    entry.found.type !== TABLE_TYPE &&
    !wantsUpgrade(entry)
  );
}

// One link repainted, whatever it is made of. The row a button acts on is the
// last scan's, so the shape it names may have been deleted or dragged out of
// its group since: that is the one host error this hands back in words. An
// update asks for `upgrade`: a picture link then becomes a native object, and a
// chart the host refuses to draw is painted as the picture it was.
export async function refreshLink(
  found: FoundLink,
  payload: Payload,
  rev: number,
  upgrade = false,
): Promise<string | undefined> {
  const to = upgrade ? upgradeOf(found, payload) : null;
  if (to === null) return paintLink(found, payload, rev, null);
  try {
    return await paintLink(found, payload, rev, to);
  } catch (error) {
    // A table has no picture to repaint, and a host that stopped answering is
    // the row's failure: only a refused chart draw falls back to the picture.
    if (to.to === "table" || !refusedUpgrade(error)) throw error;
    const note = await paintLink(found, payload, rev, null);
    return joinNotes(declinedNote(error), note);
  }
}

async function paintLink(
  found: FoundLink,
  payload: Payload,
  rev: number,
  to: Upgrade | null,
): Promise<string | undefined> {
  const kind = to?.to === "table" ? "table" : found.tag.kind;
  const entry = { record: refreshRecord(found, payload, rev, kind), payload };
  const { result, notes } = await withDeckRecords(
    [entry],
    () => repaintLink(found, payload, rev, to),
    true,
  ).catch((error: unknown) => {
    throw missingShapeError(refreshStage(found), error);
  });
  return joinNotes(result, notes.get(found.tag.id));
}

// The batch of one on a host with fill.setImage, and the reinsertion fallback
// below it. Answers with a note when the repaint has something to say about
// what it painted - why a chart group became a picture, or what a drawn chart
// could not keep of its source (smooth lines, a value under zero in a stack) -
// and with nothing when it has not.
async function repaintLink(
  found: FoundLink,
  payload: Payload,
  rev: number,
  to: Upgrade | null,
): Promise<string | undefined> {
  if (to?.to === "chart") {
    return upgradeChart(found, to, tagFor(found.tag, payload, rev));
  }
  if (to?.to === "table") {
    const tag = tagFor({ ...found.tag, kind: "table" }, payload, rev);
    return upgradeTable(found, to.payload, tag);
  }
  if (payload.kind === "table") {
    await refreshTable(found, payload, tagFor(found.tag, payload, rev));
    return undefined;
  }
  if (payload.kind === "text") {
    await refreshText(found, payload, tagFor(found.tag, payload, rev));
    return undefined;
  }
  if (found.type === GROUP_TYPE) {
    // charts.ts owns the reason; a version of it with none to give simply
    // answers nothing, and the summary then falls back to its own.
    const note: unknown = await refreshChartGroup(
      found,
      payload,
      tagFor(found.tag, payload, rev),
    );
    return typeof note === "string" ? note : undefined;
  }
  // A link that became a table cannot take a picture back (a revert to the
  // revision before it, a push made before it): say so, never hand the host a
  // fill for a table.
  if (found.type === TABLE_TYPE) {
    throw new Error(
      `${refreshStage(found)}: this link is a table now, so it cannot go back to a picture`,
    );
  }
  if (supportsInPlaceRefresh()) {
    await refreshLinks([{ found, payload, rev }]);
    return undefined;
  }
  const stage = refreshStage(found);
  // Reinsertion drops the picture on the slide, not back into its group, so
  // a grouped link is left alone and the row says why.
  if (isGrouped(found)) {
    throw new Error(
      `${stage}: grouped pictures need PowerPoint 2504/16.96 or newer`,
    );
  }
  const size = pngSize(base64ToBytes(payload.png));
  const tag = tagFor(found.tag, payload, rev);
  await reinsertLink(
    stage,
    found,
    payload.png,
    tag,
    refreshedHeight(found, size),
  );
  return undefined;
}

// Every in-place repaint of an "Update all" in one round trip: the pictures,
// the tags and the heights are queued for the whole batch and sent with a
// single sync, because a sync per shape is what makes a sixty-link deck crawl.
// False when the host has no fill.setImage, so the caller repaints row by row
// through the reinsertion above. A host that refuses one shape rejects the
// whole batch; retrying those rows one at a time is safe, because a repaint
// writes the same picture, tag and height however often it runs.
export async function refreshLinks(
  batch: RefreshRequest[],
  onNote?: (found: FoundLink, note: string) => void,
): Promise<boolean> {
  if (!supportsInPlaceRefresh()) return false;
  // A table is written cell by cell, not with one setImage: a batch holding
  // one goes back to the caller, which replays every row on its own.
  if (!batch.every(isPicture)) return false;
  if (batch.length === 0) return true;
  const entries = batch.map((one) => ({
    record: refreshRecord(one.found, one.payload, one.rev),
    payload: one.payload,
  }));
  // The records ride this batch. One the host refuses is run again row by row
  // (links.ts paintBatch), and by then the deck rides no more (deck-flow.ts).
  const { notes } = await withDeckRecords(entries, () => repaintBatch(batch));
  for (const { found } of batch) {
    const note = notes.get(found.tag.id);
    if (note !== undefined) onNote?.(found, note);
  }
  return true;
}

function repaintBatch(batch: PictureRequest[]): Promise<void> {
  return PowerPoint.run(async (context) => {
    for (const entry of batch) queueRefresh(context, entry);
    await withSyncDeadline(context.sync(), "repainting the links");
  });
}

// One picture repainted where it sits. Nothing here reads a shape property, so
// no entry in the batch needs a load: the geometry it compares against is the
// one the scan already read.
function queueRefresh(
  context: PowerPoint.RequestContext,
  { found, payload, rev }: PictureRequest,
): void {
  const height = refreshedHeight(found, pngSize(base64ToBytes(payload.png)));
  const shape = shapeAt(context, found);
  shape.fill.setImage(payload.png);
  if (height !== found.height) shape.height = height;
  shape.tags.add(TAG_LINK, encodeTag(tagFor(found.tag, payload, rev)));
  rideDeckCopy(context, found.tag.id);
}

// The fallback repaint, in the only safe order: the new picture lands on the
// same slide at the same box and is tagged first, and the old shape goes last.
// A failure before the delete leaves a duplicate the user can remove; a delete
// first would lose the picture and both tags for good. The new shape is still
// the last one on the slide while the old one is there, so it is found the
// same way as on a first insert.
async function reinsertLink(
  stage: string,
  found: FoundLink,
  png: string,
  tag: LinkTag,
  height: number,
): Promise<void> {
  const box = { left: found.left, top: found.top, width: found.width, height };
  const shapeId = await insertPictureBySelection(
    stage,
    found.slideId,
    png,
    box,
  );
  await writeTags(found.slideId, shapeId, tag, found.token);
  await PowerPoint.run(async (context) => {
    context.presentation.slides
      .getItem(found.slideId)
      .shapes.getItem(found.shapeId)
      .delete();
    await withSyncDeadline(context.sync(), "removing the old picture");
  });
}
