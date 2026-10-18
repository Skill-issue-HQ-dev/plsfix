// The PowerPoint Office.js code every link flow goes through: scan by tags,
// insert by payload kind, re-point, break, show a slide (repaints: refresh.ts,
// re-exported here). Identity is the PLSFIX_LINK tag, never a shape id; one
// sync per batch, never per shape; deck-flow.ts's hooks keep the deck's copy.

import {
  decodeTag,
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type Payload,
  type PicturePayload,
} from "../link/model";
import { pictureNote } from "../link/chart-model";
import { base64ToBytes, pngSize } from "../link/png";
import { fitToSlide, type Box } from "../link/status";
import { cleanupByToken } from "./chart-cleanup";
import { withSyncDeadline } from "./chart-draw";
import { chartPlan, declineReason, insertChart } from "./charts";
import { deckBreak, deckInsert, deckScan, rideDeckCopy } from "./deck-flow";
import { scannedBox } from "./found-box";
import type { InsertResult } from "./insert-result";
import { missingShapeError } from "./missing-shape";
import { insertPictureBySelection } from "./picture";
import { supportsInPlaceRefresh, tagFor, writeTags } from "./refresh";
import {
  DEFAULT_TARGET,
  finishTarget,
  readSelectedSlideId,
  resolveTarget,
  type InsertTarget,
} from "./placement";
import {
  expandGroups,
  shapeAt,
  SHAPE_PROPERTIES,
  type PlacedShape,
  type ShapePath,
} from "./shapes";
import { insertTable } from "./tables";
import { insertText } from "./texts";

// The repaint half of the adapter, kept on this module's surface so no caller
// has to know it moved (src/ppt/links.ts builds its PptHost from `typeof
// realHost`, and revert.ts and fetch.ts take RefreshRequest from here).
export {
  refreshLink,
  refreshLinks,
  supportsInPlaceRefresh,
  type RefreshRequest,
} from "./refresh";
// What an insert hands back and the overlap sentence (insert-result.ts).
export { OVERLAP_NOTE, overlapNote, type InsertResult } from "./insert-result";

export interface FoundLink extends ShapePath {
  slideIndex: number;
  tag: LinkTag;
  token: string;
  // What the deck holds this link as: "Group" is a native chart, anything
  // else the picture. An update may upgrade a picture (upgrade.ts).
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

const TAG_PROPERTIES = "items/key,items/value";

interface TaggedShape extends PlacedShape {
  tags: PowerPoint.TagCollection;
}

// Three round trips, and office.js allows no fewer: the slides, then every
// slide's shapes, then every shape's tags - the only place identity is read.
// Groups add one sync per nesting level, between the second and the third.
async function scanUnrecorded(
  report: (slides: number) => void,
): Promise<FoundLink[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await withSyncDeadline(context.sync(), "reading the links");
    report(slides.items.length);
    const sets = slides.items.map((slide, slideIndex) => {
      const shapes = slide.shapes;
      shapes.load(SHAPE_PROPERTIES);
      return { slideId: slide.id, slideIndex, shapes };
    });
    await withSyncDeadline(context.sync(), "reading the links");
    const placed = sets.flatMap((set) =>
      set.shapes.items.map((shape): PlacedShape => ({
        slideId: set.slideId,
        slideIndex: set.slideIndex,
        shape,
        groupPath: [],
      })),
    );
    const tagged = (await expandGroups(context, placed)).map(
      (entry): TaggedShape => {
        const tags = entry.shape.tags;
        tags.load(TAG_PROPERTIES);
        return { ...entry, tags };
      },
    );
    await withSyncDeadline(context.sync(), "reading the links");
    return tagged
      .map((entry) => toFoundLink(entry, tagged))
      .filter((link): link is FoundLink => link !== null);
  });
}
export const scanLinks = deckScan(scanUnrecorded);

// A shape is a link only when it carries both tags: the identity and the key
// that opens its payload. Anything else on the slide is somebody's picture. Its
// box is the host's own (found-box.ts reads a Mac group's off its children).
function toFoundLink(entry: TaggedShape, all: PlacedShape[]): FoundLink | null {
  const items = entry.tags.items;
  const tag = decodeTag(items.find((item) => item.key === TAG_LINK)?.value);
  const token = items.find((item) => item.key === TAG_KEY)?.value;
  if (tag === null || token === undefined || token === "") return null;
  const shape = entry.shape;
  return {
    slideId: entry.slideId,
    slideIndex: entry.slideIndex,
    shapeId: shape.id,
    groupPath: entry.groupPath.length > 0 ? entry.groupPath : undefined,
    tag,
    token,
    type: shape.type,
    ...scannedBox(entry, all),
  };
}

// What "Update this slide" acts on: PowerPoint's own selection, not a tick in
// the pane - the pane cannot see the selection any other way. Null when
// nothing is selected, so the caller can say so instead of guessing a slide.
export async function activeSlideId(): Promise<string | null> {
  return PowerPoint.run((context) => readSelectedSlideId(context));
}

async function insertUnrecorded(
  item: InboxItem,
  payload: Payload,
  rev: number,
  target: InsertTarget = DEFAULT_TARGET,
): Promise<InsertResult> {
  const stage = `insert ${item.label}`;
  const tag = tagFor(item, payload, rev);
  if (payload.kind === "table") {
    return insertTable(stage, item, payload, tag, target);
  }
  if (payload.kind === "text") {
    return insertText(stage, item, payload, tag, target);
  }
  // A chart this host can draw lands as shapes; the rest take the picture,
  // with the reason when there is one.
  const plan = chartPlan(payload);
  const note =
    plan === null ? issueNote(payload) : (declineReason(plan) ?? undefined);
  if (plan !== null && note === undefined) {
    return insertChart(stage, item, plan, tag, target);
  }
  return insertPictureLink(stage, item, payload, tag, note, target);
}
export const insertLink = deckInsert(insertUnrecorded);

// How far a plain picture link may shrink before the free-space search gives
// up and overlaps instead: half its natural size. A picture (unlike a chart)
// has no MIN_SIZE of its own below which it stops being readable, but half
// size is still a real image and a better answer than a postage stamp.
const PICTURE_MIN_SCALE = 0.5;

// The two ways an un-drawable chart or a plain picture lands: the selection
// API below PowerPointApi 1.8, or an in-place rectangle with the PNG as its
// fill (insertPictureInPlace, its own function to stay under the line cap)
// on 1.8 and above.
async function insertPictureLink(
  stage: string,
  item: InboxItem,
  payload: PicturePayload,
  tag: LinkTag,
  note: string | undefined,
  target: InsertTarget,
): Promise<InsertResult> {
  const size = pngSize(base64ToBytes(payload.png));
  const fitted = fitToSlide(size.width, size.height);
  if (!supportsInPlaceRefresh()) {
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, stage, target, fitted, PICTURE_MIN_SCALE),
    );
    const { slideId, placement, consume } = resolved;
    const shapeId = await insertPictureBySelection(
      stage,
      slideId,
      payload.png,
      placement.box,
    );
    await writeTags(slideId, shapeId, tag, item.token);
    await finishTarget(target, slideId, consume);
    return {
      slideId,
      shapeId,
      overlapping: placement.overlapping,
      freeSpot: placement.freeSpot,
      freeSpotSize: placement.freeSpotSize,
      note,
    };
  }
  return insertPictureInPlace(stage, item, payload, tag, note, target, fitted);
}

async function insertPictureInPlace(
  stage: string,
  item: InboxItem,
  payload: PicturePayload,
  tag: LinkTag,
  note: string | undefined,
  target: InsertTarget,
  fitted: Box,
): Promise<InsertResult> {
  const placed = await PowerPoint.run(async (context) => {
    const resolved = await resolveTarget(
      context,
      stage,
      target,
      fitted,
      PICTURE_MIN_SCALE,
    );
    const { slideId, placement, consume, before } = resolved;
    const shape = context.presentation.slides
      .getItem(slideId)
      .shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        placement.box,
      );
    shape.name = `pls,fix link ${item.label}`;
    shape.lineFormat.visible = false;
    shape.fill.setImage(payload.png);
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, item.token);
    rideDeckCopy(context, tag.id);
    shape.load("id");
    // The picture's id is only known once this sync answers, so a host
    // that swallows it (rolls the batch back) leaves nothing to delete -
    // but one that applies the add and both tags before answering the
    // sync itself with an error leaves a fully tagged, orphaned picture;
    // this catch is what finds and removes it. cleanupByToken matches by
    // the token, which the link owns, not this one push - `before` (read
    // by resolveTarget ahead of the add) is what keeps it off an older,
    // already-finished copy of the same link.
    try {
      await withSyncDeadline(context.sync(), "inserting the picture");
    } catch (error) {
      await withSyncDeadline(cleanupByToken(slideId, item.token, before)).catch(
        () => undefined,
      );
      throw error;
    }
    return {
      slideId,
      shapeId: shape.id,
      overlapping: placement.overlapping,
      freeSpot: placement.freeSpot,
      freeSpotSize: placement.freeSpotSize,
      consume,
    };
  });
  await finishTarget(target, placed.slideId, placed.consume);
  return {
    slideId: placed.slideId,
    shapeId: placed.shapeId,
    overlapping: placed.overlapping,
    freeSpot: placed.freeSpot,
    freeSpotSize: placed.freeSpotSize,
    note,
  };
}

// Why Excel shipped a chart link with no chart data: the picture is certain,
// and the pane says so the way it says a decline of this host's own.
export function issueNote(payload: PicturePayload): string | undefined {
  return payload.chartIssue === undefined
    ? undefined
    : pictureNote(payload.chartIssue);
}

// Re-pointing a link at another export: both tags rewritten in one sync, and
// nothing else - no picture, no geometry, no z-order. The shape keeps
// everything the user gave it and only changes what it tracks.
export async function retagLink(
  found: FoundLink,
  tag: LinkTag,
  token: string,
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const tags = shapeAt(context, found).tags;
    tags.add(TAG_LINK, encodeTag(tag));
    tags.add(TAG_KEY, token);
    await withSyncDeadline(context.sync(), "retagging the link");
  });
}

// Breaking a link leaves the picture exactly where it is; only the identity
// and the key go, so nothing in the deck ever refreshes it again.
async function breakUnrecorded(found: FoundLink): Promise<void> {
  await PowerPoint.run(async (context) => {
    rideDeckCopy(context, found.tag.id);
    const tags = shapeAt(context, found).tags;
    tags.delete(TAG_LINK);
    tags.delete(TAG_KEY);
    await withSyncDeadline(context.sync(), "breaking the link");
  }).catch((error: unknown) => {
    throw missingShapeError(
      `break ${sourceLabel(found.tag.src, found.tag.kind)}`,
      error,
    );
  });
}
export const breakLink = deckBreak(breakUnrecorded);

export async function goToSlide(slideId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await withSyncDeadline(context.sync(), "going to the slide");
  });
}

// How many slides the deck holds right now: what the inbox's Slide picker
// renders as "Slide 1".."Slide N" (target.ts).
export async function slideCount(): Promise<number> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await withSyncDeadline(context.sync(), "reading the slides");
    return slides.items.length;
  });
}

// The id of the slide at that position (0-based), so the picker's "Slide N"
// option - which only ever names a position, never an id - resolves to a
// real slide right before an insert reads it. Read alongside every slide's
// id in the one sync, so a deck that lost a slide since the picker was last
// rendered answers the pane's own sentence, never a raw office.js one.
export async function slideIdAt(index: number): Promise<string> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await withSyncDeadline(context.sync(), "reading the slides");
    const slide = slides.items[index];
    if (!slide) {
      throw new Error(
        `Slide ${String(index + 1)} is gone: pick a slide again.`,
      );
    }
    return slide.id;
  });
}
