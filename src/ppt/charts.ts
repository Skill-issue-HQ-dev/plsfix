// Chart links on a slide: what a payload's chart data would take to draw,
// whether this host has the shape API for the kind, the insert that lands it
// as a tagged group of native shapes, the repaint that rebuilds it where the
// user left it, and the fall back to a picture when the source stopped being
// drawable. Invariant: a link drawn as a group refreshes as a group, and only
// as one; a budget, a size or a stalled host never makes a picture.

import { chartDecline, kindApi, kindWords } from "../chart-kinds";
import {
  chartSize,
  layoutChart,
  MIN_SIZE,
  type Primitive,
} from "../chart-shapes";
import type { Box, Size } from "../layout";
import {
  INSERT_STOPPED,
  needsNewerPowerPoint,
  NO_SHAPE_API,
  PICTURE_CHOSEN,
  pictureNote,
  REFRESH_STOPPED,
  drawnNote,
  type ChartData,
} from "../link/chart-model";
import {
  sourceLabel,
  type InboxItem,
  type LinkTag,
  type Payload,
  type PicturePayload,
} from "../link/model";
import { base64ToBytes, pngSize } from "../link/png";
import { boxAtFound, chartBoxAt } from "./chart-box";
import { addPicture, pictureSynced } from "./chart-picture";
import { decideDrawing, populationOf } from "./chart-decide";
import { drawGroup, isDrawTimeout, withSyncDeadline } from "./chart-draw";
import { rideDeckCopy } from "./deck-flow";
import { drawProgress } from "./draw-progress";
import type { FoundLink, InsertResult } from "./host";
import {
  CONTENT_WIDTH,
  DEFAULT_TARGET,
  finishTarget,
  resolveTarget,
  SLIDE,
  type InsertTarget,
} from "./placement";
import { hasPowerPointApi, isGrouped } from "./shapes";

export { SHAPES_PER_SYNC } from "./chart-draw";

// Shape groups, text boxes and the formats a chart wears arrived in
// PowerPointApi 1.8; what a kind needs beyond that (a pie's start and end
// angle, an area's triangle tips and half turns, 1.10) is its rule in
// chart-kinds.ts.
const CHART_API = "1.8";

// A picture's pixels are 96 to the inch and a slide's points are 72.
const PNG_TO_POINTS = 0.75;
// The same margin CONTENT_WIDTH keeps, on the other side: a chart taller than
// this is scaled down to it, exactly as fitToSlide scales the picture, or the
// group hangs off the top and the bottom of the slide.
const CONTENT_HEIGHT = SLIDE.height - (SLIDE.width - CONTENT_WIDTH);

// The slide has the last word on the box: a busy one makes placeInFreeSpace
// shrink an object a tenth at a time rather than overlap what is already
// there, and 0.4 of a small chart is a smudge. This is the floor that scan may
// not go under - past it the chart is centred at full size over what is there
// and the pane says so (overlapping), never turned into a picture.
function minPlacementScale(size: Size): number {
  return Math.min(
    1,
    Math.max(MIN_SIZE.width / size.width, MIN_SIZE.height / size.height),
  );
}

// What a chart payload would become on a slide: the data, the size the picture
// beside it says the chart is, and the primitives that fill a box that size.
export interface ChartPlan {
  data: ChartData;
  size: Size;
  primitives: Primitive[];
  // The picture that travelled beside the data: what goes on the slide when
  // the user answers "Insert as picture" to the long-draw question. Empty for
  // a payload kept without picture bytes (a re-insert from the deck's copy),
  // which can then only be drawn or cancelled.
  png: string;
  // Set when the layout itself threw for this chart: the reason, after "as a
  // picture". A plan never throws, so a chart that cannot be laid out is a
  // picture with a reason and never an Update all that rejects for every row.
  refused?: string;
}

// chartSize caps the width alone, because how tall a chart may be is the
// slide's business and not the layout's: a chart the sheet made taller than
// the slide is scaled back, aspect kept, before anything is laid out in it.
// A very tall one scales to a sliver, and the layout is only drawn for boxes
// from MIN_SIZE up (below it the bars are hairlines and a label's box goes
// negative), so the box is lifted back to the minimum: a narrow chart is
// drawn wider than its source, never turned into a picture.
function onSlide(size: Size): Size {
  const scale = Math.min(1, CONTENT_HEIGHT / size.height);
  return {
    width: Math.max(size.width * scale, MIN_SIZE.width),
    height: Math.max(size.height * scale, MIN_SIZE.height),
  };
}

// The layout of a plan, or - when the layout throws - no primitives and the
// reason. The kind rules decline what they know cannot be drawn (chart-kinds.ts);
// this is the net under them, for the chart nobody foresaw.
function laidOut(
  data: ChartData,
  size: Size,
): Pick<ChartPlan, "primitives" | "refused"> {
  try {
    return { primitives: layoutChart(data, { left: 0, top: 0, ...size }) };
  } catch (error) {
    const why = error instanceof Error ? error.message : String(error);
    return {
      primitives: [],
      refused: `the chart could not be laid out (${why})`,
    };
  }
}

// Null for every payload with no chart data at all: a table, a plain range,
// and a chart type Excel could not describe, which are pictures and stay so.
export function chartPlan(payload: Payload): ChartPlan | null {
  if (payload.kind !== "picture" || payload.chart === undefined) return null;
  const data = payload.chart;
  const size = onSlide(
    chartSize(
      {
        width: payload.width * PNG_TO_POINTS,
        height: payload.height * PNG_TO_POINTS,
      },
      CONTENT_WIDTH,
    ),
  );
  return { data, size, ...laidOut(data, size), png: payload.png };
}

// Null means draw it. Anything else is the sentence the pane shows beside the
// picture it inserted instead, and every one of them names the picture first.
// Only what this host cannot do declines: no shape API at all, a kind that
// needs the newer one, a kind with no shape version yet (or a combo with a
// series that has none). How many shapes the chart is, how big the box is and
// how slow the host is are not reasons.
export function declineReason(plan: ChartPlan): string | null {
  if (plan.refused !== undefined) return pictureNote(plan.refused);
  return declineReasonFor(plan.data);
}

// The same answer from the chart's data alone: it depends on the chart (its
// kind, and for a combo its series) and the host, on nothing else in the plan,
// so a caller that only asks (the deck's records) need not lay a chart out.
export function declineReasonFor(data: ChartData): string | null {
  // A chart with no layout yet is a picture on every host: no upgrade helps.
  const unshaped = chartDecline(data);
  if (unshaped !== null) return pictureNote(unshaped);
  if (!hasPowerPointApi(CHART_API)) return NO_SHAPE_API;
  const needs = kindApi(data.kind);
  return needs === undefined || hasPowerPointApi(needs)
    ? null
    : needsNewerPowerPoint(kindWords(data.kind));
}

// The layout at the box the group actually gets: the plan's own primitives
// when the placement kept its size, laid out again when the slide had to
// shrink it or the user left the group at another width.
function primitivesAt(plan: ChartPlan, box: Box): Primitive[] {
  if (box.width === plan.size.width && box.height === plan.size.height) {
    return plan.primitives;
  }
  return layoutChart(plan.data, {
    left: 0,
    top: 0,
    width: box.width,
    height: box.height,
  });
}

function refreshStage(found: FoundLink): string {
  return `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
}

// Reinserting a group would drop it on the slide, not back into the group the
// user built, so a chart inside one is left alone and the row says why.
function requireUngrouped(found: FoundLink, stage: string): void {
  if (isGrouped(found)) {
    throw new Error(`${stage}: ungroup the chart before it can update`);
  }
}

// What one run hands back: the result the pane shows, and the empty layout
// placeholder the insert is to consume once the group is confirmed.
interface Drawn {
  result: InsertResult;
  consume?: string;
}

// The insert, as one PowerPoint run. A round trip the host swallowed
// (ChartDrawTimeout, whichever of the placement reads, the chunks or the
// grouping sync it was) has already had its shapes taken down again by
// drawGroup - by id, by the chart's name, by the link's token - and the Inbox
// row is only deleted after a successful insert (links.ts), so what is left
// is to say so. No picture takes the chart's place, however slow the host.
export async function insertChart(
  stage: string,
  item: InboxItem,
  plan: ChartPlan,
  tag: LinkTag,
  target: InsertTarget = DEFAULT_TARGET,
): Promise<InsertResult> {
  let drawn: Drawn;
  try {
    drawn = await PowerPoint.run((context) =>
      drawPlaced(context, stage, item, plan, tag, target),
    );
  } catch (error) {
    throw isDrawTimeout(error)
      ? new Error(INSERT_STOPPED, { cause: error })
      : error;
  }
  await finishTarget(target, drawn.result.slideId, drawn.consume);
  return drawn.result;
}

// The placement reads under the draw's deadline: the slide, the box, the ids
// already on it. The scale the free-space scan may not shrink the chart below
// is also the one a selected shape may not fit it at: past it the chart goes to
// free space and says why.
function placeChart(
  context: PowerPoint.RequestContext,
  stage: string,
  plan: ChartPlan,
  target: InsertTarget,
) {
  const floor = minPlacementScale(plan.size);
  return withSyncDeadline(
    resolveTarget(context, stage, target, plan.size, floor, floor),
  );
}

// One run: the placement, the decision, then the group - or, at his word, the
// picture.
async function drawPlaced(
  context: PowerPoint.RequestContext,
  stage: string,
  item: InboxItem,
  plan: ChartPlan,
  tag: LinkTag,
  target: InsertTarget,
): Promise<Drawn> {
  const at = await placeChart(context, stage, plan, target);
  const { slideId, placement, consume, before, spotNote } = at;
  const { box, overlapping, freeSpot, freeSpotSize } = placement;
  const primitives = primitivesAt(plan, box);
  const { choice, existing } = await decideDrawing(
    context,
    slideId,
    item,
    primitives.length,
    plan.png !== "",
    before.size,
  );
  const landed = { slideId, overlapping, freeSpot, freeSpotSize, spotNote };
  if (choice === "picture") {
    const shapeId = await pictureSynced(context, at, item, plan.png, tag);
    return { result: { ...landed, shapeId, note: PICTURE_CHOSEN }, consume };
  }
  const shapes = context.presentation.slides.getItem(slideId).shapes;
  const shapeId = await drawGroup(context, shapes, {
    primitives,
    box,
    font: plan.data.font,
    name: `pls,fix chart ${item.label}`,
    tag,
    token: item.token,
    slideId,
    existing,
    beforeIds: before,
    progress: drawProgress(item.label),
  });
  // A drawn chart says what it could not keep of its source (smooth lines, a
  // value under zero in a stack).
  const note = drawnNote(plan.data);
  const kept = note === undefined ? {} : { note };
  return { result: { ...landed, shapeId, ...kept }, consume };
}

// The chart drawn again where it sits. The new group is built first and the
// old one deleted in the same sync afterwards, the order tables.ts recreates
// in: a host that refuses the add never gets as far as taking the old group -
// and both its tags - down.
export async function refreshChart(
  found: FoundLink,
  plan: ChartPlan,
  tag: LinkTag,
): Promise<void> {
  const stage = refreshStage(found);
  requireUngrouped(found, stage);
  const box = chartBoxAt(found, plan.data, plan.size);
  const label = sourceLabel(found.tag.src, found.tag.kind);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(found.slideId).shapes;
    const old = shapes.getItem(found.shapeId);
    // The old group is still on the slide while the new one is drawn, so every
    // add pays for its shapes too (S1: 292 s for the 92-shape chart, against
    // 108 s with the old group gone).
    const existing = await populationOf(context, found.slideId, 0);
    await drawGroup(context, shapes, {
      primitives: primitivesAt(plan, box),
      box,
      font: plan.data.font,
      name: `pls,fix chart ${label}`,
      tag,
      token: found.token,
      slideId: found.slideId,
      existing,
      progress: drawProgress(label),
      before: () => {
        old.delete();
      },
    });
  });
}

// The modeller turned the source into a chart type no slide can draw, or this
// host lost the API that drew it: the picture goes where the group was, in one
// sync, with the add queued before the delete for the same reason as above. The
// record rides first: a host that cannot take it halts the batch before the
// picture is added, so the group stays the one tagged object and the repaint
// can run again.
export async function replaceGroupWithPicture(
  found: FoundLink,
  payload: PicturePayload,
  tag: LinkTag,
): Promise<void> {
  const stage = refreshStage(found);
  requireUngrouped(found, stage);
  const box = boxAtFound(found, pngSize(base64ToBytes(payload.png)));
  const label = sourceLabel(found.tag.src, found.tag.kind);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(found.slideId).shapes;
    const old = shapes.getItem(found.shapeId);
    rideDeckCopy(context, tag.id);
    addPicture(shapes, box, label, payload.png, { tag, token: found.token });
    old.delete();
    await withSyncDeadline(context.sync());
  });
}

// What "Update all" does to a link the deck holds as a group: draw the chart
// again when this host still can (the answer is what the redraw could not keep
// of the source, if anything), and put the picture there when it cannot (the
// answer is why).
export async function refreshChartGroup(
  found: FoundLink,
  payload: PicturePayload,
  tag: LinkTag,
): Promise<string | undefined> {
  const plan = chartPlan(payload);
  const reason = plan === null ? null : declineReason(plan);
  if (plan !== null && reason === null) {
    try {
      await refreshChart(found, plan, tag);
      return drawnNote(plan.data);
    } catch (error) {
      // A host that swallowed the redraw never took the old group down (it
      // goes in the sync that groups, which did not land), and drawGroup has
      // taken the new shapes down again: the chart is as it was, the row
      // fails with the sentence that says so, and the next Update all tries
      // again. A stalled host is never a reason for a picture.
      throw isDrawTimeout(error)
        ? new Error(REFRESH_STOPPED, { cause: error })
        : error;
    }
  }
  await replaceGroupWithPicture(found, payload, tag);
  // The reason travels back so "Update all" can say why a group is a picture:
  // the host that declined the draw has one, a source Excel could not
  // describe has its own (noteChartFallback in links.ts).
  return reason ?? undefined;
}
