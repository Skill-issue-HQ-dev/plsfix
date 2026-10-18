// Picture links that become native objects on the paint that was coming anyway:
// what a picture link may become (a chart group, a table), why one stays a
// picture, and the table build. The chart draw is charts.ts's own refreshChart.
// Invariant: the old picture is deleted only in the batch that tags the new
// object, and a picture the user chose is never turned into anything.

import { drawnNote, labelledNote, pictureNote } from "../link/chart-model";
import {
  encodeTag,
  overTableCap,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  TAG_PAINT,
  type LinkTag,
  type Payload,
  type PicturePayload,
  type TablePayload,
} from "../link/model";
import { encodePaintMap } from "../link/paint-map";
import { cleanupByToken, cleanupShapes } from "./chart-cleanup";
import { isDrawTimeout, withSyncDeadline } from "./chart-draw";
import {
  chartPlan,
  declineReason,
  declineReasonFor,
  refreshChart,
  type ChartPlan,
} from "./charts";
import { outcomeOf } from "./deck-build";
import { rideDeckCopyOf } from "./deck-flow";
import type { FoundLink } from "./host";
import type { RefreshRequest } from "./refresh";
import { awaitShapeRegistered } from "./shape-ready";
import { GROUP_TYPE, hasPowerPointApi, isGrouped, shapeAt } from "./shapes";
import {
  addTable,
  requireTableApi,
  requireTableSize,
  TABLE_API,
  tableSize,
} from "./tables";
import { FORMATTING, NO_CLEARS, writeCellsInChunks } from "./table-cells";
import { uniformSize } from "./table-font";
import { queueTableStyle } from "./table-style";

export const NOW_CHART = "now a native chart";
export const NOW_TABLE = "now a table";
export const UPGRADE_STOPPED =
  "PowerPoint stopped answering; the picture stays";
export const UNGROUP_FOR_CHART =
  "stays a picture: it sits inside a group. Ungroup it and the next update draws a chart";

const BUILDING_TABLE = "building the table";
const TAKING_DOWN = "taking the new table down";

// What a picture link becomes, with what the draw needs.
export type Upgrade =
  | { to: "chart"; payload: PicturePayload; plan: ChartPlan }
  | { to: "table"; payload: TablePayload };

// A picture the user chose, at the long-draw question: the tag remembers it.
export function isChosen(found: FoundLink): boolean {
  return found.tag.choice === "picture";
}

// A chart link whose picture nobody chose: it is a picture because a host or a
// source could not draw it then.
function isPictureEraChart(found: FoundLink): boolean {
  return found.tag.kind === "chart" && !isChosen(found);
}

// What this picture link becomes when it is painted with this payload, or null
// when it stays a picture: a chosen one, a picture inside a user's group, a
// payload this host cannot draw, an explicit picture (a range link's payload is
// a picture until the user turns it into a table).
export function upgradeOf(found: FoundLink, payload: Payload): Upgrade | null {
  if (found.type === GROUP_TYPE || isGrouped(found)) return null;
  if (found.tag.kind === "range" && payload.kind === "table") {
    const drawable =
      hasPowerPointApi(TABLE_API) && !overTableCap(payload.rows, payload.cols);
    return drawable ? { to: "table", payload } : null;
  }
  if (
    !isPictureEraChart(found) ||
    payload.kind !== "picture" ||
    payload.chart === undefined
  ) {
    return null;
  }
  // The host's answer needs the data alone; the layout is only worth drawing
  // once the chart is going to be.
  if (declineReasonFor(payload.chart) !== null) return null;
  const plan = chartPlan(payload);
  return plan !== null && plan.refused === undefined
    ? { to: "chart", payload, plan }
    : null;
}

// A request an update made (not a revert or a re-point) that will upgrade.
export function wantsUpgrade(entry: RefreshRequest): boolean {
  return (
    entry.upgrade === true && upgradeOf(entry.found, entry.payload) !== null
  );
}

// Each upgrade travels alone, so the pictures beside it keep the one round trip
// the batch route gives them.
export function isolateUpgrades(parts: RefreshRequest[][]): RefreshRequest[][] {
  return parts.flatMap((part) => {
    const alone = part.filter(wantsUpgrade);
    if (alone.length === 0) return [part];
    const rest = part.filter((entry) => !wantsUpgrade(entry));
    return [...(rest.length > 0 ? [rest] : []), ...alone.map((one) => [one])];
  });
}

// Why a chart link is a picture after this paint, or undefined when nothing
// declined it: the reason Excel sent no chart data, this host's own reason, a
// group around the picture. A picture the user chose says nothing.
export function whyNotNative(
  found: FoundLink,
  payload: Payload,
): string | undefined {
  if (payload.kind !== "picture") return undefined;
  const native = found.type === GROUP_TYPE;
  if (!native && !isPictureEraChart(found)) return undefined;
  const plan = chartPlan(payload);
  const issue =
    payload.chartIssue === undefined ? null : pictureNote(payload.chartIssue);
  const reason = plan === null ? issue : declineReason(plan);
  if (reason !== null) return reason;
  return !native && isGrouped(found) ? UNGROUP_FOR_CHART : undefined;
}

// A chart draw the host refused falls back to painting the picture; a host that
// stopped answering and a shape the deck no longer holds do not (the row fails).
export function refusedUpgrade(error: unknown): boolean {
  return outcomeOf(error) === "failed";
}

export function declinedNote(error: unknown): string {
  const reason = error instanceof Error ? error.message : String(error);
  return `stays a picture: PowerPoint would not draw it as a chart (${reason})`;
}

// The chart drawn where the picture sits (charts.ts refreshChart: the group is
// built first and the picture goes in the sync that groups it).
export async function upgradeChart(
  found: FoundLink,
  upgrade: Extract<Upgrade, { to: "chart" }>,
  tag: LinkTag,
): Promise<string> {
  try {
    await refreshChart(found, upgrade.plan, tag);
  } catch (error) {
    throw isDrawTimeout(error)
      ? new Error(UPGRADE_STOPPED, { cause: error })
      : error;
  }
  const kept = drawnNote(upgrade.plan.data);
  return kept === undefined ? NOW_CHART : labelledNote(NOW_CHART, kept);
}

// The table built beside the picture at its corner and width; tags, the deck's
// record and the picture's delete all wait for the last format round trip, so a
// refusal anywhere before it leaves the picture and its link as they were.
export async function upgradeTable(
  found: FoundLink,
  payload: TablePayload,
  tag: LinkTag,
): Promise<string> {
  requireTableApi();
  requireTableSize(payload);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(found.slideId).shapes;
    // Resolved before the add: a picture that is gone halts the batch before
    // anything is built beside it.
    const old = shapeAt(context, found);
    const built = addBuilt(shapes, found, payload);
    // Known once the add's sync has answered, and not before.
    let builtId: string | undefined;
    try {
      await withSyncDeadline(context.sync(), BUILDING_TABLE);
      builtId = built.id;
      await awaitShapeRegistered(context, shapes, builtId, BUILDING_TABLE);
      await formatAndSwap(context, old, built, found, payload, tag);
    } catch (error) {
      await takeDownBuilt(found, builtId);
      throw isDrawTimeout(error)
        ? new Error(UPGRADE_STOPPED, { cause: error })
        : error;
    }
  });
  return NOW_TABLE;
}

// The key a table wears while it is built beside the picture: the link's own
// token would also name every copy of the link, so it carries a suffix that
// only a half-built table has, and the swap writes the real token over it.
function buildKey(token: string): string {
  return `${token}.building`;
}

// The table added at the picture's corner and width, wearing its build key: a
// host that applies this batch and then refuses its sync leaves a table whose
// id was never read, and the key is how takeDownBuilt finds it.
function addBuilt(
  shapes: PowerPoint.ShapeCollection,
  found: FoundLink,
  payload: TablePayload,
): PowerPoint.Shape {
  const box = {
    left: found.left,
    top: found.top,
    width: found.width,
    height: tableSize(payload).height,
  };
  const built = addTable(
    shapes,
    payload,
    box,
    sourceLabel(payload.src, "table"),
  );
  built.tags.add(TAG_KEY, buildKey(found.token));
  built.load("id");
  return built;
}

// The formats, a few cells per round trip, and in the last one the swap: the
// record first (a host that refuses it halts the batch before the tags and the
// delete, and the picture is still the link), then the tags, the real token
// over the build key, and the picture's delete.
function formatAndSwap(
  context: PowerPoint.RequestContext,
  old: PowerPoint.Shape,
  built: PowerPoint.Shape,
  found: FoundLink,
  payload: TablePayload,
  tag: LinkTag,
): Promise<void> {
  queueTableStyle(built.getTable(), payload.h === true);
  return writeCellsInChunks(context, built.getTable(), payload, {
    withText: false,
    clear: NO_CLEARS,
    what: FORMATTING,
    uniformFontSize: uniformSize(payload),
    beforeLast: () => {
      rideDeckCopyOf(context, payload);
      built.tags.add(TAG_LINK, encodeTag(tag));
      built.tags.add(TAG_KEY, found.token);
      built.tags.add(TAG_PAINT, encodePaintMap(payload.cells));
      old.delete();
    },
  });
}

// The half-built table goes, unless the batch that swapped it for the picture
// landed after all (a host that applies a batch and then answers it with an
// error): then the table is the link, and it stays. By its id when the add's
// sync told it, otherwise by the key it was built with (cleanupByToken, the way
// a fresh insert whose id was never read back is found). Both run under the
// sync deadline: a cleanup that hangs would hold the pane busy as the build did.
async function takeDownBuilt(
  found: FoundLink,
  id: string | undefined,
): Promise<void> {
  try {
    const stillThere = await PowerPoint.run(async (context) => {
      const probe = context.presentation.slides
        .getItem(found.slideId)
        .shapes.getItemOrNullObject(found.shapeId);
      probe.load("isNullObject");
      await withSyncDeadline(context.sync(), TAKING_DOWN);
      return !probe.isNullObject;
    });
    if (!stillThere) return;
    await withSyncDeadline(
      id === undefined
        ? cleanupByToken(found.slideId, buildKey(found.token), new Set())
        : cleanupShapes(found.slideId, [id]),
      TAKING_DOWN,
    );
  } catch {
    // Best effort: the caller already has the error it means to surface.
  }
}
