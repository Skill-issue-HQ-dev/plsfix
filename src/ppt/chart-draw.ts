// One chart layout turned into native PowerPoint shapes: a primitive per add,
// the adds chunked into syncs (draw-pace.ts sizes each chunk before its sync
// from the cost model, and a sync that came back slow halves what the rest may
// carry), a pie's angles and a triangle's tip written in the batch that adds
// the shape on the web and in the sync after it on the desktop
// (chart-primitives.ts), and one final sync that groups, names and tags the
// lot. Owns every shape-drawing Office.js call.
// Invariant: the group carries both tags and its children carry none; on the
// web no batch after a shape's own add resolves that shape by id, because the
// web cannot do that on a slide the editor never showed.

import type { Primitive } from "../chart-shapes";
import type { Box } from "../layout";
import { encodeTag, TAG_KEY, TAG_LINK, type LinkTag } from "../link/model";
import { cleanupByToken, cleanupShapes } from "./chart-cleanup";
import { SYNC_TIMEOUT_MS, type DrawHost } from "./chart-cost";
import {
  addPrimitive,
  adjustShape,
  adjustsInAddBatch,
  shapeAdjustments,
  turnShape,
  type Added,
} from "./chart-primitives";
import { rideDeckCopy } from "./deck-flow";
import { capAfter, drawClock, nextChunk, SHAPES_PER_SYNC } from "./draw-pace";

export { SHAPES_PER_SYNC, SYNC_TIMEOUT_MS };

// Which cost table, which deadline and which wording a draw runs under: the
// web is the host that charges per shape and tells the user to keep the tab in
// front. With no Office host at all (a bare unit test) it is the desktop.
export function drawHost(): DrawHost {
  if (typeof Office === "undefined") return "desktop";
  return Office.context?.platform === Office.PlatformType.OfficeOnline
    ? "web"
    : "desktop";
}

// A write batch on PowerPoint for the web can be swallowed whole: it neither
// applies nor rejects, the shapes of the chunks before it stay on the slide
// and the pane waits for ever (lessons, 08.09: a 21-shape pie stopped after
// its first chunk of twelve). Every round trip of a draw therefore has a
// deadline, per host (SYNC_TIMEOUT_MS in chart-cost.ts): the web's is just
// past the longest add sync that ever answered (27.1 s, S1) because none the
// model put past 30 s answered at all and a longer wait is dead time.

// PowerPoint for Mac (16.107 on macOS 26.6, 13.09.2026) dies on its own
// repaint after ShapeCollection.addGroup takes a whole chart at once: six
// shapes grouped in every run of the bisection, nineteen never did, whatever
// their kind, and the same shapes grouped through the ribbon were fine. So
// the Mac groups in tiers: sub-groups of this many members first, then those
// into the link's own group. Windows and the web keep the flat group.
export const GROUP_TIER_MAC = 6;

export function groupTier(): number | null {
  return Office.context?.platform === Office.PlatformType.Mac
    ? GROUP_TIER_MAC
    : null;
}

// What a round trip that never came back rejects with, so the caller can tell
// a host that stopped answering from one that refused the shapes. `what`
// names the round trip that stopped, so the sentence is true wherever this
// runs, not only for a chart's own draw.
export class ChartDrawTimeout extends Error {
  constructor(what: string) {
    super(`PowerPoint stopped answering while ${what}`);
    this.name = "ChartDrawTimeout";
  }
}

export function isDrawTimeout(error: unknown): boolean {
  return error instanceof ChartDrawTimeout;
}

// One piece of host work under that deadline: whatever it answers, unless it
// answers nothing at all. The race keeps a handler on the abandoned promise,
// so a batch that rejects long afterwards is still nobody's unhandled error.
// Every PowerPoint.run sync in the pane goes through here now, not only a
// chart's own draw, so `what` defaults to that call's own wording and every
// other caller names its own round trip.
export async function withSyncDeadline<T>(
  work: Promise<T>,
  what = "drawing the chart",
): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const deadline = new Promise<never>((_resolve, reject) => {
    timer = setTimeout(() => {
      reject(new ChartDrawTimeout(what));
    }, SYNC_TIMEOUT_MS[drawHost()]);
  });
  try {
    return await Promise.race([work, deadline]);
  } finally {
    clearTimeout(timer);
  }
}

// The teardown runs on the host that has just stopped answering, so it gets
// the same deadline and no more: a cleanup that hangs too would hold the
// pane's busy flag open exactly as the draw did.
async function cleanupWithin(
  slideId: string,
  ids: string[],
  namePrefix: string,
  // Set only for a fresh insert (insertChart never passes GroupSpec.before;
  // only refreshChart does, to delete the OLD group its own draw replaces).
  // A refresh's old group legitimately carries this SAME token already -
  // that is the whole point of a link keeping its identity across updates -
  // so cleanupByToken must never run there: a batch this sync rolled back
  // leaves that still-valid old group exactly as it was, tag and all, and a
  // token match would delete a link the deck still needs. Undefined means
  // "do not run it" rather than "run it for every token", so a caller that
  // forgets the distinction cannot silently reintroduce this.
  freshInsertToken: string | undefined,
  // The ids already on the target slide before THIS insert started
  // (GroupSpec.beforeIds): cleanupByToken never touches one of these, so an
  // older, finished copy of the same link survives even though it shares
  // the token being cleaned up by.
  before: ReadonlySet<string>,
): Promise<void> {
  try {
    await withSyncDeadline(cleanupShapes(slideId, ids, namePrefix));
    if (freshInsertToken === undefined) return;
    // The GROUPING sync itself is the one chunk cleanup cannot reach by id
    // or by name: addGroup nests the primitives (found above, by id) inside
    // a new shape whose own id was never read back (group.load("id") sat in
    // the same batch as the sync that just failed), and that new shape's
    // name is spec.name exactly, with none of the ": " suffix sweepByName
    // matches - by design, so the sweep never takes down an unrelated,
    // already-finished link that happens to share the chart's name. A host
    // that applied the addGroup and both tags before answering the sync
    // with an error leaves exactly that: a fully tagged, orphaned group.
    await withSyncDeadline(
      cleanupByToken(slideId, freshInsertToken, before),
    ).catch(() => undefined);
  } catch {
    // Best effort, the same swallow cleanupShapes makes of its own errors.
  }
}

// What one chart draw needs to know: the primitives in the layout's own
// coordinates, the box on the slide they are drawn at, the brand font every
// label wears, the group's name, its two tags, and whatever the caller wants
// queued in the grouping sync (a refresh deletes the old group there).
export interface GroupSpec {
  primitives: Primitive[];
  box: Box;
  font: string;
  name: string;
  tag: LinkTag;
  token: string;
  // The slide the group is drawn on, for the cleanup run a rejection sends
  // after it, which cannot reuse the context that failed.
  slideId: string;
  // How many shapes the slide already holds, group children included (a
  // refresh counts the old group it has not deleted yet): every add costs more
  // for each of them (S1), so each chunk is sized from this plus what the draw
  // has added so far. A caller that has not counted them passes none.
  existing?: number;
  before?: () => void;
  // resolveTarget's own `before` (placement.ts): the slide's ids as they
  // stood before a FRESH insert started. Never set by refreshChart, which
  // never reaches cleanupByToken at all (see cleanupWithin).
  beforeIds?: ReadonlySet<string>;
  // Told how many of the primitives the host has confirmed, once before the
  // first chunk (0) and once after every chunk's sync, so the pane can say
  // "48 of 92 shapes" while a slow host works.
  progress?: DrawProgress;
}

export type DrawProgress = (done: number, total: number) => void;

function chunks<T>(items: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let at = 0; at < items.length; at += size) {
    out.push(items.slice(at, at + size));
  }
  return out;
}

// What every shape of a chart is named with before its own part: the
// primitives, and on the Mac the tiers, all in the batch that adds them, so
// a cleanup can find by name what a refused batch left with no id.
export function chartPrefix(name: string): string {
  return `${name}: `;
}

// The ids the link's group takes: the shapes themselves, or on a host that
// groups in tiers (groupTier) the sub-groups of them, made in one sync per
// level and appended to the cleanup list - a shape inside a sub-group is no
// longer an id the slide's own collection can delete, but its sub-group is.
// Levels repeat until at most `tier` members remain, so a 40-point chart
// never asks addGroup for 19+ ids (the count that killed Mac 16.107). A
// remainder of one is left ungrouped: addGroup of one is InvalidArgument,
// and on Mac 16.107 (14.09) that sync left the sub-groups untagged.
async function tierUp(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  ids: string[],
  name: string,
): Promise<string[]> {
  const tier = groupTier();
  if (tier === null || ids.length <= tier) return [...ids];
  let members = [...ids];
  while (members.length > tier) {
    const next: string[] = [];
    const subs: PowerPoint.Shape[] = [];
    for (const part of chunks(members, tier)) {
      if (part.length < 2) {
        next.push(...part);
        continue;
      }
      const sub = shapes.addGroup(part);
      sub.name = `${chartPrefix(name)}tier`;
      sub.load("id");
      subs.push(sub);
    }
    await withSyncDeadline(context.sync());
    const grouped = subs.map((sub) => sub.id);
    next.push(...grouped);
    ids.push(...grouped);
    members = next;
  }
  return members;
}

// The primitives added in chunks, each chunk added and its ids read back in
// one sync. Each chunk is sized BEFORE its sync from the cost model for the
// slide as it stands (draw-pace.ts), and timed from its first add to its
// sync's answer: one that took over 1.5 times what the model said, and over
// 12 s, halves the cap for every chunk after it. `ids` is the caller's cleanup
// list, filled as syncs confirm, so a rejection halfway still knows what is on
// the slide.
async function drawChunks(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  spec: GroupSpec,
  ids: string[],
): Promise<void> {
  const total = spec.primitives.length;
  const clock = drawClock();
  const host = drawHost();
  const withTheAdd = adjustsInAddBatch(host);
  let cap = SHAPES_PER_SYNC;
  spec.progress?.(0, total);
  for (let done = 0; done < total;) {
    const { size, predictedMs } = nextChunk(
      cap,
      (spec.existing ?? 0) + done,
      total - done,
      host,
    );
    const started = clock.now();
    const added = spec.primitives
      .slice(done, done + size)
      .map((primitive): Added => {
        const shape = addPrimitive(shapes, primitive, spec.box, spec.font);
        shape.name = `${chartPrefix(spec.name)}${primitive.name}`;
        turnShape(shape, primitive);
        if (withTheAdd) adjustShape(shape, primitive);
        shape.load("id");
        return { shape, primitive };
      });
    await withSyncDeadline(context.sync());
    cap = capAfter(cap, clock.now() - started, predictedMs);
    ids.push(...added.map((one) => one.shape.id));
    if (!withTheAdd) shapeAdjustments(added);
    done += added.length;
    spec.progress?.(done, total);
  }
}

// The whole chart in ceil(primitives / chunk) + 1 round trips: the chunks
// above, and a last sync that groups them, names the group, tags it and runs
// whatever the caller queued there.
//
// A sync that rejects has already put shapes on the host: PowerPoint.run
// does not roll those back, and a refused batch keeps the adds queued before
// the call it refused (Mac 16.107, 14.09). So this tracks every id a prior
// sync confirmed and, on any rejection, deletes them itself (chart-cleanup.ts)
// and then sweeps the slide for the rest by the chart's name, before
// rethrowing the rejection unchanged. A round trip the host swallows without
// answering reaches the same path through its deadline, as a ChartDrawTimeout
// the caller turns into a sentence.
export async function drawGroup(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  spec: GroupSpec,
): Promise<string> {
  const ids: string[] = [];
  try {
    await drawChunks(context, shapes, spec, ids);
    const members = await tierUp(context, shapes, ids, spec.name);
    // On the desktop the last chunk's tip and angle writes (shapeAdjustments)
    // open this batch; the web wrote them in the batch that added the shapes.
    // The deck's record follows, ahead of the group and the old group's
    // delete: a write the host cannot take halts the batch before anything
    // changes, so the cleanup below leaves the old chart as it was and the
    // repaint can run again. Not on the Mac: this is the batch 16.107 died
    // after (GROUP_TIER_MAC).
    if (groupTier() === null) rideDeckCopy(context, spec.tag.id);
    const group = shapes.addGroup(members);
    group.name = spec.name;
    group.tags.add(TAG_LINK, encodeTag(spec.tag));
    group.tags.add(TAG_KEY, spec.token);
    group.load("id");
    spec.before?.();
    await withSyncDeadline(context.sync());
    return group.id;
  } catch (err) {
    await cleanupWithin(
      spec.slideId,
      ids,
      chartPrefix(spec.name),
      // No pre-add ids, no token sweep: without them an older copy of the
      // same link is indistinguishable from this draw's own orphan.
      spec.before === undefined && spec.beforeIds ? spec.token : undefined,
      spec.beforeIds ?? new Set(),
    );
    throw err;
  }
}
