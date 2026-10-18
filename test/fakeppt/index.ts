// A hand-built PowerPoint host: an in-memory deck answering exactly the
// office.js surface src/ppt/host.ts touches, so link insert, refresh and break
// can be driven without a sideload. Strict load semantics are opt-in and
// global, the way the Excel fake does it; the two fakes share no code.

import type { AddCostClock } from "../clock";
import { bindAddCosts } from "./add-cost";
import type { FakePresentation } from "./model";
import type { FormatPart } from "./refusals";
import { installGlobals, removeGlobals } from "./office";
import type { RangeColor } from "./text-selection";

export type {
  FakePptShape,
  FakeShapeFont,
  FakeShapeGroup,
  FakeShapeInit,
  FakeShapeMargins,
  FakeSlide,
  ShapeSite,
} from "./model";
export type { CustomXmlSnapshot, XmlOwner } from "./customxml";
export type { RangeColor } from "./text-selection";
export type {
  FakeTable,
  FakeTableCell,
  FakeTableCellFill,
  FakeTableStyleSettings,
  TableUniformCellProperties,
} from "./tables";
// What the deck's custom XML parts hold, owner by owner: with the open batch's
// writes in (Snapshot) and as the last landed sync left them (Committed).
export { customXmlCommitted, customXmlSnapshot } from "./customxml";
export { FakePresentation } from "./model";
// The current selection, in order: what getSelectedShapes() hands back.
export { selectedShapeIds } from "./objects";

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

// One Office.context.document.setSelectedDataAsync call, as the fake saw it.
export interface SelectionInsert {
  slideId: string;
  png: string;
  box: Box;
}

export interface FakePptOptions {
  presentation?: FakePresentation;
  slides?: number;
  isSetSupported?: (set: string, version: string) => boolean;
  strictLoad?: boolean;
  storage?: Map<string, string>;
  // The clock every shapes.add* charges (test/clock.ts costedClock): a host
  // slow in the web's two terms, so a test can prove an adaptive chunk size by
  // the time a chunk took. Left out, adds are free and no clock is read.
  clock?: AddCostClock;
}

export interface FakePptHelpers {
  selectSlide(id: string): void;
  clearSelection(): void;
  // Sets the shape selection directly, order preserved, skipping the
  // same-slide rule Slide.setSelectedShapes enforces: a fixture shortcut
  // mirroring selectSlide.
  selectShapes(shapeIds: string[]): void;
  setSupported(check: (set: string, version: string) => boolean): void;
  // Registers a style the NEXT shapes.addTable on this shape id starts with,
  // in place of the Mac's own default of no style at all - the only way to
  // make a just-inserted or just-rebuilt table arrive already styled, the
  // way a host other than the Mac might hand one back. Ids are sequential
  // per deck (`shape-1`, `shape-2`, ...), so a fresh presentation's first
  // table is always "shape-1".
  setTableStyle(shapeId: string, style: string): void;
  // PowerPoint for the web, rig 27.09 (src/ppt/shape-ready.ts): the NEXT
  // shape addTable creates stays invisible to shapes.getItem(id), to the
  // slide's own item list and to any proxy addressing it by id, for this
  // many syncs following its own add - though that add's own sync still
  // hands back its id, same as the real host.
  delayRegistration(syncs: number): void;
  // PowerPoint for the web, S1 spike section 6: a slide outside the editor's
  // loaded range reads as empty (its shape list and its count) for about 15 s
  // after a deck opens, though its shapes are there and a write by id works.
  // Held from now until released; the model keeps the shapes either way.
  holdSlideReads(slideId: string): void;
  releaseSlideReads(slideId: string): void;
  // PowerPoint for the web, S1 spike section 6.3 and the rig of 05.10: a slide
  // the editor has never shown. It reads empty like a held slide, and a shape
  // added to it in this session cannot be addressed by id once the sync of its
  // own add has landed: getItem(id), and any proxy a LATER batch uses (a
  // write, a delete, an adjustment), fail with "InvalidParam passed to
  // GetItem(id)" (5010), getItemOrNullObject answers null; an add and
  // addGroup(ids) still work, and a table is found as on any slide (its add
  // loads the slide; rig 05.10). This is the worst case the rig met, not every
  // run: code that passes here resolves nothing by id in a later batch.
  // Showing the slide (setSelectedSlides) ends it. Throws unless
  // setPlatform("OfficeOnline") came first: no other host has such a slide.
  leaveSlideUnshown(slideId: string): void;
  // Office.context.platform, "Mac" until a test asks for another host: the web
  // is the one that draws fewer shapes before it gives up.
  setPlatform(platform: string): void;
  storage(): Map<string, string>;
  // Round trips to the host since this fake was installed: what a batching
  // change is measured in, counted from zero per installFakePpt.
  syncCount(): number;
  // The writes PowerPoint for the web's editor counts as edits (F1 spike,
  // 05.10): a document property, a shape's name or box, whatever value is
  // written. A tag or a custom XML part never counts, which is why a Break alone
  // is not saved there. What src/ppt/save-nudge.ts is measured by.
  editorWrites(): number;
  // Everything a person can see of the deck - slides, shapes, their names,
  // boxes, fills and text, and the document's title - as one string, with no tag
  // and no custom XML part: equal before and after means nothing visible moved.
  visibleState(): string;
  insertedViaSelection(): SelectionInsert[];
  // Makes the next setSelectedDataAsync report failure, the way a host that is
  // out of memory or has the slide locked does. Consumed by that one call.
  failNextSelectionInsert(message?: string): void;
  // Makes a future context.sync() reject, the way a hung or refused round
  // trip does: the very next one by default, or the one after skipping
  // `afterSyncs` more that still succeed. Calls stack, so more than one
  // future sync can be armed to fail. Whichever sync it lands on takes back
  // off the deck everything added since the sync before it.
  failNextSync(error?: Error, afterSyncs?: number): void;
  // Makes a future context.sync() reject the way a real refusal does: the
  // host applied the batch's adds and then refused, so they stay on the deck
  // with ids no load answered (Mac 16.107, 14.09: four sub-groups landed and
  // the sync rejected on the fifth addGroup). Same ordinal rule as
  // failNextSync; every add of that batch is kept, the loads are lost.
  refuseNextSync(error?: Error, afterSyncs?: number): void;
  // G audit: makes a future context.sync() never settle - neither applying
  // its batch nor rejecting it - the way PowerPoint for the web swallowed a
  // chart draw's second chunk (tasks/lessons.md, 2026-09-08). Same ordinal
  // rule as failNextSync; the shapes of that batch never reach the deck.
  hangNextSync(afterSyncs?: number): void;
  // Makes every Shape.rotation write throw InvalidArgument, as a host that
  // cannot rotate does; the writes queued after it in that batch never run.
  refuseRotation(): void;
  // Makes every write of a triangle's tip (adjustment 0) halt its batch at that
  // write: what was queued before it lands, the rest never runs, and the sync
  // rejects. The tips are the first writes of the sync after a chunk's, so a
  // refusal there stops the grouping sync before the group and the old group's
  // delete.
  refuseTriangleTips(): void;
  // PowerPointApi 1.10 theme colors (themes.ts): the slide masters of the deck,
  // each answering Accent 1-6 and the rest like the Office theme until written.
  slideMasterIds(): string[];
  // A new master: the Office theme with `colors` (by theme color name) over it.
  addSlideMaster(colors?: Record<string, string>): string;
  // What a master holds now, by index or id, "#RRGGBB".
  themeColor(master: string | number, name: string): string;
  // Seeds a master's colors straight, no batch: a deck that arrives themed.
  setThemeColors(master: string | number, colors: Record<string, string>): void;
  // getThemeColor answers a bare "RRGGBB", the d.ts example's spelling.
  answerBareThemeColors(bare?: boolean): void;
  // A batch that reads a theme color is refused at its sync (Mac 16.107 reports
  // 1.10 and refuses some loads).
  // `letThrough`: the color reads answered before the refusal starts.
  refuseThemeReads(on?: boolean, letThrough?: number): void;
  // The write after `letThrough` more writes halts its batch: what came before
  // landed, what came after never ran. `forever`: every later batch halts at its
  // first write too (a host that cannot change theme colors at all).
  refuseThemeWrites(letThrough?: number, forever?: boolean): void;
  // PowerPointApi 1.5: highlights `length` characters of the shape from `start`
  // and selects it. Selecting other shapes drops the highlight.
  selectText(shapeId: string, start: number, length: number): void;
  // The colors written through a selected range of the shape, in order: the
  // shape's own font color is not touched by them.
  rangeColors(shapeId: string): RangeColor[];
  // From now on this shape refuses every fill, outline and text font write, or
  // only those of the parts named: its batch halts there and fails at the sync
  // (refusals.ts).
  refuseFormatOn(shapeId: string, parts?: FormatPart[]): void;
  // The ribbon FunctionNames registerCommands associated with this host.
  commandIds(): string[];
  // Presses one of them the way the ribbon does, resolving when the command
  // calls event.completed().
  runCommand(id: string): Promise<void>;
}

let strictByDefault = false;

// Switches strict load semantics on for every host installed afterwards, which
// is how a suite runs the add-in against real office.js load ordering.
// installFakePpt({ strictLoad }) overrides it per host.
export function enableStrictLoadSemantics(on = true): void {
  strictByDefault = on;
}

export function installFakePpt(options: FakePptOptions = {}): {
  presentation: FakePresentation;
  helpers: FakePptHelpers;
} {
  const installed = installGlobals(options, strictByDefault);
  bindAddCosts(installed.presentation, options.clock);
  return installed;
}

export function uninstallFakePpt(): void {
  removeGlobals();
}
