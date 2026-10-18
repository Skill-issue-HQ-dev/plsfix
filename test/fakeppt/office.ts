// The host itself: the runtime a test installs, the request context PowerPoint
// hands a batch, and the PowerPoint, Office and OfficeRuntime globals the pane
// talks to. The slide, shape and tag objects the context leads to live in
// objects.ts, over the deck in model.ts.

import type { Box, FakePptHelpers, FakePptOptions, SelectionInsert } from ".";
import { presentationXmlParts } from "./customxml";
import { editorWrites, visibleState } from "./editor-state";
import { GEOMETRIC_SHAPE_TYPE } from "./geometry";
import { FakePresentation } from "./model";
import { leaveSlideUnshown, showSlide } from "./never-shown";
import { SlideCollectionProxy } from "./objects";
import { bindPlatform } from "./platform";
import { DocumentPropertiesProxy } from "./properties";
import { holdSlideReads, releaseSlideReads } from "./read-hold";
import { armNextShapeDelay, onSyncSucceeded } from "./register-delay";
import { refuseFormatOn } from "./refusals";
import {
  SelectedShapesCollectionProxy,
  selectShapesForTest,
} from "./selection";
import { Loadable, StrictLoads } from "./strict";
import { setPendingTableStyle } from "./tables";
import { rangeColorsOf, selectedTextRange, selectText } from "./text-selection";
import {
  addSlideMaster,
  answerBareColors,
  masterIds,
  refuseThemeReads,
  refuseThemeWrites,
  seedThemeColors,
  slideMasterCollection,
  themeColorOf,
  THEME_COLOR,
} from "./themes";

interface SelectionOptions {
  coercionType?: string;
  imageLeft?: number;
  imageTop?: number;
  imageWidth?: number;
  imageHeight?: number;
}

// One armed context.sync() failure: which sync ordinal (1-based, across the
// whole host) it fires on and what it rejects with. Several can be queued at
// once, for a caller that needs more than one future sync to fail. `applied`
// is the real host's refusal (Mac 16.107, 14.09: the sub-groups queued before
// the refused addGroup were on the slide when the sync rejected): the batch's
// adds stay on the deck, only the loads are lost. Without it the batch never
// reached the host at all and its adds come back off.
interface ArmedSyncFailure {
  at: number;
  error: Error;
  applied: boolean;
}

// What Office.actions.associate hands the ribbon: the handler completes the
// event when the command is done, which is the host's own signal.
type CommandHandler = (event?: { completed: () => void }) => void;

// The message and code a rejected sync carries when a test does not supply
// its own, styled like the host's own "the operation could not complete".
function syncFailure(message?: string): Error {
  return Object.assign(
    new Error(message ?? "PowerPoint could not complete the request."),
    { code: "GeneralException" },
  );
}

class FakeRuntime {
  strict: StrictLoads | null;
  supported: (set: string, version: string) => boolean;
  storage: Map<string, string>;
  insertions: SelectionInsert[] = [];
  nextInsertFailure: string | null = null;
  // Armed by helpers.failNextSync(): consumed by the sync whose ordinal
  // matches, in case more than one is queued for the same host.
  syncFailures: ArmedSyncFailure[] = [];
  // Office.context.platform: the Windows desktop until a test says otherwise,
  // because the web is the host with the tighter shape budget and the Mac the
  // one that keeps every chart a picture (charts.ts hostDrawsCharts).
  platform = "PC";
  // Every context.sync() this host served: one round trip to PowerPoint, and
  // the only cost a deck's size is allowed to multiply.
  syncs = 0;
  // G audit: PowerPoint for the web can take a write batch and never answer at
  // all - a 21-shape pie stopped after its first chunk of twelve and the pane
  // stayed busy for ever (tasks/lessons.md, 2026-09-08). Armed by
  // helpers.hangNextSync(): the sync ordinals whose promise never settles.
  hangSyncs: number[] = [];
  // Every ribbon FunctionName registerCommands associated, by id.
  commands = new Map<string, CommandHandler>();

  constructor(
    public presentation: FakePresentation,
    options: FakePptOptions,
    strictDefault: boolean,
  ) {
    this.supported = options.isSetSupported ?? (() => true);
    this.storage = options.storage ?? new Map();
    this.strict =
      (options.strictLoad ?? strictDefault) ? new StrictLoads() : null;
  }
}

class FakeContext extends Loadable {
  presentation: PresentationProxy;

  constructor(private runtime: FakeRuntime) {
    super();
    this.presentation = new PresentationProxy(runtime.presentation);
  }

  // A no-op flush: reads come from the deck, so only load state moves here.
  // Counted, because a batch that syncs per shape is the performance bug.
  // Every shape added in the batch is now one the host has heard of - unless
  // this is the sync a test armed to fail, in which case none of what it
  // queued reached the host: the adds since the last sync come back off the
  // deck and the requested loads stay unreadable, the way a batch that never
  // arrived leaves nothing behind.
  sync(): Promise<void> {
    this.runtime.syncs += 1;
    const at = this.runtime.syncs;
    // G audit: a batch the host swallowed - neither applied nor refused. The
    // slide kept only the chunks before it (lessons 08.09: a 21-shape pie
    // left its first twelve behind), so its own adds come back off the deck
    // exactly as a rejection's do; the promise simply never settles.
    if (this.runtime.hangSyncs.includes(at)) {
      this.runtime.presentation.rollbackPending();
      this.runtime.strict?.drop();
      // A promise with no settle path at all: the executor takes none.
      return new Promise<void>(() => undefined);
    }
    const index = this.runtime.syncFailures.findIndex(
      (failure) => failure.at === at,
    );
    if (index !== -1) {
      const [failure] = this.runtime.syncFailures.splice(index, 1);
      if (failure!.applied) {
        this.runtime.presentation.confirmPending();
        this.runtime.presentation.markSynced();
      } else {
        this.runtime.presentation.rollbackPending();
      }
      this.runtime.strict?.drop();
      return Promise.reject(failure!.error);
    }
    // An id the batch used that is not there (customxml.ts): the host halted
    // at that op, so this sync lands what came before it and refuses with it.
    const halted = this.runtime.presentation.settleHalt();
    if (halted) {
      this.runtime.presentation.markSynced();
      this.runtime.strict?.drop();
      return Promise.reject(halted);
    }
    this.runtime.presentation.confirmPending();
    this.runtime.presentation.markSynced();
    this.runtime.strict?.commit();
    // register-delay.ts: a landed sync is what ages a delayed shape closer
    // to visible, whatever the batch was for.
    onSyncSucceeded(this.runtime.presentation);
    return Promise.resolve();
  }
}

class PresentationProxy extends Loadable {
  constructor(private deck: FakePresentation) {
    super();
  }

  get slides(): SlideCollectionProxy {
    return new SlideCollectionProxy(this.deck, () => this.deck.slides);
  }
  getSelectedSlides(): SlideCollectionProxy {
    return new SlideCollectionProxy(this.deck, () =>
      this.deck.selectedSlides(),
    );
  }
  setSelectedSlides(slideIds: string[]): void {
    this.deck.selectedSlideIds = [...slideIds];
    // A slide brought into view is loaded by the editor: it reads what it
    // holds from then on (read-hold.ts), and a never-shown one resolves its
    // shapes by id again (never-shown.ts, which releases the read hold too).
    for (const slideId of slideIds) showSlide(this.deck, slideId);
  }
  getSelectedShapes(): SelectedShapesCollectionProxy {
    return new SelectedShapesCollectionProxy(this.deck);
  }
  // PowerPointApi 1.5: the highlighted characters, or none (text-selection.ts).
  getSelectedTextRange() {
    return selectedTextRange(this.deck, false);
  }
  getSelectedTextRangeOrNullObject() {
    return selectedTextRange(this.deck, true);
  }
  // PowerPointApi 1.3 masters, each with a 1.10 theme color scheme (themes.ts).
  get slideMasters() {
    return slideMasterCollection(this.deck);
  }
  get customXmlParts() {
    return presentationXmlParts(this.deck);
  }
  // PowerPointApi 1.7: the document's own properties, its title among them.
  get properties(): DocumentPropertiesProxy {
    return new DocumentPropertiesProxy(this.deck);
  }
}

// ---------------------------------------------------------------------------
// Globals
// ---------------------------------------------------------------------------

function pickCallback(
  ...candidates: unknown[]
): ((result: unknown) => void) | undefined {
  for (const candidate of candidates) {
    if (typeof candidate === "function") {
      return candidate as (result: unknown) => void;
    }
  }
  return undefined;
}

// Inserting a picture through the selection API, the only route on hosts
// without the shape APIs: it lands on the selected slide, at the given box.
function insertViaSelection(
  runtime: FakeRuntime,
  png: string,
  options: SelectionOptions,
): void {
  const failure = runtime.nextInsertFailure;
  if (failure !== null) {
    runtime.nextInsertFailure = null;
    throw new Error(failure);
  }
  const deck = runtime.presentation;
  const slideId = deck.selectedSlideIds[0] ?? deck.slides[0]?.id ?? "";
  const slide = deck.findSlideOrThrow(slideId);
  const box: Box = {
    left: options.imageLeft ?? 0,
    top: options.imageTop ?? 0,
    width: options.imageWidth ?? 0,
    height: options.imageHeight ?? 0,
  };
  // The picture is the shape's own image; nothing called fill.setImage.
  deck.addShape(slide, { type: "Image", fillImage: png, ...box });
  runtime.insertions.push({ slideId: slide.id, png, box });
}

function documentApi(runtime: FakeRuntime): Record<string, unknown> {
  return {
    setSelectedDataAsync(data: unknown, first?: unknown, second?: unknown) {
      const done = pickCallback(first, second);
      const options = (
        typeof first === "object" && first !== null ? first : {}
      ) as SelectionOptions;
      queueMicrotask(() => {
        try {
          insertViaSelection(runtime, String(data), options);
          done?.({ status: "succeeded", value: undefined });
        } catch (error) {
          done?.({ status: "failed", error });
        }
      });
    },
    getFilePropertiesAsync(first?: unknown, second?: unknown) {
      const done = pickCallback(first, second);
      const url = runtime.presentation.fileUrl;
      queueMicrotask(() => {
        done?.({ status: "succeeded", value: { url } });
      });
    },
  };
}

function officeGlobal(runtime: FakeRuntime): Record<string, unknown> {
  return {
    // Office.actions.associate: the ribbon's FunctionName table. Without it
    // registerCommands returns early, so a pane test could never fire a
    // ribbon press; helpers.runCommand(id) is how a test does.
    actions: {
      associate: (id: string, handler: CommandHandler) => {
        runtime.commands.set(id, handler);
      },
    },
    // Office.addin.showAsTaskpane: what a ribbon or shortcut error brings to
    // the front when the pane is closed. A plain resolved promise, the way
    // test/fakehost.ts's Excel fake already answers it; a test wraps this in
    // vi.spyOn to see whether and how often it was called.
    addin: { showAsTaskpane: () => Promise.resolve() },
    context: {
      get platform(): string {
        return runtime.platform;
      },
      requirements: {
        isSetSupported: (set: string, version: string) =>
          runtime.supported(set, version),
      },
      document: documentApi(runtime),
    },
    HostType: { Excel: "Excel", Word: "Word", PowerPoint: "PowerPoint" },
    // office.js gives these string values, whatever the .d.ts enum looks like.
    PlatformType: {
      PC: "PC",
      OfficeOnline: "OfficeOnline",
      Mac: "Mac",
      iOS: "iOS",
      Android: "Android",
      Universal: "Universal",
    },
    CoercionType: { Text: "text", Image: "image", SlideRange: "slideRange" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    // Office.onReady both calls back and resolves with the host it found.
    onReady: (callback?: (info: { host: string }) => unknown) => {
      const info = { host: "PowerPoint" };
      callback?.(info);
      return Promise.resolve(info);
    },
  };
}

function powerPointGlobal(runtime: FakeRuntime): Record<string, unknown> {
  return {
    run(first: unknown, second?: unknown): Promise<unknown> {
      const callback = (typeof first === "function" ? first : second) as (
        context: unknown,
      ) => unknown;
      const context = new FakeContext(runtime);
      const handed = runtime.strict?.root(context, "context") ?? context;
      return Promise.resolve().then(() => callback(handed));
    },
    GeometricShapeType: GEOMETRIC_SHAPE_TYPE,
    ConnectorType: { straight: "Straight", elbow: "Elbow", curve: "Curve" },
    ShapeAutoSize: {
      autoSizeNone: "AutoSizeNone",
      autoSizeTextToFitShape: "AutoSizeTextToFitShape",
      autoSizeShapeToFitText: "AutoSizeShapeToFitText",
      autoSizeMixed: "AutoSizeMixed",
    },
    ShapeType: {
      unsupported: "Unsupported",
      image: "Image",
      geometricShape: "GeometricShape",
      group: "Group",
      line: "Line",
      table: "Table",
    },
    ThemeColor: THEME_COLOR,
    TableStyle: {
      noStyleNoGrid: "NoStyleNoGrid",
      noStyleTableGrid: "NoStyleTableGrid",
      mediumStyle2Accent1: "MediumStyle2Accent1",
    },
  };
}

function officeRuntimeGlobal(runtime: FakeRuntime): Record<string, unknown> {
  return {
    storage: {
      getItem: (key: string) =>
        Promise.resolve(runtime.storage.get(key) ?? null),
      setItem: (key: string, value: string) => {
        runtime.storage.set(key, value);
        return Promise.resolve();
      },
      removeItem: (key: string) => {
        runtime.storage.delete(key);
        return Promise.resolve();
      },
    },
  };
}

function makeHelpers(runtime: FakeRuntime): FakePptHelpers {
  return {
    selectSlide(id) {
      runtime.presentation.findSlideOrThrow(id);
      runtime.presentation.selectedSlideIds = [id];
    },
    clearSelection() {
      runtime.presentation.selectedSlideIds = [];
    },
    selectShapes(ids) {
      selectShapesForTest(runtime.presentation, ids);
    },
    setSupported(check) {
      runtime.supported = check;
    },
    setTableStyle: (shapeId, style) =>
      setPendingTableStyle(runtime.presentation, shapeId, style),
    delayRegistration(syncs) {
      armNextShapeDelay(runtime.presentation, syncs);
    },
    holdSlideReads: (slideId) => holdSlideReads(runtime.presentation, slideId),
    releaseSlideReads: (slideId) =>
      releaseSlideReads(runtime.presentation, slideId),
    leaveSlideUnshown: (slideId) =>
      leaveSlideUnshown(runtime.presentation, slideId),
    setPlatform(platform) {
      runtime.platform = platform;
    },
    storage: () => runtime.storage,
    syncCount: () => runtime.syncs,
    editorWrites: () => editorWrites(runtime.presentation),
    visibleState: () => visibleState(runtime.presentation),
    failNextSelectionInsert(
      message = "PowerPoint could not insert the image.",
    ) {
      runtime.nextInsertFailure = message;
    },
    failNextSync(error, afterSyncs = 0) {
      runtime.syncFailures.push({
        at: runtime.syncs + afterSyncs + 1,
        error: error ?? syncFailure(),
        applied: false,
      });
    },
    refuseNextSync(error, afterSyncs = 0) {
      runtime.syncFailures.push({
        at: runtime.syncs + afterSyncs + 1,
        error: error ?? syncFailure(),
        applied: true,
      });
    },
    insertedViaSelection: () =>
      runtime.insertions.map((insert) => ({
        ...insert,
        box: { ...insert.box },
      })),
    // G audit: the web host that takes a batch and never answers.
    hangNextSync(afterSyncs = 0) {
      runtime.hangSyncs.push(runtime.syncs + afterSyncs + 1);
    },
    // A host whose Shape.rotation setter refuses (PowerPointApi 1.10 absent in
    // practice, or an old build): every later write of a batch never happens.
    refuseRotation() {
      runtime.presentation.refuseRotation = true;
    },
    refuseTriangleTips() {
      runtime.presentation.refuseTips = true;
    },
    slideMasterIds: () => masterIds(runtime.presentation),
    addSlideMaster: (colors) => addSlideMaster(runtime.presentation, colors),
    themeColor: (master, name) =>
      themeColorOf(runtime.presentation, master, name),
    setThemeColors: (master, colors) =>
      seedThemeColors(runtime.presentation, master, colors),
    answerBareThemeColors: (bare) =>
      answerBareColors(runtime.presentation, bare),
    refuseThemeReads: (on, letThrough) =>
      refuseThemeReads(runtime.presentation, on, letThrough),
    refuseThemeWrites: (letThrough, forever) =>
      refuseThemeWrites(runtime.presentation, letThrough, forever),
    selectText: (shapeId, start, length) =>
      selectText(runtime.presentation, shapeId, start, length),
    rangeColors: (shapeId) => rangeColorsOf(runtime.presentation, shapeId),
    refuseFormatOn: (shapeId, parts) =>
      refuseFormatOn(runtime.presentation, shapeId, parts),
    commandIds: () => [...runtime.commands.keys()],
    runCommand(id) {
      const handler = runtime.commands.get(id);
      if (!handler) throw new Error(`no ribbon command "${id}"`);
      // The host waits for event.completed(); so does the caller.
      return new Promise<void>((done) => {
        handler({ completed: () => done() });
      });
    },
  };
}

export function installGlobals(
  options: FakePptOptions,
  strictDefault: boolean,
): { presentation: FakePresentation; helpers: FakePptHelpers } {
  const presentation =
    options.presentation ?? new FakePresentation(options.slides);
  const runtime = new FakeRuntime(presentation, options, strictDefault);
  bindPlatform(presentation, () => runtime.platform);
  const scope = globalThis as unknown as Record<string, unknown>;
  scope.PowerPoint = powerPointGlobal(runtime);
  scope.Office = officeGlobal(runtime);
  scope.OfficeRuntime = officeRuntimeGlobal(runtime);
  return { presentation, helpers: makeHelpers(runtime) };
}

export function removeGlobals(): void {
  const scope = globalThis as unknown as Record<string, unknown>;
  delete scope.PowerPoint;
  delete scope.Office;
  delete scope.OfficeRuntime;
}
