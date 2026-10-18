// Walking the shapes of a deck, groups included: PowerPointApi 1.8 exposes
// Shape.group.shapes, so a picture the user dragged into a group is still
// reachable - by the path of group ids running from the slide down to it, and
// that path is how the same shape is addressed again on a refresh.

import { withSyncDeadline } from "./chart-draw";

export const SHAPE_PROPERTIES =
  "items/id,items/type,items/left,items/top,items/width,items/height";

// Groups (Shape.type "Group", Shape.group, ShapeGroup.shapes) and
// fill.setImage both arrived in PowerPointApi 1.8, so one version gates both.
export const GROUP_API = "1.8";
// PowerPoint.ShapeType.group; the typings have no TextBox member, so a shape's
// type is compared as the string the host reports.
export const GROUP_TYPE = "Group";
// PowerPoint reports a layout placeholder as its own shape type; an empty one
// is the slide's "click to add" furniture, not an object to place around.
export const PLACEHOLDER_TYPE = "Placeholder";

// A user can nest groups without limit; three levels is every real deck, and
// the cap is what keeps a pathological one from costing a sync per level.
const MAX_DEPTH = 3;

// Where a shape sits: on a slide, or inside groups nested in one.
export interface ShapePath {
  slideId: string;
  shapeId: string;
  groupPath?: string[];
}

export interface PlacedShape {
  slideId: string;
  slideIndex: number;
  shape: PowerPoint.Shape;
  groupPath: string[];
}

// Who has PowerPointApi 1.5, which the pane boots on and every object tool
// needs: Microsoft 365 2208, PowerPoint for Mac 16.64 and Office 2024. Office
// 2021 volume licences stop at 1.2. The one wording of the boot badge and of
// the object tools' refusal.
export const NEEDS_API_1_5 =
  "PowerPoint Microsoft 365 2208 / Mac 16.64 / Office 2024 or newer";

// A host that reports no requirements at all is the newest one, the web.
export function hasPowerPointApi(version: string): boolean {
  const requirements = Office.context?.requirements;
  return requirements
    ? requirements.isSetSupported("PowerPointApi", version)
    : true;
}

// The whole selection surface (getSelectedShapes, setSelectedShapes,
// getParentSlide) is PowerPointApi 1.5; the fill/line members below it are
// 1.4, so gating every entry point at 1.5 covers both.
const OBJECT_TOOLS_API = "1.5";

// The one gate and the one wording of every tool that acts on the selection:
// the object tools, the brand swatches, Smart Painter and the shape styles
// (re-exported by object-tools.ts, where placement.ts and shape-paint.ts read it).
export function requireObjectToolsApi(): void {
  if (!hasPowerPointApi(OBJECT_TOOLS_API)) {
    throw new Error(`Object tools need ${NEEDS_API_1_5}.`);
  }
}

export function isGrouped(path: ShapePath): boolean {
  return (path.groupPath?.length ?? 0) > 0;
}

// The shape a link lives on, addressed the way it was found: down through its
// groups when it has any, straight off the slide when it has none.
export function shapeAt(
  context: PowerPoint.RequestContext,
  path: ShapePath,
): PowerPoint.Shape {
  const shapes = context.presentation.slides.getItem(path.slideId).shapes;
  const groups = path.groupPath ?? [];
  const outermost = groups[0];
  if (outermost === undefined) return shapes.getItem(path.shapeId);
  let shape = shapes.getItem(outermost);
  for (const groupId of groups.slice(1)) {
    shape = shape.group.shapes.getItem(groupId);
  }
  return shape.group.shapes.getItem(path.shapeId);
}

// Every group at one nesting level is opened in a single batch, so depth costs
// a round trip and width costs none. Below 1.8 there is nothing to open.
export async function expandGroups(
  context: PowerPoint.RequestContext,
  top: PlacedShape[],
): Promise<PlacedShape[]> {
  if (!hasPowerPointApi(GROUP_API)) return top;
  const all = [...top];
  let frontier = top;
  for (let depth = 0; depth < MAX_DEPTH && frontier.length > 0; depth += 1) {
    frontier = await openGroups(context, frontier);
    all.push(...frontier);
  }
  return all;
}

// How many shapes a slide holds as PowerPoint for the web counts them: every
// shape, a group's children and the group itself. Each add pays for every one
// (S1), so a chart drawn onto a slide that already holds a 92-shape chart pays
// for 93, not for the one group that shows. One read of the slide and one per
// level of nesting, so a slide with no group costs a single sync.
export async function slidePopulation(
  context: PowerPoint.RequestContext,
  slideId: string,
): Promise<number> {
  const slide = context.presentation.slides.getItem(slideId);
  slide.shapes.load("items/id,items/type");
  await withSyncDeadline(context.sync(), "counting the shapes on the slide");
  // The slide's index is not read below, only the paths it carries.
  const top = slide.shapes.items.map((shape) => ({
    slideId,
    slideIndex: 0,
    shape,
    groupPath: [],
  }));
  return (await expandGroups(context, top)).length;
}

// The children of every group in the frontier, each carrying the path that
// reaches it: its parent's path plus the id of the group it came out of.
async function openGroups(
  context: PowerPoint.RequestContext,
  frontier: PlacedShape[],
): Promise<PlacedShape[]> {
  const opened = frontier
    .filter((entry) => entry.shape.type === GROUP_TYPE)
    .map((entry) => {
      const shapes = entry.shape.group.shapes;
      shapes.load(SHAPE_PROPERTIES);
      return { entry, shapes };
    });
  if (opened.length === 0) return [];
  // The one round trip a grouped deck adds per level; a host that swallows
  // it would otherwise hang every scan of the deck.
  await withSyncDeadline(context.sync(), "reading the groups");
  return opened.flatMap(({ entry, shapes }) =>
    shapes.items.map((shape) => ({
      slideId: entry.slideId,
      slideIndex: entry.slideIndex,
      shape,
      groupPath: [...entry.groupPath, entry.shape.id],
    })),
  );
}
