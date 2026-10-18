// Whether and how a chart insert goes ahead: the slide's population (PowerPoint
// for the web charges every add for each shape already there), the one question
// when the draw is long or the slide too full, and the answer. Owns that
// decision and nothing else; the drawing is chart-draw.ts.
// Invariant: the picture is only ever his answer, and calling it off is a
// CalledOff - cancelling leaves the slide and the Inbox as they were.

import { INSERT_CANCELLED } from "../link/chart-model";
import type { InboxItem } from "../link/model";
import { CalledOff } from "../ui/guard";
import { drawHost } from "./chart-draw";
import { chooseDrawing } from "./draw-progress";
import { slidePopulation } from "./shapes";

// The shapes the slide holds as the host that charges for them counts them,
// group children and the groups themselves included. Only the web charges per
// shape already there, so only the web pays the extra read; elsewhere the
// count the caller already has stands in.
export async function populationOf(
  context: PowerPoint.RequestContext,
  slideId: string,
  known: number,
): Promise<number> {
  return drawHost() === "web" ? slidePopulation(context, slideId) : known;
}

// What an insert does with this chart on this slide.
export interface Decision {
  choice: "draw" | "picture";
  // The slide's population, for sizing the chunks of a draw.
  existing: number;
}

export async function decideDrawing(
  context: PowerPoint.RequestContext,
  slideId: string,
  item: InboxItem,
  shapeCount: number,
  hasPicture: boolean,
  known: number,
): Promise<Decision> {
  const existing = await populationOf(context, slideId, known);
  const choice = await chooseDrawing({
    label: item.label,
    shapes: shapeCount,
    existing,
    hasPicture,
  });
  if (choice === "cancel") throw new CalledOff(INSERT_CANCELLED);
  return { choice, existing };
}
