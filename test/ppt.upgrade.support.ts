// What the picture-era upgrade suites share: a pane over the fake deck, a chart
// link that landed as a picture before charts drew as shapes, a picture the user
// chose at the long-draw question, and the reads every assertion makes. Owns no
// assertions.

import type { ChartData } from "../src/link/chart-model";
import { TAG_LINK, type InboxItem, type LinkTag } from "../src/link/model";
import type { DrawChoice } from "../src/ppt/draw-progress";
import type { UpdateSummary } from "../src/ppt/links";
import { fakePng } from "./fakepng";
import type { FakePptShape } from "./fakeppt";
import { bootActions, type ActionsPane } from "./ppt.deck-actions.support";
import { columnChart, seedChart } from "./ppt.support";

export const PNG = fakePng(800, 400);
// 800 x 400 pixels are 600 x 300 points: the chart's own aspect is 1 to 2.
export const LABEL = "Model: B4:F12";
export const COLUMN = columnChart(6);
export const GREW = columnChart(7);

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export type UpgradePane = ActionsPane;

export const bootUpgrade = bootActions;

export function shapesOf(pane: UpgradePane): FakePptShape[] {
  return pane.presentation.slides[0]!.shapes;
}

export function boxOf(shape: FakePptShape): Box {
  return {
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

export function tagOf(shape: FakePptShape): LinkTag & { choice?: string } {
  return JSON.parse(shape.tags.get(TAG_LINK)!) as LinkTag;
}

// What a user does with an object on the slide: drags it to a corner and pulls
// it to a width.
export function moveTo(shape: FakePptShape, box: Box): void {
  Object.assign(shape, box);
}

export async function updateAll(pane: UpgradePane): Promise<UpdateSummary> {
  return pane.links.updateLinks(
    await pane.links.listLinks(pane.relay),
    pane.relay,
  );
}

// A chart link inserted before charts were drawn as shapes: a picture carrying
// the link's tags, whose payload then described no chart Excel could read.
export async function pictureEraChart(pane: UpgradePane): Promise<InboxItem> {
  const item = await seedChart(
    null,
    PNG,
    "this chart type has no data to read",
  );
  await pane.links.insertFromInbox(item, pane.ws, pane.relay);
  return item;
}

// Six series over forty points: minutes on the web, so the insert asks once.
export function bigChart(): ChartData {
  const base = columnChart(40);
  return {
    ...base,
    series: Array.from({ length: 6 }, (_, j) => ({
      ...base.series[0]!,
      name: `S${String(j)}`,
    })),
  };
}

// The picture the user chose when the question came: the web, a chart too long
// to draw, the answer "Insert as picture". The draw UI is the pane's own module
// graph's, the one bootUpgrade imported.
export async function chosenPicture(pane: UpgradePane): Promise<InboxItem> {
  const { installDrawUi } = await import("../src/ppt/draw-progress");
  pane.helpers.setPlatform("OfficeOnline");
  const answer: DrawChoice = "picture";
  installDrawUi({
    progress: () => undefined,
    ask: () => Promise.resolve(answer),
  });
  const item = await seedChart(bigChart(), PNG);
  await pane.links.insertFromInbox(item, pane.ws, pane.relay);
  return item;
}

export async function clearDrawUi(): Promise<void> {
  (await import("../src/ppt/draw-progress")).installDrawUi(null);
}
