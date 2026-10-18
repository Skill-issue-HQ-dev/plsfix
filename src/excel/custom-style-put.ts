// Puts a cell style on every area of the selection: the one write path a row's Apply and the
// style-cycle key share. pls,fix Undo is captured first, each area is written, and the one synced
// write answers a protected sheet in words (a refused write spends no Undo slot).

import { protectedNote, syncWrite } from "./protection";
import { captureUndoAreas } from "./undo";

export async function putStyleOnAreas(
  context: Excel.RequestContext,
  areas: Excel.Range[],
  name: string,
  stage: string,
): Promise<void> {
  await captureUndoAreas(context, areas);
  for (const area of areas) area.style = name;
  await syncWrite(context, stage, protectedNote, areas.length);
}
