// The style-cycle key: reads the list, the styles the workbook has and the active cell's style in
// one batch, takes the next style of the ring (src/custom-style-cycle.ts) and puts it on every
// area through the write Apply uses. Invariants: it never writes the list, only the cells; presses
// run one after the other, each reading the style the one before wrote.

import {
  cycleReceipt,
  listedNames,
  nextStyle,
  NO_STYLES_YET,
} from "../custom-style-cycle";
import { selectedAreas } from "./areas";
import { putStyleOnAreas } from "./custom-style-put";
import { queueRegistryRead, registryEntries } from "./custom-style-store";

const STAGE_CYCLE = "Style cycling";

async function step(): Promise<string> {
  return Excel.run(async (context) => {
    const styles = context.workbook.styles;
    styles.load("items/name");
    const registry = queueRegistryRead(context);
    const active = context.workbook.getActiveCell();
    active.load("style");
    // The areas' own sync carries the three reads above.
    const areas = await selectedAreas(context, STAGE_CYCLE);

    const names = listedNames(
      registryEntries(registry),
      styles.items.map((style) => style.name),
    );
    const next = nextStyle(names, active.style);
    if (next === null) throw new Error(NO_STYLES_YET);
    await putStyleOnAreas(context, areas, next.name, STAGE_CYCLE);
    return cycleReceipt(next);
  });
}

// Office fires a key's handler as it is pressed, so two presses a moment apart would both read
// the style the first had not written yet and make one step. A press waits for the one before,
// whether that one worked or was refused.
let queue: Promise<unknown> = Promise.resolve();

/**
 * Puts the style after the active cell's on the whole selection: the pls,fix
 * styles in list order, then Excel's Normal after the last, and the first from
 * any other style. Answers "Style: Margin % (2 of 5)" or "Style: Normal".
 */
export function cycleCustomStyle(): Promise<string> {
  const run = queue.then(step);
  queue = run.catch(() => undefined);
  return run;
}
