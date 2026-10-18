// A test knob for the fake host: the host refuses a style property write. `after`
// more writes land, then the next one throws, so the writes a caller queues behind
// it never run, the way a batch stops at the first statement the host refuses.
// Every style property write goes through StyleProxy.touched, which checks it.

import type { FakeWorkbook } from "./fakehost";

const refusing = new WeakMap<FakeWorkbook, { after: number; error: Error }>();

export function refuseStyleWriteAfter(
  workbook: FakeWorkbook,
  after: number,
  error: Error,
): void {
  refusing.set(workbook, { after, error });
}

/** Called by every style write: throws the armed error once its turn has come. */
export function checkRefusal(workbook: FakeWorkbook): void {
  const armed = refusing.get(workbook);
  if (!armed) return;
  if (armed.after > 0) {
    armed.after -= 1;
    return;
  }
  refusing.delete(workbook);
  throw armed.error;
}
