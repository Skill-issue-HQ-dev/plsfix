// Guard: the one place every pane action funnels through - busy state on,
// run the action, refresh, decorate and notify on success, describe and
// notify on failure, then always clear busy. Host-agnostic: Office.js only
// reaches this module through the callbacks the caller supplies.
// Invariant: with `busyMessage` set this is also the pane's re-entrancy latch
// - one action at a time, the second answered rather than interleaved.

import type { ToastKind } from "./toast";

export interface GuardDeps {
  setBusy(busy: boolean): void;
  // Runs as each action starts, before its work and not for a press the latch
  // refuses: the pane takes the last answer (a standing error) off the banner.
  begin?(): void;
  notify(message: string, kind?: ToastKind, details?: string): void;
  describe(
    error: unknown,
    action: string | undefined,
  ): { message: string; details: string };
  after?(): Promise<void>;
  decorate?(message: string): string;
  finally?(): void;
  // The sentence a second action gets while the first is still in the host.
  // Absent (the Excel pane) leaves the latch off and nothing changes: every
  // caller runs the moment it is called, as it always did.
  busyMessage?: string;
}

export type Guard = (
  run: () => Promise<string>,
  action?: string,
) => Promise<void>;

// An action the user called off, in his own words: not a failure. The guard
// shows its message as an ordinary success toast, with no error look and no
// Copy details, and clears busy like after any other action.
export class CalledOff extends Error {
  constructor(message: string) {
    super(message);
    this.name = "CalledOff";
  }
}

export function makeGuard(deps: GuardDeps): Guard {
  const busyMessage = deps.busyMessage;
  let running = false;
  return async (run, action) => {
    if (busyMessage !== undefined && running) {
      deps.notify(busyMessage, "error");
      return;
    }
    running = true;
    deps.begin?.();
    deps.setBusy(true);
    try {
      const message = await run();
      await deps.after?.();
      deps.notify(deps.decorate ? deps.decorate(message) : message);
    } catch (error) {
      if (error instanceof CalledOff) {
        // "" and not none: a pane that falls back to its own buffered details
        // when the guard passes none must not attach them to a plain answer.
        deps.notify(error.message, "success", "");
      } else {
        const { message, details } = deps.describe(error, action);
        deps.notify(message, "error", details);
      }
    } finally {
      deps.finally?.();
      deps.setBusy(false);
      running = false;
    }
  };
}
