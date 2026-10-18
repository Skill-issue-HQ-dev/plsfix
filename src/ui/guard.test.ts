import { describe, expect, it, vi } from "vitest";
import { CalledOff, makeGuard } from "./guard";

describe("makeGuard", () => {
  it("busy on, run, after, decorate, notify, finally, busy off", async () => {
    const calls: string[] = [];
    const guard = makeGuard({
      setBusy: (b) => calls.push(`busy:${b}`),
      notify: (m, k) => calls.push(`notify:${k ?? "success"}:${m}`),
      describe: (e) => ({ message: String(e), details: "" }),
      after: async () => {
        calls.push("after");
      },
      decorate: (m) => `${m}!`,
      finally: () => calls.push("finally"),
    });
    await guard(async () => "Done");
    expect(calls).toEqual([
      "busy:true",
      "after",
      "notify:success:Done!",
      "finally",
      "busy:false",
    ]);
  });
  it("routes a failure to notify(error) with details and still clears busy", async () => {
    const notify = vi.fn();
    const guard = makeGuard({
      setBusy: () => undefined,
      notify,
      describe: (e, action) => ({
        message: (e as Error).message,
        details: `action=${action}`,
      }),
    });
    await guard(async () => {
      throw new Error("boom");
    }, "export");
    expect(notify).toHaveBeenCalledWith("boom", "error", "action=export");
  });

  it("runs one action at a time once busyMessage is set", async () => {
    const notify = vi.fn();
    const guard = makeGuard({
      setBusy: () => undefined,
      notify,
      describe: (e) => ({ message: String(e), details: "" }),
      busyMessage: "Wait for the last action to finish.",
    });
    let release = (): void => undefined;
    const held = new Promise<void>((done) => {
      release = () => done();
    });
    const second = vi.fn(async () => "second");

    const first = guard(async () => {
      await held;
      return "first";
    });
    await guard(second, "update-all");

    expect(second).not.toHaveBeenCalled();
    expect(notify).toHaveBeenCalledWith(
      "Wait for the last action to finish.",
      "error",
    );
    release();
    await first;
    // The latch clears with the flow: the next press runs.
    await guard(second);
    expect(second).toHaveBeenCalledTimes(1);
  });

  it("leaves the latch off for a guard with no busyMessage", async () => {
    const running: string[] = [];
    const guard = makeGuard({
      setBusy: () => undefined,
      notify: () => undefined,
      describe: (e) => ({ message: String(e), details: "" }),
    });
    let release = (): void => undefined;
    const held = new Promise<void>((done) => {
      release = () => done();
    });

    const first = guard(async () => {
      await held;
      return "first";
    });
    await guard(async () => {
      running.push("second");
      return "second";
    });

    // The Excel pane's guard is unchanged: the second action still runs.
    expect(running).toEqual(["second"]);
    release();
    await first;
  });

  it("notifies an action the user called off as a plain success: no error, no details", async () => {
    const notify = vi.fn();
    const describe = vi.fn();
    const guard = makeGuard({ setBusy: () => undefined, notify, describe });

    await guard(async () => {
      throw new CalledOff("Cancelled: nothing was inserted");
    }, "insert");

    // An empty details string, not none: a pane that falls back to its own
    // buffered details when the guard passes none must not attach them here.
    expect(notify).toHaveBeenCalledWith(
      "Cancelled: nothing was inserted",
      "success",
      "",
    );
    expect(describe).not.toHaveBeenCalled();
  });

  it("clears busy and the latch after a called-off action like after any other", async () => {
    const calls: string[] = [];
    const guard = makeGuard({
      setBusy: (busy) => calls.push(`busy:${String(busy)}`),
      notify: () => undefined,
      describe: (e) => ({ message: String(e), details: "" }),
      finally: () => calls.push("finally"),
      busyMessage: "Wait.",
    });

    await guard(async () => {
      throw new CalledOff("Cancelled");
    });
    const next = vi.fn(async () => "next");
    await guard(next);

    expect(calls.slice(0, 3)).toEqual(["busy:true", "finally", "busy:false"]);
    expect(next).toHaveBeenCalledTimes(1);
  });

  it("starts every action with begin, before the work runs: the next action takes the last answer off the banner", async () => {
    const calls: string[] = [];
    const guard = makeGuard({
      setBusy: (busy) => calls.push(`busy:${String(busy)}`),
      begin: () => calls.push("begin"),
      notify: (message) => calls.push(`notify:${message}`),
      describe: (e) => ({ message: String(e), details: "" }),
    });

    await guard(async () => {
      calls.push("run");
      return "Done";
    });
    await guard(async () => {
      throw new Error("boom");
    });

    expect(calls).toEqual([
      "begin",
      "busy:true",
      "run",
      "notify:Done",
      "busy:false",
      "begin",
      "busy:true",
      "notify:Error: boom",
      "busy:false",
    ]);
  });

  it("does not call begin for a press the busy latch refuses", async () => {
    const begin = vi.fn();
    const guard = makeGuard({
      setBusy: () => undefined,
      begin,
      notify: () => undefined,
      describe: (e) => ({ message: String(e), details: "" }),
      busyMessage: "Wait for the last action to finish.",
    });
    let release = (): void => undefined;
    const held = new Promise<void>((done) => {
      release = () => done();
    });
    const first = guard(async () => {
      await held;
      return "first";
    });
    await guard(async () => "second");

    expect(begin).toHaveBeenCalledTimes(1);
    release();
    await first;
  });

  it("is an Error that says which kind it is, and nothing else is called off", () => {
    const off = new CalledOff("Cancelled");
    expect(off).toBeInstanceOf(Error);
    expect(off.name).toBe("CalledOff");
    expect(off.message).toBe("Cancelled");
    expect(new Error("Cancelled")).not.toBeInstanceOf(CalledOff);
  });
});
