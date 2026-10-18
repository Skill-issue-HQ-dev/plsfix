// The virtual clocks of the suite: the debouncer's virtualClock as it always
// was, and costedClock, which the fake PowerPoint host charges for every
// shapes.add* so a test can prove a chunk size by the time a chunk took.

import { describe, expect, it } from "vitest";
import { costedClock, virtualClock } from "./clock";

describe("virtualClock", () => {
  it("moves only when told to and runs what falls due inside advance", () => {
    const clock = virtualClock();
    const ran: number[] = [];
    clock.after(100, () => ran.push(clock.now()));
    clock.after(300, () => ran.push(clock.now()));
    expect([clock.now(), clock.waiting()]).toEqual([0, 2]);
    clock.advance(150);
    expect(ran).toEqual([150]);
    clock.advance(200);
    expect(ran).toEqual([150, 350]);
    expect(clock.waiting()).toBe(0);
  });

  it("forgets a timer that was cancelled", () => {
    const clock = virtualClock();
    let ran = false;
    clock
      .after(10, () => {
        ran = true;
      })
      .cancel();
    clock.advance(50);
    expect(ran).toBe(false);
  });
});

describe("costedClock", () => {
  it("is free until it is priced, and is a virtualClock all the same", () => {
    const clock = costedClock();
    clock.charge(0);
    clock.charge(500);
    expect(clock.now()).toBe(0);
    expect(clock.charged()).toBe(2);
    const ran: number[] = [];
    clock.after(5, () => ran.push(clock.now()));
    clock.advance(5);
    expect(ran).toEqual([5]);
  });

  it("charges a flat price per add plus one per shape already on the slide", () => {
    const clock = costedClock({ msPerAdd: 100, msPerExistingShape: 10 });
    clock.charge(0);
    expect(clock.now()).toBe(100);
    clock.charge(3);
    expect(clock.now()).toBe(100 + 130);
    expect(clock.charged()).toBe(2);
  });

  it("can be repriced between adds", () => {
    const clock = costedClock({ msPerAdd: 100 });
    clock.charge(0);
    clock.costs.msPerAdd = 40;
    clock.costs.msPerExistingShape = 1;
    clock.charge(60);
    expect(clock.now()).toBe(100 + 40 + 60);
  });

  it("runs a timer that falls due inside an add, after the add", () => {
    const clock = costedClock({ msPerAdd: 100 });
    const ran: number[] = [];
    clock.after(60, () => ran.push(clock.now()));
    clock.charge(0);
    expect(ran).toEqual([100]);
  });
});
