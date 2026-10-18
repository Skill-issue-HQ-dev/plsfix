// @vitest-environment jsdom
// The PowerPoint pane's banner over the real pptpane.html and the fake deck: a
// failure has no timer and stands until the next pane action starts, even when
// that action has nothing to say (the ribbon's Object tools press only brings
// the pane up). The Excel pane's guard is held to the same rule in
// src/pane/shared.test.ts.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  type FakePptHelpers,
  installFakePpt,
  uninstallFakePpt,
} from "./fakeppt";
import { settlePpt, trackPptBoot } from "./ppt-ready";

async function boot(): Promise<FakePptHelpers> {
  vi.resetModules();
  uninstallFakePpt();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const { helpers } = installFakePpt({ slides: 1 });
  const booted = trackPptBoot();
  await import("../src/ppt/main");
  await booted;
  return helpers;
}

const banner = (): HTMLElement => document.getElementById("toast")!;
const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";

afterEach(() => {
  uninstallFakePpt();
});

describe("the PowerPoint pane's banner", () => {
  it("keeps a failure until the next action starts, then takes it off", async () => {
    const helpers = await boot();

    document.querySelector<HTMLButtonElement>("#revert-selected")?.click();
    await settlePpt(1);
    expect(toastText()).toBe("Tick the rows to revert.");
    expect(banner().className).toContain("error");

    await helpers.runCommand("PLSFIX_PPT_TOOLS");
    expect(banner().className).toBe("toast");
  });

  it("lets the next failure replace the last one rather than stack on it", async () => {
    await boot();

    document.querySelector<HTMLButtonElement>("#revert-selected")?.click();
    await settlePpt(1);
    document.querySelector<HTMLButtonElement>("#change-source-cancel")?.click();
    await settlePpt(1);

    expect(banner().querySelectorAll(".toast-text")).toHaveLength(1);
    expect(toastText()).toBe(
      "Change source cancelled. Nothing was re-pointed.",
    );
  });

  it("goes when it is clicked", async () => {
    await boot();

    document.querySelector<HTMLButtonElement>("#revert-selected")?.click();
    await settlePpt(1);
    banner()
      .querySelector(".toast-text")!
      .dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(banner().className).toBe("toast");
  });
});
