// @vitest-environment jsdom
// The shipped PowerPoint pane drawing a chart over the fake host: the progress
// line the toast shows while the shapes go on, replaced by the action's own
// final toast. Drives pptpane.html through src/ppt/main.ts (harness:
// stress.ppt.support). Strict load semantics are on.

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  NO_SHAPE_API,
  SMOOTH_STRAIGHT,
  type ChartData,
} from "../src/link/chart-model";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  sourceLabel,
  type InboxItem,
  type PicturePayload,
} from "../src/link/model";
import type * as RelayModule from "../src/link/relay";
import { fakePng } from "./fakepng";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { columnChart } from "./ppt.support";
import {
  bootPane,
  button,
  deck,
  hostHelpers,
  linkRows,
  paneRelay,
  paneWorkspace,
  pressEnterOnKeyField,
  settle,
  SRC,
  toastText,
} from "./stress.ppt.support";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return paneRelay();
    },
  };
});

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

const LABEL = sourceLabel(SRC, "chart");

// One chart export waiting in the Inbox for this workspace, the way Excel
// would have left it: the data the slide draws and the picture beside it - or,
// for a chart Excel could not describe, the picture and the reason why.
async function waitingChart(
  chart: ChartData | undefined,
  png = fakePng(800, 400),
  chartIssue?: string,
): Promise<InboxItem> {
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  const payload: PicturePayload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src: SRC,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
    ...(chart === undefined ? {} : { chart }),
    ...(chartIssue === undefined ? {} : { chartIssue }),
  };
  const keys = await deriveLinkKeys(token);
  await paneRelay().putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
  const ws = paneWorkspace();
  const item: InboxItem = {
    id,
    token,
    kind: "chart",
    label: LABEL,
    src: SRC,
    createdAt: payload.pushedAt,
  };
  await paneRelay().postInbox(
    ws.id,
    ws.auth,
    id,
    await seal(ws.enc, ws.id, encodeInboxItem(item)),
  );
  return item;
}

// Every text the toast showed, in order: a mutation observer reads the
// message at each change, so a line replaced by the next is still seen.
function watchToast(): { texts: string[]; stop: () => void } {
  const texts: string[] = [];
  const toast = document.getElementById("toast")!;
  const observer = new MutationObserver(() => {
    const text = toast.querySelector(".toast-text")?.textContent ?? "";
    if (text !== "" && texts.at(-1) !== text) texts.push(text);
  });
  observer.observe(toast, { childList: true, subtree: true });
  return { texts, stop: () => observer.disconnect() };
}

function insertButton(): HTMLButtonElement {
  return document.querySelector<HTMLButtonElement>(
    `button.inbox-insert[aria-label="Insert ${LABEL}"]`,
  )!;
}

async function openInbox(): Promise<void> {
  document.getElementById("refresh-inbox")!.click();
  await settle();
}

describe("the pane's toast while a chart is drawn", () => {
  it("says how far the draw is, then the action's own sentence", async () => {
    await bootPane();
    await waitingChart(columnChart(6));
    await openInbox();
    const seen = watchToast();

    insertButton().click();
    await settle();
    seen.stop();

    // The fake host is the Windows desktop until a test says otherwise, so no
    // tab sentence; twenty shapes are two chunks of twelve and eight.
    expect(seen.texts).toEqual([
      `Drawing ${LABEL}: 0 of 20 shapes.`,
      `Drawing ${LABEL}: 12 of 20 shapes.`,
      `Drawing ${LABEL}: 20 of 20 shapes.`,
      `Inserted ${LABEL}.`,
    ]);
    expect(toastText()).toBe(`Inserted ${LABEL}.`);
    expect(deck().slides[0]!.shapes[0]!.type).toBe("Group");
  });

  it("tells the web to keep the tab in front", async () => {
    await bootPane();
    hostHelpers().setPlatform("OfficeOnline");
    await waitingChart(columnChart(6));
    await openInbox();
    const seen = watchToast();

    insertButton().click();
    await settle();
    seen.stop();

    expect(seen.texts[0]).toBe(
      `Drawing ${LABEL}: 0 of 20 shapes. Keep this tab in front.`,
    );
    expect(seen.texts.at(-1)).toBe(`Inserted ${LABEL}.`);
  });
});

// A chart that lands as a picture (05.10.2026, real Office for the web: "Inserted
// Combo chart. as a picture: ..."), or that is drawn with something it could not
// keep, says so in the same sentence as its label, not after a full stop.
describe("the pane's toast for a chart that lands as a picture", () => {
  it("reads Excel's reason on from the label, as one sentence", async () => {
    await bootPane();
    await waitingChart(
      undefined,
      fakePng(800, 400),
      "this chart type is not drawn as shapes",
    );
    await openInbox();

    insertButton().click();
    await settle();

    expect(toastText()).toBe(
      `Inserted ${LABEL} as a picture: this chart type is not drawn as shapes.`,
    );
    expect(deck().slides[0]!.shapes[0]!.type).not.toBe("Group");
  });

  it("reads this host's own reason the same way", async () => {
    await bootPane();
    hostHelpers().setSupported((_set, version) => version !== "1.8");
    await waitingChart(columnChart(6));
    await openInbox();

    insertButton().click();
    await settle();

    expect(toastText()).toBe(`Inserted ${LABEL} ${NO_SHAPE_API}.`);
  });

  it("puts what a drawn line could not keep in brackets after the label", async () => {
    await bootPane();
    const base = columnChart(6);
    await waitingChart({
      ...base,
      kind: "line",
      series: [{ ...base.series[0]!, smooth: true }],
    });
    await openInbox();

    insertButton().click();
    await settle();

    expect(toastText()).toBe(`Inserted ${LABEL} (${SMOOTH_STRAIGHT}).`);
    expect(deck().slides[0]!.shapes[0]!.type).toBe("Group");
  });
});

// Six series over forty points: 240 bars and what goes with them, minutes of
// drawing on the web, so Insert asks before it starts.
function bigChart(): ChartData {
  const base = columnChart(40);
  return {
    ...base,
    series: Array.from({ length: 6 }, (_, j) => ({
      ...base.series[0]!,
      name: `S${String(j)}`,
    })),
  };
}

function questionButtons(): HTMLButtonElement[] {
  return [
    ...document.querySelectorAll<HTMLButtonElement>("#toast .toast-copy"),
  ];
}

// The pane is busy while the question is open, and has to say so to everything
// but the answer: the question is in the toast, outside the shell the latch
// disables.
async function openTheQuestion(): Promise<void> {
  await bootPane();
  hostHelpers().setPlatform("OfficeOnline");
  await waitingChart(bigChart());
  await openInbox();
  insertButton().click();
  await vi.waitFor(() => {
    expect(questionButtons()).toHaveLength(2);
  });
}

describe("the question in the pane's toast", () => {
  it("is two buttons the busy latch leaves alone, over a pane that is otherwise busy", async () => {
    await openTheQuestion();

    expect(toastText()).toMatch(
      new RegExp(
        `^${LABEL}: \\d+ shapes will take about \\d+ min to draw here\\.$`,
      ),
    );
    expect(questionButtons().map((one) => one.textContent)).toEqual([
      expect.stringMatching(/^Draw it \(about \d+ min\)$/),
      "Insert as picture",
    ]);
    expect(questionButtons().every((one) => !one.disabled)).toBe(true);
    expect(button("update-all").disabled).toBe(true);
    expect(insertButton().disabled).toBe(true);
    // Nothing is drawn until the user has answered.
    expect(deck().slides[0]!.shapes).toHaveLength(0);

    questionButtons()[1]!.click();
    await settle();
  });

  it("Insert as picture lands the picture and the toast says it was his choice", async () => {
    await openTheQuestion();

    questionButtons()[1]!.click();
    await settle();

    expect(toastText()).toBe(`Inserted ${LABEL}. Picture (your choice)`);
    const picture = deck().slides[0]!.shapes[0]!;
    expect(picture.type).not.toBe("Group");
    expect(picture.fillImage).toBeTruthy();
    expect(questionButtons()).toHaveLength(0);
    // The export left the Inbox: it was inserted.
    expect(document.querySelector("button.inbox-insert")).toBeNull();
    expect(button("update-all").disabled).toBe(false);
  });

  // The row in the deck's Linked objects list is read from the shape's tag:
  // the choice he made at the question is what the second line says.
  it("Insert as picture: the Linked objects row reads Picture (your choice), not Chart", async () => {
    await openTheQuestion();

    questionButtons()[1]!.click();
    await settle();

    const rows = linkRows();
    expect(rows).toHaveLength(1);
    expect(rows[0]!.querySelector(".link-meta")!.textContent).toBe(
      `${SRC.workbook} · Picture (your choice)`,
    );
  });

  it("Draw it: the Linked objects row still reads Chart", async () => {
    await openTheQuestion();

    questionButtons()[0]!.click();
    await settle();

    const rows = linkRows();
    expect(rows).toHaveLength(1);
    expect(rows[0]!.querySelector(".link-meta")!.textContent).toBe(
      `${SRC.workbook} · Chart`,
    );
  });

  it("Draw it draws the group, with the progress line, and the usual toast at the end", async () => {
    await openTheQuestion();
    const seen = watchToast();

    questionButtons()[0]!.click();
    await settle();
    seen.stop();

    expect(deck().slides[0]!.shapes[0]!.type).toBe("Group");
    expect(seen.texts[0]).toMatch(
      new RegExp(
        `^Drawing ${LABEL}: 0 of \\d+ shapes\\. Keep this tab in front\\.$`,
      ),
    );
    expect(toastText()).toBe(`Inserted ${LABEL}.`);
  });

  it("is dismissed by anything else that takes the toast: nothing is inserted, the export stays", async () => {
    await openTheQuestion();

    // A second action while the first waits: the guard answers it with the
    // busy sentence in the same toast, which answers the question first.
    pressEnterOnKeyField();
    await settle();

    expect(toastText()).toBe(
      "Cancelled: nothing was inserted, the export is still in the Inbox",
    );
    expect(deck().slides[0]!.shapes).toHaveLength(0);
    expect(insertButton()).not.toBeNull();
    expect(insertButton().disabled).toBe(false);
  });

  // Calling the insert off is not a failure: the sentence is a plain answer,
  // not a red error with a Copy details button to paste into a bug report.
  it("says the call-off as a plain success: no error class, no Copy details button", async () => {
    await openTheQuestion();

    pressEnterOnKeyField();
    await settle();

    const toast = document.getElementById("toast")!;
    expect(toast.className).toContain("visible");
    expect(toast.className).toContain("success");
    expect(toast.className).not.toContain("error");
    expect(questionButtons()).toHaveLength(0);
    expect(toast.querySelector("button")).toBeNull();
  });

  it("puts the keyboard on the first answer when the question opens", async () => {
    await openTheQuestion();

    expect(document.activeElement).toBe(questionButtons()[0]);

    questionButtons()[1]!.click();
    await settle();
  });

  it("takes Escape as the answer that cancels: a plain toast, nothing inserted, the export stays", async () => {
    await openTheQuestion();

    document.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Escape", bubbles: true }),
    );
    await settle();

    expect(toastText()).toBe(
      "Cancelled: nothing was inserted, the export is still in the Inbox",
    );
    expect(document.getElementById("toast")!.className).not.toContain("error");
    expect(questionButtons()).toHaveLength(0);
    expect(deck().slides[0]!.shapes).toHaveLength(0);
    expect(insertButton().disabled).toBe(false);
  });
});
