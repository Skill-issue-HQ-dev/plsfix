// What the Linked objects list suites against the deck's own records share, on
// top of the pane harness (stress.ppt.support.ts): exports of a text and of a
// chart waiting in the pane's Inbox, a push of new text, the pane's Insert
// button pressed, the rows of the table read by kind, and the clock moved on.
// Owns no assertions.

import { vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import type { ChartData } from "../src/link/chart-model";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  sourceLabel,
  type InboxItem,
  type Payload,
} from "../src/link/model";
import { SCAN_SETTLE_MS } from "../src/ppt/scan-settle";
import { fakePng } from "./fakepng";
import {
  click,
  deck,
  linkRows,
  paneRelay,
  paneWorkspace,
  settle,
  SRC,
} from "./stress.ppt.support";

async function publish(
  id: string,
  token: string,
  payload: Payload,
): Promise<void> {
  const keys = await deriveLinkKeys(token);
  await paneRelay().putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
}

function textPayload(text: string, hash: string): Payload {
  return {
    v: 1,
    kind: "text",
    text,
    src: SRC,
    pushedAt: new Date().toISOString(),
    hash,
  };
}

async function announce(item: InboxItem): Promise<void> {
  const ws = paneWorkspace();
  await paneRelay().postInbox(
    ws.id,
    ws.auth,
    item.id,
    await seal(ws.enc, ws.id, encodeInboxItem(item)),
  );
}

// One text export waiting in the Inbox of the pane's workspace.
export async function waitingText(
  text: string,
  project?: string,
): Promise<InboxItem> {
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  const payload = textPayload(text, "0".repeat(64));
  await publish(id, token, payload);
  const item: InboxItem = {
    id,
    token,
    kind: "text",
    label: sourceLabel(SRC, "text"),
    src: SRC,
    createdAt: new Date().toISOString(),
    ...(project === undefined ? {} : { project }),
  };
  await announce(item);
  return item;
}

// One chart export waiting in the Inbox, its data and its picture both.
export async function waitingChart(chart: ChartData): Promise<InboxItem> {
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  await publish(id, token, {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png: fakePng(800, 400),
    src: SRC,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
    chart,
  });
  const item: InboxItem = {
    id,
    token,
    kind: "chart",
    label: sourceLabel(SRC, "chart"),
    src: SRC,
    createdAt: new Date().toISOString(),
  };
  await announce(item);
  return item;
}

// Excel pushing the text link again.
export async function pushedText(item: InboxItem, text: string): Promise<void> {
  await publish(item.id, item.token, textPayload(text, "1".repeat(64)));
}

// The export re-announced to the Inbox, as Excel's Push does.
export function reannounced(item: InboxItem): Promise<void> {
  return announce(item);
}

// The Inbox read and its first Insert pressed, the way a user inserts.
export async function insertFirstWaiting(): Promise<void> {
  click("refresh-inbox");
  await settle();
  document.querySelector<HTMLButtonElement>(".inbox-insert")!.click();
  await settle();
}

export function offSlideRows(): HTMLTableRowElement[] {
  return [
    ...document.querySelectorAll<HTMLTableRowElement>(
      "#link-rows tr.link-offslide",
    ),
  ];
}

export function deckButton(
  action: string,
  row = 0,
): HTMLButtonElement | undefined {
  return (
    offSlideRows()[row]?.querySelector<HTMLButtonElement>(
      `[data-action="${action}"]`,
    ) ?? undefined
  );
}

// The clock moved on, the way the minutes after a deck opens pass: Date and
// performance both, which are the clocks the settle reads.
export function afterBoot(ms: number): void {
  vi.advanceTimersByTime(ms);
}

// A text link inserted from the Inbox onto the first slide and its object then
// deleted by hand, the pane left open on the Links tab.
export async function deletedByHand(
  text = "EUR 3.2m",
  project?: string,
): Promise<InboxItem> {
  const item = await waitingText(text, project);
  await insertFirstWaiting();
  if (linkRows().length !== 1) throw new Error("test: no link row to delete");
  deck().deleteShape(deck().slides[0]!.shapes[0]!.id);
  return item;
}

// Two late scans a full interval apart, both the user's own refresh: the list as
// it reads once the deck can be believed (about 30 s after it opened).
export async function refreshedAfterTheSettle(): Promise<void> {
  for (let scan = 0; scan < 2; scan += 1) {
    afterBoot(SCAN_SETTLE_MS + 1_000);
    click("refresh-links");
    await settle();
  }
}
