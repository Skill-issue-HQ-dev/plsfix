// A deck as it opens with the deck store already in it: one text link whose
// record is in the file, its object on a slide of the file or not, and that
// slide's reads held (PowerPoint for the web not having loaded it yet) or not.
// The pane boots over it. Owns no assertions.

import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import { buildDeckRecord, encodeDeckRecord } from "../src/link/deck-record";
import {
  encodePayload,
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type LinkTag,
  type TextPayload,
} from "../src/link/model";
import { seedParts } from "./ppt.deck-store.support";
import {
  bootPane,
  deck,
  hostHelpers,
  paneRelay,
  SRC,
} from "./stress.ppt.support";

export interface OpenedDeck {
  id: string;
  token: string;
}

export interface OpenOptions {
  // The slide (0-based) the link's object sits on; left out, the object is gone.
  objectOn?: number;
  // That slide reads as empty until it is shown.
  hold?: boolean;
}

// The link as the relay holds it, and its record as the file holds it.
async function seedRecord(
  id: string,
  token: string,
  payload: TextPayload,
): Promise<void> {
  const keys = await deriveLinkKeys(token);
  await paneRelay().putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
  const identity = {
    id,
    kind: "text" as const,
    label: "Model!B4:F12 text",
    token,
    sentAt: payload.pushedAt,
  };
  await seedParts([
    encodeDeckRecord(buildDeckRecord(identity, payload, 1, payload.pushedAt)),
  ]);
}

// The link's object on a slide of the file, tagged like an insert would have.
function plantObject(
  options: OpenOptions & { objectOn: number },
  id: string,
  token: string,
  payload: TextPayload,
): void {
  const slide = deck().slides[options.objectOn]!;
  const shape = deck().addShape(slide, {
    type: "TextBox",
    text: payload.text,
    left: 100,
    top: 100,
    width: 200,
    height: 40,
  });
  const tag: LinkTag = {
    v: 1,
    id,
    kind: "text",
    rev: 1,
    src: SRC,
    pushedAt: payload.pushedAt,
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, token);
  if (options.hold) hostHelpers().holdSlideReads(slide.id);
}

export async function openWithRecord(
  options: OpenOptions = {},
): Promise<OpenedDeck> {
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  const payload: TextPayload = {
    v: 1,
    kind: "text",
    text: "EUR 3.2m",
    src: SRC,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
  await bootPane(true, async () => {
    await seedRecord(id, token, payload);
    if (options.objectOn === undefined) return;
    plantObject({ ...options, objectOn: options.objectOn }, id, token, payload);
  });
  return { id, token };
}

// A second link's object on a slide, tagged and with no record: what a scan
// that has seen more than the one before it finds.
export function plantLate(slideIndex: number): void {
  const shape = deck().addShape(deck().slides[slideIndex]!, {
    type: "TextBox",
    text: "late",
    left: 10,
    top: 10,
    width: 100,
    height: 20,
  });
  const tag: LinkTag = {
    v: 1,
    id: "f".repeat(32),
    kind: "text",
    rev: 1,
    src: SRC,
    pushedAt: new Date().toISOString(),
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, newToken());
}
