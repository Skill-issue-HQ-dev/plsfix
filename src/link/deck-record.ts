// The deck's own copy of one link: the record a PowerPoint custom XML part
// holds (the link's data minus its picture bytes), its pure codec and the caps
// a deck keeps them under. Invariant: decode trusts nothing - a part it cannot
// read in full is refused with a reason, never half used.

import type { ChartData } from "./chart-model";
import { fromBase64Url, toBase64Url } from "./crypto";
import {
  decodePayload,
  encodePayload,
  isKind,
  isLinkId,
  isSource,
  type LinkId,
  type LinkKind,
  type Payload,
  type Source,
  type TablePayload,
  type TextPayload,
} from "./model";
import { optionalProject, readProject } from "./project";

// The namespace every part of the deck store lives in. getByNamespace needs
// the fully qualified URI; the trailing 1 is the record format's own version.
export const DECK_NAMESPACE = "urn:plsfix:deck-link:1";
export const DECK_RECORD_VERSION = 1;

// The two caps, in one place. S1 (the measuring slice) may refine both once
// real decks have been measured; nothing else in the code names a byte count.
export const DECK_RECORD_MAX_BYTES = 256 * 1024;
export const DECK_STORE_MAX_BYTES = 4 * 1024 * 1024;

// What a host may add when it hands a part back (a declaration, a mark, white
// space): reading allows this much past the cap a record is written under.
const READ_SLACK_BYTES = 1024;
const ELEMENT = "link";

// What a chart keeps: everything a slide needs to draw it again.
export interface ChartKept {
  v: 1;
  kind: "picture";
  width: number;
  height: number;
  hash: string;
  chart: ChartData;
}

// The payload a record keeps: the decoded Payload without its picture bytes,
// and without what the record already says itself (src, pushedAt).
export type DeckPayload =
  | Omit<TablePayload, "src" | "pushedAt">
  | Omit<TextPayload, "src" | "pushedAt">
  | ChartKept;

export interface DeckRecord {
  v: 1;
  id: LinkId;
  kind: LinkKind;
  label: string;
  project?: string;
  src: Source;
  token: string;
  rev: number;
  // When Excel first sent the link to this deck (the inbox row's createdAt).
  sentAt: string;
  // When Excel pushed the revision this record holds (the payload's own).
  pushedAt: string;
  // When this pane wrote the record into the deck.
  savedAt: string;
  // False when the record keeps identity only: a picture link, or a link too
  // large for the caps. `payload` is null exactly then.
  dataKept: boolean;
  payload: DeckPayload | null;
}

export interface DeckLinkIdentity {
  id: LinkId;
  kind: LinkKind;
  label: string;
  project?: string;
  token: string;
  sentAt: string;
}

export type DeckRefusal =
  | "notXml"
  | "tooBig"
  | "wrongNamespace"
  | "notALink"
  | "newerVersion"
  | "badId"
  | "badBody"
  | "badRecord";

export type DeckDecode =
  | { ok: true; record: DeckRecord }
  | { ok: false; reason: DeckRefusal; message: string };

export const NEWER_DECK_RECORD =
  "This deck's link data was saved by a newer pls,fix. Update the add-in.";

const MESSAGES: Record<DeckRefusal, string> = {
  notXml: "That part is not one XML element.",
  tooBig: "That part is larger than any link record.",
  wrongNamespace: "That part is in another namespace.",
  notALink: "That part is not a link record.",
  newerVersion: NEWER_DECK_RECORD,
  badId: "That link record does not carry a valid link id.",
  badBody: "That link record's body cannot be read.",
  badRecord: "That link record is not in a shape this pls,fix knows.",
};

// What a payload keeps in a record, field by field: a key this build does not
// know is dropped, so what is read back is exactly what this build wrote.
export function deckPayloadOf(payload: Payload): DeckPayload | null {
  if (payload.kind === "table") {
    const { rows, cols, cells, widths, hash } = payload;
    const header = payload.h === true ? { h: true as const } : {};
    return { v: 1, kind: "table", rows, cols, cells, widths, hash, ...header };
  }
  if (payload.kind === "text") {
    return { v: 1, kind: "text", text: payload.text, hash: payload.hash };
  }
  if (payload.chart === undefined) return null;
  const { width, height, hash, chart } = payload;
  return { v: 1, kind: "picture", width, height, hash, chart };
}

// The payload a record stands for, as the insert and repaint code takes it. A
// chart's picture bytes are not kept, so its png is empty.
export function payloadFromRecord(record: DeckRecord): Payload | null {
  const kept = record.payload;
  if (kept === null) return null;
  const identity = { src: record.src, pushedAt: record.pushedAt };
  if (kept.kind === "picture") {
    return { ...kept, ...identity, mime: "image/png", png: "" };
  }
  return { ...kept, ...identity };
}

export function buildDeckRecord(
  link: DeckLinkIdentity,
  payload: Payload,
  rev: number,
  savedAt: string,
): DeckRecord {
  const kept = deckPayloadOf(payload);
  const project = readProject(link.project);
  return {
    v: DECK_RECORD_VERSION,
    id: link.id,
    kind: link.kind,
    label: link.label,
    ...(project === undefined ? {} : { project }),
    src: payload.src,
    token: link.token,
    rev,
    sentAt: link.sentAt,
    pushedAt: payload.pushedAt,
    savedAt,
    dataKept: kept !== null,
    payload: kept,
  };
}

export function identityOnly(record: DeckRecord): DeckRecord {
  return { ...record, dataKept: false, payload: null };
}

function serialize(record: DeckRecord): string {
  if (!isLinkId(record.id)) {
    throw new Error("encodeDeckRecord: the record's id is not a link id");
  }
  const json = JSON.stringify(record);
  const body = toBase64Url(new TextEncoder().encode(json));
  return (
    `<${ELEMENT} xmlns="${DECK_NAMESPACE}" v="${String(DECK_RECORD_VERSION)}" ` +
    `id="${record.id}">${body}</${ELEMENT}>`
  );
}

// The XML a record is stored as, and the record that XML says: over the cap, a
// record is cut to identity only (never refused), so the deck still knows the
// link even where it cannot keep its data.
export function packDeckRecord(record: DeckRecord): {
  xml: string;
  record: DeckRecord;
} {
  const xml = serialize(record);
  if (xml.length <= DECK_RECORD_MAX_BYTES) return { xml, record };
  const lean = identityOnly(record);
  return { xml: serialize(lean), record: lean };
}

export function encodeDeckRecord(record: DeckRecord): string {
  return packDeckRecord(record).xml;
}

function refuse(reason: DeckRefusal): DeckDecode {
  return { ok: false, reason, message: MESSAGES[reason] };
}

// One element, whatever a host may have done to the text around it: an XML
// declaration or a mark in front, white space, either quote, any attribute
// order. Anything with more in it than that is not what this build wrote.
const ELEMENT_TEXT =
  /^\s*(?:<\?xml[^>]*\?>\s*)?<([A-Za-z_][\w.:-]*)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*>([^<]*)<\/([A-Za-z_][\w.:-]*)\s*>\s*$/;
const ATTRIBUTE = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

interface Parsed {
  name: string;
  attrs: Map<string, string>;
  body: string;
}

function parseElement(xml: string): Parsed | null {
  const match = ELEMENT_TEXT.exec(xml.startsWith("﻿") ? xml.slice(1) : xml);
  if (match === null || match[1] !== match[4]) return null;
  const attrs = new Map<string, string>();
  for (const found of (match[2] ?? "").matchAll(ATTRIBUTE)) {
    if (attrs.has(found[1]!)) return null;
    attrs.set(found[1]!, found[2] ?? found[3] ?? "");
  }
  return { name: match[1]!, attrs, body: (match[3] ?? "").trim() };
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

// The body: base64url of UTF-8 JSON, every step of which can fail.
function readBody(body: string): Record<string, unknown> | null {
  if (body === "") return null;
  try {
    const bytes = fromBase64Url(body);
    const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    const json: unknown = JSON.parse(text);
    return isObject(json) ? json : null;
  } catch {
    return null;
  }
}

// What the version of a part says: the one this build reads, a newer one, or
// nothing it can use at all.
function versionOf(value: unknown): "this" | "newer" | "unreadable" {
  if (typeof value === "string" && /^[0-9]+$/.test(value))
    value = Number(value);
  if (typeof value !== "number" || !Number.isInteger(value))
    return "unreadable";
  if (value > DECK_RECORD_VERSION) return "newer";
  return value === DECK_RECORD_VERSION ? "this" : "unreadable";
}

// The payload is judged by the same validator every relay payload passes: it is
// put back together as the Payload it was cut from and read as one.
function readPayload(
  raw: unknown,
  src: Source,
  pushedAt: string,
): DeckPayload | null {
  if (!isObject(raw)) return null;
  const picture = raw.kind === "picture" ? { mime: "image/png", png: "" } : {};
  const whole = { ...raw, ...picture, src, pushedAt };
  try {
    return deckPayloadOf(
      decodePayload(encodePayload(whole as unknown as Payload)),
    );
  } catch {
    return null;
  }
}

function isTime(value: unknown): value is string {
  return typeof value === "string";
}

function readRecord(json: Record<string, unknown>): DeckRecord | null {
  const { id, kind, label, token, src, rev, dataKept } = json;
  if (typeof id !== "string" || !isKind(kind) || !isSource(src)) return null;
  if (typeof label !== "string" || typeof token !== "string") return null;
  if (typeof rev !== "number" || !Number.isInteger(rev) || rev < 0) return null;
  const { sentAt, pushedAt, savedAt, project } = json;
  if (!isTime(sentAt) || !isTime(pushedAt) || !isTime(savedAt)) return null;
  if (!optionalProject(project) || typeof dataKept !== "boolean") return null;
  const kept = dataKept ? readPayload(json.payload, src, pushedAt) : null;
  // Data claimed but unreadable, or a payload where none was claimed.
  if (dataKept && kept === null) return null;
  if (!dataKept && json.payload !== null) return null;
  const named = readProject(project);
  return {
    v: DECK_RECORD_VERSION,
    id,
    kind,
    label,
    ...(named === undefined ? {} : { project: named }),
    src,
    token,
    rev,
    sentAt,
    pushedAt,
    savedAt,
    dataKept,
    payload: kept,
  };
}

// Strict, in the order a refusal is most useful: what the part is, which
// version it is (before its shape is judged, because a newer part has another
// shape), whose it is, then every field of its body.
export function decodeDeckRecord(xml: string): DeckDecode {
  if (typeof xml !== "string") return refuse("notXml");
  if (xml.length > DECK_RECORD_MAX_BYTES + READ_SLACK_BYTES) {
    return refuse("tooBig");
  }
  const parsed = parseElement(xml);
  if (parsed === null) return refuse("notXml");
  if (parsed.attrs.get("xmlns") !== DECK_NAMESPACE) {
    return refuse("wrongNamespace");
  }
  if (parsed.name !== ELEMENT) return refuse("notALink");
  const version = versionOf(parsed.attrs.get("v"));
  if (version === "newer") return refuse("newerVersion");
  if (version === "unreadable") return refuse("badRecord");
  const id = parsed.attrs.get("id");
  if (id === undefined || !isLinkId(id)) return refuse("badId");
  const json = readBody(parsed.body);
  if (json === null) return refuse("badBody");
  const inside = versionOf(json.v);
  if (inside === "newer") return refuse("newerVersion");
  if (inside === "unreadable") return refuse("badRecord");
  if (json.id !== id) return refuse("badId");
  const record = readRecord(json);
  return record === null ? refuse("badRecord") : { ok: true, record };
}
