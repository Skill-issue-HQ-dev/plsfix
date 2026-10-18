// PowerPointApi 1.7 custom XML parts of the presentation, a slide and a shape:
// the part lists (off the model), the XML check and the proxies. Writes are
// journaled (journal.ts), so a failed or swallowed sync takes the batch back. A
// part that is not there is answered at the sync, as office.js does, never at
// the call. Invariant: a part is one well-formed element, in its namespace.

import { createRequire } from "node:module";
import { journalOnConfirm, journalRecord } from "./journal";
import { gone, invalidArgument, type FakePresentation } from "./model";
import { FakeClientResult, Loadable } from "./strict";

// Who holds a list of parts. A copy of a shape does not carry its parts: what
// PowerPoint does on a copy is for the L6 spike to measure.
export type XmlOwner =
  | { kind: "presentation" }
  | { kind: "slide"; id: string }
  | { kind: "shape"; id: string };

interface XmlPart {
  id: string;
  xml: string;
  namespace: string;
}

// What a test reads back through customXmlSnapshot, never the live part.
export interface CustomXmlSnapshot {
  id: string;
  namespaceUri: string;
  xml: string;
}

type Lists = Map<string, XmlPart[]>;

// `lists` is the deck as the batch in flight sees it, its own writes in; the
// host only holds them once a sync lands, which is when `committed` catches up.
interface Store {
  lists: Lists;
  committed: Lists;
  seq: number;
}

const stores = new WeakMap<FakePresentation, Store>();

function copyOf(lists: Lists): Lists {
  const copy: Lists = new Map();
  for (const [key, list] of lists) {
    copy.set(
      key,
      list.map((part) => ({ ...part })),
    );
  }
  return copy;
}

function storeOf(deck: FakePresentation): Store {
  let store = stores.get(deck);
  if (!store) {
    const made: Store = { lists: new Map(), committed: new Map(), seq: 0 };
    journalOnConfirm(deck, () => {
      made.committed = copyOf(made.lists);
    });
    stores.set(deck, made);
    store = made;
  }
  return store;
}

function keyOf(owner: XmlOwner): string {
  return owner.kind === "presentation"
    ? "presentation"
    : `${owner.kind}:${owner.id}`;
}

// The ids Office hands out are braced GUIDs; these count, so a test can tell
// them apart and a run reads the same twice.
function guid(seq: number): string {
  const tail = seq.toString(16).toUpperCase().padStart(12, "0");
  return `{C0575A00-0000-4000-8000-${tail}}`;
}

function generalException(why: string): Error {
  const error = new Error(`GeneralException: ${why}`);
  return Object.assign(error, { code: "GeneralException" });
}

// jsdom writes this root element, in this namespace, for text it cannot parse.
const PARSER_ERROR_NAMESPACE =
  "http://www.mozilla.org/newlayout/xml/parsererror.xml";
let nodeParser: typeof DOMParser | undefined;

function loadNodeParser(): typeof DOMParser {
  const { JSDOM } = createRequire(import.meta.url)("jsdom") as {
    JSDOM: new (html: string) => { window: { DOMParser: typeof DOMParser } };
  };
  return new JSDOM("").window.DOMParser;
}

// A real XML parser either way: the DOM a jsdom test already has, or in a
// plain node test a jsdom of the fake's own, so what is refused here is what a
// host's parser refuses and not what a hand-made check happened to catch.
function xmlParser(): DOMParser {
  if (typeof DOMParser !== "undefined") return new DOMParser();
  nodeParser ??= loadNodeParser();
  return new nodeParser();
}

// The namespace of the one element `xml` is - "" when it has none - or the
// GeneralException the host answers a fragment that is not that.
function namespaceOf(xml: unknown): string {
  if (typeof xml !== "string") throw generalException("the XML must be text");
  const doc = xmlParser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  if (
    root.localName === "parsererror" &&
    root.namespaceURI === PARSER_ERROR_NAMESPACE
  ) {
    throw generalException("the XML is not one well-formed element");
  }
  return root.namespaceURI ?? "";
}

// The owner's live list. An owner that has gone says ItemNotFound, like every
// other proxy over a deleted shape or slide.
function listOf(deck: FakePresentation, owner: XmlOwner): XmlPart[] {
  if (owner.kind === "slide") deck.findSlideOrThrow(owner.id);
  if (owner.kind === "shape") deck.findShape(owner.id);
  const { lists } = storeOf(deck);
  const list = lists.get(keyOf(owner)) ?? [];
  lists.set(keyOf(owner), list);
  return list;
}

function findPart(list: XmlPart[], id: string): XmlPart | undefined {
  return list.find((part) => part.id === id);
}

function partOrThrow(list: XmlPart[], id: string): XmlPart {
  const part = findPart(list, id);
  if (!part) throw gone("custom XML part", id);
  return part;
}

function addPart(
  deck: FakePresentation,
  owner: XmlOwner,
  xml: unknown,
): XmlPart {
  const namespace = namespaceOf(xml);
  const list = listOf(deck, owner);
  const store = storeOf(deck);
  store.seq += 1;
  const part = { id: guid(store.seq), xml: xml as string, namespace };
  list.push(part);
  journalRecord(deck, () => {
    const at = list.indexOf(part);
    if (at >= 0) list.splice(at, 1);
  });
  return part;
}

function rewrite(
  deck: FakePresentation,
  owner: XmlOwner,
  id: string,
  xml: unknown,
): void {
  const namespace = namespaceOf(xml);
  const part = partOrThrow(listOf(deck, owner), id);
  const before = { xml: part.xml, namespace: part.namespace };
  Object.assign(part, { xml, namespace });
  journalRecord(deck, () => {
    Object.assign(part, before);
  });
}

function remove(deck: FakePresentation, owner: XmlOwner, id: string): void {
  const list = listOf(deck, owner);
  const part = partOrThrow(list, id);
  const at = list.indexOf(part);
  list.splice(at, 1);
  journalRecord(deck, () => {
    list.splice(at, 0, part);
  });
}

// PowerPoint.CustomXmlPart. `scope` is the namespace of the scoped collection
// it came out of: a part that left it is no longer found through it.
export class CustomXmlPartProxy extends Loadable {
  constructor(
    private deck: FakePresentation,
    private owner: XmlOwner,
    private partId: string,
    private nullable = false,
    private scope?: string,
  ) {
    super();
  }

  get id(): string {
    return this.nullable && !this.peek() ? "" : this.part().id;
  }
  get namespaceUri(): string {
    return this.nullable && !this.peek() ? "" : this.part().namespace;
  }
  get isNullObject(): boolean {
    return this.peek() === undefined;
  }
  getXml(): FakeClientResult<string> {
    return new FakeClientResult(this.present()?.xml ?? "");
  }
  setXml(xml: string): void {
    if (this.present()) rewrite(this.deck, this.owner, this.partId, xml);
  }
  delete(): void {
    if (this.present()) remove(this.deck, this.owner, this.partId);
  }

  // office.js finds out that an id is not there when the batch runs: the call
  // is queued and the batch halts at this op, which the sync says. A null
  // object is the one thing refused at the call.
  private present(): XmlPart | undefined {
    const part = this.peek();
    if (part) return part;
    const error = gone("custom XML part", this.partId);
    if (this.nullable) throw error;
    this.deck.haltAtSync(error);
    return undefined;
  }

  private peek(): XmlPart | undefined {
    const part = findPart(listOf(this.deck, this.owner), this.partId);
    return part && (this.scope === undefined || part.namespace === this.scope)
      ? part
      : undefined;
  }
  private part(): XmlPart {
    const part = this.peek();
    if (!part) throw gone("custom XML part", this.partId);
    return part;
  }
}

// What the whole collection and the scoped one share: the list they read, its
// count, items and the three ways of reaching one part.
abstract class PartsRead extends Loadable {
  constructor(
    protected deck: FakePresentation,
    protected owner: XmlOwner,
    protected scope?: string,
  ) {
    super();
  }

  protected abstract parts(): XmlPart[];

  get items(): CustomXmlPartProxy[] {
    return this.parts().map((part) => this.item(part.id));
  }
  getCount(): FakeClientResult<number> {
    return new FakeClientResult(this.parts().length);
  }
  getItem(id: string): CustomXmlPartProxy {
    return this.item(id);
  }
  // PowerPointApi 1.10: InvalidArgument past either end.
  getItemAt(index: number): CustomXmlPartProxy {
    const part = this.parts()[index];
    if (!part) throw invalidArgument(`no custom XML part at ${String(index)}`);
    return this.item(part.id);
  }
  getItemOrNullObject(id: string): CustomXmlPartProxy {
    return this.item(id, true);
  }

  protected item(id: string, nullable = false): CustomXmlPartProxy {
    return new CustomXmlPartProxy(
      this.deck,
      this.owner,
      id,
      nullable,
      this.scope,
    );
  }
}

// PowerPoint.CustomXmlPartScopedCollection: what getByNamespace answers.
export class ScopedCustomXmlPartCollectionProxy extends PartsRead {
  protected parts(): XmlPart[] {
    return listOf(this.deck, this.owner).filter(
      (part) => part.namespace === this.scope,
    );
  }

  // Exactly one part, or the GeneralException the host answers.
  getOnlyItem(): CustomXmlPartProxy {
    const [only, ...rest] = this.parts();
    if (!only || rest.length > 0) {
      throw generalException("the collection does not hold exactly one part");
    }
    return this.item(only.id);
  }
  getOnlyItemOrNullObject(): CustomXmlPartProxy {
    const parts = this.parts();
    return this.item(parts.length === 1 ? parts[0]!.id : "", true);
  }
}

// PowerPoint.CustomXmlPartCollection of the presentation, a slide or a shape.
export class CustomXmlPartCollectionProxy extends PartsRead {
  protected parts(): XmlPart[] {
    return listOf(this.deck, this.owner);
  }

  add(xml: string): CustomXmlPartProxy {
    return this.item(addPart(this.deck, this.owner, xml).id);
  }
  getByNamespace(namespaceUri: string): ScopedCustomXmlPartCollectionProxy {
    if (typeof namespaceUri !== "string") {
      throw invalidArgument("the namespace must be a string");
    }
    return new ScopedCustomXmlPartCollectionProxy(
      this.deck,
      this.owner,
      namespaceUri,
    );
  }
}

export function presentationXmlParts(
  deck: FakePresentation,
): CustomXmlPartCollectionProxy {
  return new CustomXmlPartCollectionProxy(deck, { kind: "presentation" });
}

export function slideXmlParts(
  deck: FakePresentation,
  slideId: string,
): CustomXmlPartCollectionProxy {
  return new CustomXmlPartCollectionProxy(deck, { kind: "slide", id: slideId });
}

export function shapeXmlParts(
  deck: FakePresentation,
  shapeId: string,
): CustomXmlPartCollectionProxy {
  return new CustomXmlPartCollectionProxy(deck, { kind: "shape", id: shapeId });
}

function snapshotOf(lists: Lists, owner: XmlOwner): CustomXmlSnapshot[] {
  return (lists.get(keyOf(owner)) ?? []).map((part) => ({
    id: part.id,
    namespaceUri: part.namespace,
    xml: part.xml,
  }));
}

// What the deck holds for an owner, as a test sees it: the presentation's by
// default, every write made so far in. Reads the store itself, so a shape that
// is gone shows what it left.
export function customXmlSnapshot(
  deck: FakePresentation,
  owner: XmlOwner = { kind: "presentation" },
): CustomXmlSnapshot[] {
  return snapshotOf(storeOf(deck).lists, owner);
}

// What the host holds: the deck as the last sync that landed left it, without
// the writes of a batch that has not synced yet.
export function customXmlCommitted(
  deck: FakePresentation,
  owner: XmlOwner = { kind: "presentation" },
): CustomXmlSnapshot[] {
  return snapshotOf(storeOf(deck).committed, owner);
}
