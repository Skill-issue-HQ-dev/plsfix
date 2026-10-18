// PowerPoint.DocumentProperties (PowerPointApi 1.7) in the fake host: the one
// member the add-in writes is the deck's title (src/ppt/save-nudge.ts). A write
// is an edit as far as the web's editor is concerned (editor-state.ts), however
// little it changes. Invariant: the title lives in editor-state.ts, so a proxy
// from an earlier batch reads and writes the same one.

import { documentTitle, setDocumentTitle, touch } from "./editor-state";
import type { FakePresentation } from "./model";
import { Loadable } from "./strict";

export class DocumentPropertiesProxy extends Loadable {
  constructor(private deck: FakePresentation) {
    super();
  }

  get title(): string {
    return documentTitle(this.deck);
  }
  set title(value: string) {
    if (typeof value !== "string") {
      throw new Error("InvalidArgument: title must be a string");
    }
    setDocumentTitle(this.deck, value);
    touch(this.deck);
  }
}
