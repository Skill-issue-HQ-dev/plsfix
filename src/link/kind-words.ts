// The words both panes print for a link's kind: the Excel list's second line and
// the deck list's and Inbox's meta line, one table so they cannot drift.
// Invariant: a picture the user chose (a range link, or a chart kept as a picture
// at the long-draw question) reads "Picture (your choice)", never "range" or "Chart".

import { PICTURE_CHOSEN, type LinkKind, type LinkTag } from "./model";

export const KIND_WORD: Record<LinkKind, string> = {
  range: PICTURE_CHOSEN,
  chart: "Chart",
  table: "Table",
  text: "Text",
};

// The word for one link: its kind's, unless the deck's tag says the user chose
// the picture it holds (only the deck's tag knows; the registry has no choice).
export function kindWord(kind: LinkKind, choice?: LinkTag["choice"]): string {
  return choice === "picture" ? PICTURE_CHOSEN : KIND_WORD[kind];
}
