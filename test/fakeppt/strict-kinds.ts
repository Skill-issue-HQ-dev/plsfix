// The strict-load descriptors of the members shape-api.ts and customxml.ts add,
// merged into strict.ts's table by one call there (it is near its line cap):
// which are scalars that need a load, which hand back a child or a fresh root.
// Invariant: a kind named here is one strict.ts would leave unpoliced.

import type { Descriptor } from "./strict";

// getCount and getXml are ClientResults; the three getItem forms hand back a
// part; getItemAt is PowerPointApi 1.10.
const READS = {
  getCount: "clientResult",
  getItem: "customXmlPart",
  getItemAt: "customXmlPart",
  getItemOrNullObject: "customXmlPart",
};

// The table strict.ts wrote, plus the kinds and members of the F0 additions.
export function extendKinds(
  base: Record<string, Descriptor>,
): Record<string, Descriptor> {
  const kinds = { ...base };
  const withParts = (kind: Descriptor | undefined): Descriptor => ({
    ...kind,
    children: { ...kind?.children, customXmlParts: "customXmlParts" },
  });
  const presentation = withParts(kinds.presentation);
  kinds.presentation = {
    ...presentation,
    children: { ...presentation.children, properties: "documentProperties" },
  };
  kinds.documentProperties = { scalars: ["title"] };
  // PowerPointApi 1.10 theme colors on the slide masters, and the 1.5 selected
  // text range: a master's id is a load, a theme color a ClientResult.
  kinds.presentation = {
    ...kinds.presentation,
    children: { ...kinds.presentation?.children, slideMasters: "slideMasters" },
    returns: {
      ...kinds.presentation?.returns,
      getSelectedTextRange: "textSelection",
      getSelectedTextRangeOrNullObject: "textSelection",
    },
  };
  kinds.slideMasters = {
    items: "slideMaster",
    returns: {
      ...READS,
      getItem: "slideMaster",
      getItemAt: "slideMaster",
      getItemOrNullObject: "slideMaster",
    },
  };
  kinds.slideMaster = {
    scalars: ["id", "isNullObject"],
    children: { themeColorScheme: "themeColorScheme" },
  };
  kinds.themeColorScheme = { returns: { getThemeColor: "clientResult" } };
  kinds.textFont = { scalars: ["name", "size", "color", "bold", "italic"] };
  kinds.textSelection = {
    scalars: ["isNullObject", "start", "length", "text"],
    children: { font: "textFont" },
  };
  kinds.slide = withParts(kinds.slide);
  const shape = withParts(kinds.shape);
  kinds.shape = { ...shape, scalars: [...(shape.scalars ?? []), "rotation"] };
  kinds.customXmlParts = {
    items: "customXmlPart",
    returns: {
      ...READS,
      add: "customXmlPart",
      getByNamespace: "scopedXmlParts",
    },
  };
  kinds.scopedXmlParts = {
    items: "customXmlPart",
    returns: {
      ...READS,
      getOnlyItem: "customXmlPart",
      getOnlyItemOrNullObject: "customXmlPart",
    },
  };
  kinds.customXmlPart = {
    scalars: ["id", "namespaceUri", "isNullObject"],
    returns: { getXml: "clientResult" },
  };
  return kinds;
}
