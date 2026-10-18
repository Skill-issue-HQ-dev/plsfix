// The sentences both panes' Starter styles receipts share: what the press added or restyled, and
// the styles the user holds under a starter's name. Pure, no Office.js; each flow adds its own
// refusals after these. Invariant: every part after the first opens with a space, so a receipt
// is the parts concatenated.

/** "Callout", "Callout and Highlight", "Brand fill, Callout and Highlight". */
export function nameList(names: readonly string[]): string {
  if (names.length <= 1) return names.join("");
  return `${names.slice(0, -1).join(", ")} and ${names.slice(-1).join("")}`;
}

/** What a Starter styles press did to the starters themselves: the receipt's first sentence. */
export function startersDone(added: number, restyled: number): string {
  if (added === 0 && restyled === 0) return "No starter styles were added.";
  if (restyled === 0) return `Added ${String(added)} starter styles.`;
  if (added === 0) {
    return `Restyled ${String(restyled)} starter styles from the palette.`;
  }
  return `Added ${String(added)} and restyled ${String(restyled)} starter styles from the palette.`;
}

/** Styles the user holds under a starter's name (saved by hand, or a starter they updated). */
export function yoursNow(names: readonly string[]): string {
  const one = names.length === 1;
  return ` ${nameList(names)} ${one ? "is" : "are"} yours now, so the starters leave ${one ? "it" : "them"} alone.`;
}
