// What the PowerPoint pane says in words that depend on how links reach this
// computer: copy and paste ("local", the default of a fresh install) or the
// relay. Pure text. Invariant: relay mode keeps the words it always had.

import type { LinkTransport } from "../link/transport-setting";

const EMPTY_INBOX: Record<LinkTransport, string> = {
  local: "Nothing waiting. In Excel press Export, then paste here.",
  relay: "Nothing waiting. Push a range or a chart from Excel.",
};

const UPDATE_FROM_EXCEL =
  "To update from Excel, copy there and paste on the Inbox tab.";

export function emptyInboxLine(mode: LinkTransport): string {
  return EMPTY_INBOX[mode];
}

// Update all repaints from what this computer holds, which in copy-and-paste
// mode is what was last pasted: its receipt says how newer numbers get here.
export function updateAllReceipt(receipt: string, mode: LinkTransport): string {
  if (mode !== "local") return receipt;
  const sentence = receipt.endsWith(".") ? receipt : `${receipt}.`;
  return `${sentence} ${UPDATE_FROM_EXCEL}`;
}
