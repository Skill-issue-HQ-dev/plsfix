// Shows each element marked data-shown-in="local" or "relay" in that link mode only (how
// links reach PowerPoint on this device), for both panes, and hides each one marked
// data-offers-relay in a build with no relay (src/link/local-only.ts). No Office.js.
// Invariant: an element with no marker is never touched.

import { localOnly } from "../link/local-only";
import type { LinkTransport } from "../link/transport-setting";

export function showForMode(root: ParentNode, mode: LinkTransport): void {
  for (const element of root.querySelectorAll<HTMLElement>("[data-shown-in]")) {
    element.hidden = element.dataset.shownIn !== mode;
  }
  if (!localOnly()) return;
  for (const element of root.querySelectorAll<HTMLElement>(
    "[data-offers-relay]",
  )) {
    element.hidden = true;
  }
}
