// The PowerPoint pane's help copy, in the same shape as the Excel half. Owns
// the wording of the deck side only. Its Linked objects section is keyed
// "deck-links" through a data-help attribute, because the Excel pane already
// owns "links-heading" and one key may mean one section.

import type { HelpCopy } from "./copy";

export const PPT_HELP: HelpCopy = {
  "deck-links": {
    about:
      "Every pls,fix object in this deck by project. A deleted one stays listed as Not on a slide. Buttons refresh, redraw or release them.",
    note: "A deleted object appears as Not on a slide about half a minute after the deck opens, never from the first scan.",
    buttons: {
      "refresh-links":
        "Scans the deck again and checks which links have a newer picture waiting.",
      "update-selected":
        "Repaints the ticked links with their latest data. Slide and position do not move. A changed picture chart becomes shapes.",
      "update-slide":
        "Repaints every link on the slide you are on. A changed picture chart becomes shapes.",
      "update-all":
        "Repaints every link the filters show (All projects: the whole deck); a changed picture chart becomes shapes. In this deck: nothing newer.",
      "revert-selected":
        "Puts the ticked links back to the picture before the last update. One step only.",
      "go-to-slide": "Jumps to the slide the first ticked link sits on.",
      "change-source":
        "Points one ticked link at another export waiting in the Inbox. Tick exactly one row.",
      "change-source-confirm":
        "Re-points the link at the chosen export. Slide, position and size stay as they are.",
      "change-source-cancel":
        "Closes the picker and leaves the link pointing where it did.",
      "deck-redraw":
        "Draws the ticked links again from the copy this deck keeps, with no Excel. Repairs a broken chart or turns a picture chart into shapes.",
      "deck-put-back":
        "Draws a Not on a slide object again at the slide and spot picked on the Inbox tab, with its link. Excel's next push refreshes it.",
      "deck-remove":
        "Deletes this deck's copy of a Not on a slide link and its link key. Its Inbox row shows again if Excel pushes it. Press twice.",
      "break-selected":
        "Takes the link marker off the ticked objects. They stay on the slide but stop updating. Press twice.",
    },
  },
  "inbox-heading": {
    about:
      "Exports from Excel not yet in this deck; they expire after 7 days. Pick a slide and spot below. A chart lands as shapes, else a picture.",
    buttons: {
      "refresh-inbox":
        "Reads what is waiting to insert: pushed to the relay, or pasted here in local mode.",
      "clear-pasted-links":
        "Deletes every pasted link and waiting item kept on this computer. The deck keeps its objects. Press twice.",
      "paste-latest-linked":
        "Inserts the newest Excel export at the picked slide and spot, and creates its tracked link.",
      "first-run-ppt-dismiss":
        "Dismisses this card. It will not show again on this machine.",
    },
  },
  "pairing-heading": {
    about:
      "How links arrive from Excel on this computer: pasted here with no relay, or through the relay once a key is paired.",
    buttons: {
      "save-key":
        "In relay mode, stores the pasted key on this computer and opens the inbox items pushed from Excel.",
      "forget-key":
        "Deletes the key from this computer. The deck's links stay, but the inbox closes. Press twice.",
    },
  },
  "object-tools-heading": {
    about:
      "Moves and sizes the objects selected on the active slide without changing their content.",
    buttons: {
      "align-objects":
        "Aligns selected objects using the direction chosen beside it.",
      "distribute-objects":
        "Spaces three or more selected objects evenly across or down.",
      "match-size":
        "Makes every target the width and height of the first selected object.",
      "select-similar":
        "Selects every object on this slide with the same type and size.",
      "swap-objects": "Swaps the positions of exactly two selected objects.",
    },
  },
  "smart-painter-heading": {
    about:
      "Copies a solid fill, an outline and the text look without changing object content or geometry.",
    buttons: {
      "capture-object-style":
        "Captures the fill, outline and text look of one selected object.",
      "apply-object-style":
        "Applies the captured fill, outline and text look to every selected target.",
    },
  },
  // The brand colors, Tools tab: the one-click swatches. Their runtime buttons
  // carry no id, so the card explains them in the section's own lines.
  "ppt-swatches": {
    about:
      "One click paints the selected objects in a brand color: fill, outline or text. One that cannot take it is skipped and counted.",
    note: "Point at a swatch and its five shades appear below, one click each. Text colors only the selected text when some is selected.",
    buttons: {
      "edit-brand-colors":
        "Opens the Settings tab at the palette, where you add, change and paste your brand colors.",
    },
  },
  // The brand colors, Settings tab: the palette of this deck and its theme.
  "ppt-brand": {
    about:
      "The brand colors of this deck, saved in it and in PowerPoint as the default for decks with none. The Tools swatches use them.",
    note: "Excel's palette comes with a link copy, or paste hex codes, rgb() values or Excel's Copy palette JSON. Brand 1 and 2 stay.",
    buttons: {
      "add-brand-color":
        "Adds one more brand color to the end of the list. Ten is the most.",
      "use-workbook-colors":
        "Takes the brand colors from the last copy pasted from Excel. Enabled when they differ from this deck's.",
      "use-deck-theme-colors":
        "Reads Accent 1 to 6 of this deck's theme into the brand colors, each color once.",
      "put-theme-colors":
        "Writes the brand colors into Accent 1 to 6 of the deck theme, so PowerPoint's own color picker offers them.",
      "restore-theme-colors":
        "Puts back the theme colors from before Put in the deck theme. Enabled once a put saved them, or while the deck's saved colors are unread.",
    },
  },
};
