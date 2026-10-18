// The help for the design loop's own sections: the Excel Tools tab's Brand colors and
// Custom styles, and the PowerPoint Tools tab's Shape styles. Owns those three
// sections; the Brand tab's palette and logo sections stay in copy.excel.ts, because
// a section key has exactly one owner.

import type { HelpCopy } from "./copy";

export const DESIGN_HELP: HelpCopy = {
  "brand-colors": {
    about:
      "One click paints the selected cells with a brand color, as a fill or as a font color. Undo takes it back.",
    note: "Point at a swatch and its five shades appear below, lighter or darker, one click each. No fill clears the fill.",
    buttons: {
      "edit-brand-colors":
        "Opens the Brand tab at the palette, where you add, change and paste the brand colors the fill and font swatches use.",
    },
  },
  // The row buttons are built per style: their sentences are listed here and
  // their labels in GENERATED_BUTTONS (src/help/copy.ts).
  "custom-styles": {
    about:
      "Save the look of a cell as a named style. It lands in Excel's own Cell Styles gallery and travels with the workbook.",
    note: "Each row previews the style's fill and text. Ctrl+Shift+Alt+1 steps the selected cells through your styles, then back to Normal.",
    buttons: {
      "style-save":
        "Saves the active cell's look under the name in the box. Enter in the box does the same.",
      "styles-starter":
        "Adds eight starter styles in your brand colors, or restyles them after the palette changed. One you updated is left alone.",
      "style-apply":
        "Puts the style on every selected cell. Undo takes it back.",
      "style-update":
        "Rewrites the style from the active cell. Every cell wearing it follows. A starter you update becomes yours. Undo does not cover it.",
      "style-delete":
        "Deletes the style. The cells wearing it go back to Normal, so it asks twice and counts them.",
    },
  },
  // The PowerPoint pane's list: its row buttons are built per style too.
  "shape-styles": {
    about:
      "Save the look of a shape as a named style: its fill, outline and text. The list is kept on this computer, so every deck has it.",
    note: "The preview on each row shows the style's fill, outline and text color. Pressing a name puts the style on the selected objects.",
    buttons: {
      "save-shape-style":
        "Saves the selected shape's look under the name in the box. Enter in the box does the same.",
      "shape-style-starters":
        "Adds six starter styles in your brand colors, or restyles them after the palette changed. One you updated is yours and stays as it is.",
      "shape-style-apply":
        "Puts the style on every selected object. One that cannot take it is skipped and counted.",
      "shape-style-update":
        "Rewrites the style from the selected shape: its fill, outline and text look replace the old ones. A starter you update becomes yours.",
      "shape-style-delete":
        "Deletes the style from the list. It asks twice, and objects that wear it keep their look.",
    },
  },
};
