// The last checklist groups: the Links tab in Excel (a table, a chart, a text, a picture, projects,
// the highlight, both transports) and the PowerPoint pane on the demo deck, slides 2 to 8. The rows
// run in an order a reader can follow: a link already on a slide leaves the Inbox, so every slide
// takes an export of its own, and Clear pasted links comes last because Revert needs the older copy.

use crate::layout::{
    a1_row, assumptions as a, cell, last_year_col, pnl, range, ASSUMPTIONS_SHEET, FIRST_YEAR_COL, HEADER_ROW,
    HEADLINES_SHEET, LABEL_COL, PNL_SHEET,
};
use crate::sheets::start::{group, step, Item};
use crate::sheets::{bridge, data, headlines, rounding};

const INBOX: &str = "PowerPoint: pls,fix tab, Inbox";
const LINKS: &str = "PowerPoint: pls,fix tab, Links";
const TOOLS: &str = "PowerPoint: pls,fix tab, Tools";
const TRANSPORT: &str = "Links tab, Transport";

pub fn items() -> Vec<Item> {
    let mut items = vec![group("Links tab in Excel: table, chart, text, picture")];
    items.extend(export_steps());
    items.extend(link_steps());
    items.push(group("PowerPoint on the demo deck: slides 2 to 8"));
    items.extend(insert_steps());
    items.extend(update_steps());
    items.extend(object_steps());
    items
}

/// The P&L block every table export reads (A10:H24).
fn pnl_block() -> String {
    range(HEADER_ROW, LABEL_COL, pnl::NET_INCOME, last_year_col())
}

fn export_steps() -> Vec<Item> {
    let (title, margin) = (headlines::text_cell(0), headlines::text_cell(1));
    vec![
        step("Export selection (table), Export active chart", "Links tab", format!("Select {} on {PNL_SHEET} and press Export selection; click the Revenue chart on {} and press Export active chart; click the Segment pie on {} and press Export active chart", pnl_block(), bridge::NAME, rounding::NAME), "Three links in the list, each copied for PowerPoint (the pane says so; if the clipboard refused, press Copy for PowerPoint). A copy holds only its own export: in PowerPoint's Inbox tab click the Paste from Excel box and press Ctrl+V (Cmd+V on a Mac) after each export, or press Copy all once every export is made and paste that one copy"),
        step("Export as text", HEADLINES_SHEET, format!("Select {title} and press Export as text; export {margin} the same way"), "Two links of kind Text in the list, copied for PowerPoint like the first three; once pasted, each sits in the Inbox as one sentence that follows the P&L"),
        step("Export as picture", data::NAME, format!("Select {} and press Export as picture", data::grid_address()), "A picture link in the list, copied for PowerPoint; Slide 4 takes it below, and Convert to table turns it into a table there"),
    ]
}

fn link_steps() -> Vec<Item> {
    let (title, margin, spare) = (headlines::text_cell(0), headlines::text_cell(1), headlines::text_cell(2));
    vec![
        step("New project, Move to project, Go to source, Remove link", "Links tab", format!("Press New project, name it Demo and press Create; tick two links and press Move to project; tick one and press Go to source; export {spare} on {HEADLINES_SHEET} as a spare, tick that new link and press Remove link (press twice); then pick All projects in the Project box"), "The list groups by project; Go to source selects the exported range; the spare link leaves the list (a removed link leaves its picture on the slide); Copy all and Push all follow the picked project, so All projects puts every export back in one copy"),
        step("Highlight linked cells", "Links tab", format!("Tick Highlight linked cells (the Audit overlay has to be off), look at {} on {PNL_SHEET}, {title} and {margin} on {HEADLINES_SHEET} and {} on {}, then untick it", pnl_block(), data::grid_address(), data::NAME), "Each exported range takes a light tint (a chart has no cells to tint); unticking puts the fills back exactly as they were"),
        step("Copy all, Copy selected: the local transport", TRANSPORT, "Keep Links travel on Copy and paste on this computer; tick two links and press Copy selected, then press Copy all (with All projects picked); paste the copy into PowerPoint's Inbox box", "Nothing leaves this computer: the pane copies a bundle (the ticked links, or all of them) and PowerPoint's Inbox box takes it"),
        step("Push selected, Push all: the relay", TRANSPORT, format!("Optional: switch Links travel to Through the relay, press Generate, then Copy and Reveal the key; in PowerPoint's Settings tab choose Through the relay too, paste the key and press Save key; back in Excel press Push selected (two links ticked) or Push all, or tick Auto-push on edit and change a number in the {} grid; press Forget (press twice) to unpair, and switch Links travel back to Copy and paste on this computer in both panes", data::NAME), "The relay needs no pasting once the key is saved: PowerPoint's Update all follows each push, and Auto-push on edit sends a changed link 3 s after the last edit while the pane stays open"),
        step("Source missing", PNL_SHEET, format!("Delete rows {}-{} (the whole exported block) and press Copy all (Push all through the relay); then undo with Cmd+Z and copy again", a1_row(HEADER_ROW), a1_row(pnl::NET_INCOME)), "The pane reports 1 missing and the link reads \"source missing\"; after the undo it copies again (deleting rows inside the block only shrinks it)"),
    ]
}

fn insert_steps() -> Vec<Item> {
    let (title, margin) = (headlines::text_cell(0), headlines::text_cell(1));
    vec![
        step("Insert and update: table, chart, picture", INBOX, format!("Press Copy all in Excel and paste the copy into the Inbox box, so every export waits there. Pick Slide 2, Where = Whole slide and Insert the P&L table; Slide 3, Left half for the Revenue chart and Right half for the Segment pie; Slide 4, select the empty placeholder, Where = Selected shape for the Data picture. Then set {ASSUMPTIONS_SHEET}!{} to 12% in Excel, press Copy all (Push all through the relay), paste into the Inbox box again and press Update all", cell(a::GROWTH, FIRST_YEAR_COL)), "Each kind lands on its own slide in the spot named and leaves the Inbox; the table and the charts refresh in place with the new numbers, sizes kept"),
        step("Convert to table", "Links tab, then PowerPoint: Inbox", format!("Tick the {}!{} link under Linked objects on the Links tab and press Convert to table, then paste the copy into PowerPoint's Inbox box (with the relay, press Update all)", data::NAME, data::grid_address()), "The link becomes a table link and the picture on Slide 4 is drawn again as a native table in the same corner, at the same width"),
        step("Text, table and chart in four quarters", INBOX, format!("In Excel export {} on {PNL_SHEET} and the Revenue chart on {} once more (every export is a new link, and the first two already sit on Slides 2 and 3), then paste them into the Inbox box. Pick Slide 5: Insert {title} with Where = Top left, {margin} with Top right, the new P&L table with Bottom left and the new Revenue chart with Bottom right", pnl_block(), bridge::NAME), "Four objects of three kinds, each fitted into its quarter with its shape kept; the text boxes keep their place when they refresh"),
        step("Paste latest linked", INBOX, format!("Add a slide at the end of the deck (Home, New Slide) and set Slide to This slide and Where to Free space; in Excel click the Segment pie on {} and press Export active chart once more, paste it into the Inbox box and click Paste latest linked", rounding::NAME), "The newest export, the pie you just made, lands on the new slide in one click, with no need to open the list; older exports still waiting are left alone"),
    ]
}

fn update_steps() -> Vec<Item> {
    vec![
        step("Update selected, Update this slide, Revert last update", LINKS, "Tick the P&L table on Slide 2 and press Revert last update, then press Update selected; go to Slide 3, tick its Revenue chart, press Revert last update and then Update this slide", "Revert puts the ticked link back one version, the numbers from before the 12% change (one step only), and its row reads Update available; Update selected repaints only the ticked row, Update this slide every link on the slide"),
        step("Go to slide, Change source, Redraw from this deck, Break link", LINKS, format!("Tick a row and press Go to slide; for Change source export {} on {PNL_SHEET} once more and paste it into the Inbox box, tick the P&L table on Slide 2, press Change source, pick that export and press Confirm; tick a chart and press Redraw from this deck; tick a row and press Break link (press twice)", pnl_block()), "The slide opens; the object keeps its slide, place and size but shows the new source; the chart is drawn again from the deck's own copy; a broken link stays on the slide and stops updating"),
        step("Not on a slide: Put back, Remove", LINKS, "Delete a linked object from its slide, wait about half a minute for its row to read Not on a slide, then press Put back (it is drawn again at the slide and spot picked on the Inbox tab) or Remove (press twice)", "Put back restores the object with its link from the copy the deck keeps, with no Excel; Remove deletes that copy and the Inbox shows the export again if Excel sends it"),
        step("Links list filters", LINKS, "Type part of a slide, object or workbook name in the search box, then try the Source workbook, Slide and Link status dropdowns", "The link table narrows to matching rows; a ticked row stays ticked even if the filter hides it"),
        step("Clear pasted links", INBOX, "Press Clear pasted links (press twice) when you want to start the paste over", "The pasted copies kept on this computer go, so Revert last update has nothing older to go back to; the objects already on the slides stay and keep their links"),
        step("Settings: Save key, Forget key", "PowerPoint: pls,fix tab, Settings", "With the relay chosen, paste the key from Excel and press Save key; Forget key (press twice) deletes it from this computer", "The Inbox opens to what Excel pushed; after Forget key the deck's links stay and only the Inbox closes"),
        step("Brand colors in the deck: palette and theme", "PowerPoint: pls,fix tab, Settings", "Under Brand colors change a color, press Add color, or paste the JSON that Copy palette JSON gave in Excel into Paste colors; press Use this deck's theme colors; press Put in the deck theme, open Home, Shape Fill and look under Theme Colors, then press Restore theme colors; after a paste from Excel whose colors differ from the deck's, press Use the workbook's colors", "The palette is saved in the deck and in PowerPoint as the default for decks with none of their own, and every copy pasted from Excel brings its brand colors along; Put in the deck theme fills Accent 1 to 6 so PowerPoint's own color picker offers your brand and Restore theme colors puts the old ones back (both need Microsoft 365 2601 or Mac 16.105); Use the workbook's colors adopts a pasted palette that differs from the deck's"),
    ]
}

fn object_steps() -> Vec<Item> {
    vec![
        step("Align and Distribute", TOOLS, "On Slide 6 select all five boxes, choose Top in the Align picker and press Apply; select A, C and E, choose Across in the Distribute picker and press Apply (the ribbon's Arrange group does the same)", "The five boxes share one top edge; A, C and E end up with equal gaps between them and the outer two stay put"),
        step("Match size, Select similar, Swap", TOOLS, "On Slide 7 click box A and press Select similar; select D first, then A and E, and press Match size; select the two circles and press Swap", "B and C join A (same kind and size; the larger D and the small E do not); A and E take D's width and height; the circles exchange places"),
        step("Smart Painter", TOOLS, "On Slide 8 select the olive box and press Capture; select the three grey boxes and press Apply", "The grey boxes take the olive fill, the ink outline and the text look (font, size, color, bold, italic); their size, place and words stay as they were"),
        step("Shape styles: save, apply, update, delete, starters", TOOLS, "On Slide 8 select the olive box, type Olive box in the Style name field under Shape styles and press Save style (Enter in the field does the same); select the three grey boxes and click Olive box in the list; select the first grey box and press the refresh icon on the Olive box row, then press the cross on that row twice; with the three grey boxes selected press Starter styles from palette under Shape styles and click Callout", "The pane says Saved style Olive box and the row shows a preview of its fill, outline and text; the grey boxes take the olive fill, the ink outline and the text look and the pane says Style Olive box on 3 objects; the refresh icon rewrites the style from the box you selected and the cross deletes it on the second press; the six starters (Brand fill, Accent fill, Soft fill, Brand outline, Callout, Highlight) come in your brand colors, pressing the button again restyles them after a palette change, and the list belongs to this computer, so every deck shows it"),
        step("Brand colors on shapes: swatches and Edit colors", TOOLS, "On Slide 8 select the three grey boxes and click a Fill swatch under Brand colors, then an Outline swatch and a Text swatch; point at a swatch and click one of its five shades; click None under Fill and Outline; press Edit colors to open the Settings tab at the palette", "The boxes take the brand color as their fill, outline or text color and the pane says which, for example Fill #6D7E12 on 3 objects; a shade is a lighter or darker step of the swatch you pointed at; None clears the fill or hides the outline"),
    ]
}
