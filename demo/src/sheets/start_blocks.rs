// The checklist groups for the number blocks, the charts and the Workbook tab: Find a
// combination, the eighteen templates, Comps stats, rounding, CAGR, unpivot, the five chart
// builders, the sheet tools, Super Find, Model check and the clean-up buttons.

use crate::layout::{COMPS_SHEET, REVIEW_SHEET, TEMPLATES_SHEET};
use crate::sheets::start::{group, step, Item};
use crate::sheets::{bridge, cagr, charts, comps, data, review, rounding, scratch, sensitivity, templates, variance};

pub fn items() -> Vec<Item> {
    let mut items = vec![group("Blocks and numbers: Variance, Templates, Comps, Rounding, CAGR, Data")];
    items.extend(block_steps());
    items.push(group("Charts: Bridge, Sensitivity, Comps, Charts"));
    items.extend(chart_steps());
    items.push(group("Workbook tab: sheets, search, review and clean-up"));
    items.extend(workbook_steps());
    items
}

fn block_steps() -> Vec<Item> {
    vec![
        step("Find a combination", variance::NAME, format!("Select {} and type the number from {} into the pane's Find a combination, then press Find cells", variance::amounts_address(), variance::target_address()), "The invoice lines that add up to the target become the selection; nothing else in the list matches; the Selection inspector counts them"),
        step("Templates: the eighteen blocks", TEMPLATES_SHEET, format!("Click a template's name in the list at the top (Excel jumps to the first cell of its pad, such as {}), then press its Insert button on the Tools tab; the search box finds a block by name, group or line", templates::anchor(0)), format!("The block lands there with inputs blue and formulas black; every one of the {} templates has an empty pad sized to it; Undo last pls,fix action takes a block back", templates::PADS.len())),
        step("Comps stats", COMPS_SHEET, format!("Select {} (header and eight peers) and press Comps stats", comps::table_address()), format!("Six live rows ({}): Min, 25th percentile, Median, Mean, 75th percentile, Max, written under every number column ({}) in one press; change a multiple and they move", comps::stats_rows_words(), comps::number_columns_words())),
        step("Consistent rounding / =PLSFIX.ROUND", rounding::NAME, format!("Select {} and press Consistent rounding, or type the formulas shown in the last column", rounding::points_address()), "PLSFIX.ROUND formulas beside the selection whose parts add up to the rounded total (the Excel ROUND column sums to 101)"),
        step("CAGR sheet", cagr::NAME, format!("Select {} and press Insert CAGR, then compare with columns I and J", cagr::first_row_range()), format!("Ribbon CAGR, =PLSFIX.CAGR and the hand formula agree on {}; {} are refused", cagr::agree_rows_words(), cagr::refused_rows_words())),
        step("Unpivot selection", data::NAME, format!("Select {} and press Unpivot selection", data::grid_address()), "A new sheet with one Row / Column / Value line per cell of the grid"),
    ]
}

fn chart_steps() -> Vec<Item> {
    vec![
        step("Waterfall from selection", bridge::NAME, format!("Select {} and press Waterfall from selection", bridge::table_address()), "A native bridge chart: opening and closing totals, the steps between"),
        step("CAGR label / Brand-format chart", bridge::NAME, "Click the Revenue chart, press CAGR label, then Brand-format chart", "A CAGR label on the chart; brand formatting applied (the pie on Rounding is there to export too)"),
        step("Tornado from selection", sensitivity::NAME, format!("Select {} and press Tornado from selection", sensitivity::block_address()), format!("A ranked sensitivity chart around the base case in {}", sensitivity::base_address())),
        step("Football field", COMPS_SHEET, format!("Select {} and press Football field", comps::football_address()), "One floating bar per valuation method, first row on top; the Method | Low | Range helper block lands in the three empty columns to its right"),
        step("Charts: every chart kind, for the link export", charts::NAME, "Click one of the twelve charts, each with its data block beside it, and press Export active chart on the Links tab; in PowerPoint insert it from the Inbox", "The chart lands with its data: as shapes where this version can draw the kind, otherwise as a picture whose Links row says why; Update all follows a changed number"),
    ]
}

fn workbook_steps() -> Vec<Item> {
    vec![
        step("Sheet explorer", "Workbook tab", format!("Click a sheet's name to go there; press Unhide all, then Show only this on any sheet and Unhide all again; try Move up, Move down and Move to end, and Bury this on {}; then tick Include very hidden and press Unhide all", scratch::NAME), format!("The hidden {} sheet appears with Unhide all; Show only this hides the rest until Unhide all; moves change the tab strip; a buried sheet leaves Excel's own unhide list, and only Unhide all with Include very hidden ticked brings it back", scratch::NAME)),
        step("Find in workbook", "Workbook tab", "Search for \"terminal\"", format!("Two hits: the export order line on {} and this row; click the {} hit to jump to the line", data::NAME, data::NAME)),
        step("Find options: Match case, Search formulas", "Workbook tab", "Search for \"riga\", then tick Match case and search again, then search for \"Riga\"; untick Search formulas and search for OFFSET", format!("\"riga\" finds Riga on {}, {} and {} in any casing, and none of them once Match case is ticked (\"Riga\" finds all three again); OFFSET is a hit on {REVIEW_SHEET} only while Search formulas is ticked, because that cell shows a number", variance::NAME, data::NAME, COMPS_SHEET)),
        step("Run model check", "Workbook tab", "Press Run model check, click a finding to jump to it, then Copy report", format!("Formula error ({}!{}), hardcodes on {REVIEW_SHEET} and the P&L, volatile functions on {REVIEW_SHEET}, the inconsistent formula, the broken name {}, the hidden {} sheet; Copy report puts the list on the clipboard", REVIEW_SHEET, review::error_address(), crate::sheets::BROKEN_NAME, scratch::NAME)),
        step("Insert contents sheet, Scan broken names, Scan styles, Prepare for sharing", "Workbook tab", "Press each button; Delete broken names asks for a second press (Delete unused styles would too, but this file has none)", format!("A contents sheet; {} reported as broken and deleted by Delete broken names; the style scan finds none unused; every sheet back to A1 and the hidden {} sheet reported", crate::sheets::BROKEN_NAME, scratch::NAME)),
        step("Clean past the data", REVIEW_SHEET, format!("Press Ctrl+End (Excel lands on {}, row {}), then on the Workbook tab press Clean past the data twice", review::stray_address(), review::stray_row_number()), format!("The empty rows and columns past the data are deleted (outside Undo); Ctrl+End now lands on the corner of the data, {}", review::data_corner_address())),
    ]
}
