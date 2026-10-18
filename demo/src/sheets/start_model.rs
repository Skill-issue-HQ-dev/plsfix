// The first checklist groups: getting started, and the model tools that act on the Assumptions,
// the P&L and the Review sheet (colour, audit, fill, guard, formats, CAGR, the colour key).
// Addresses come from the layout and from the sheets that own the blocks.

use crate::layout::{
    a1_row, assumptions as a, cell, last_year_col, pnl, range, year_col, ASSUMPTIONS_SHEET, FIRST_YEAR_COL,
    HEADER_ROW, LABEL_COL, PNL_SHEET, REVIEW_SHEET,
};
use crate::sheets::pnl::planted_cells;
use crate::sheets::review;
use crate::sheets::start::{group, step, Item};

pub fn items() -> Vec<Item> {
    let mut items = vec![group("Get started")];
    items.extend(start_steps());
    items.push(group("Model tools: Assumptions, P&L and Review"));
    items.extend(audit_steps());
    items.extend(edit_steps());
    items
}

fn start_steps() -> Vec<Item> {
    vec![
        step("Open the pane", "Excel ribbon", "Click the pls,fix tab, then Model Tools; the pane has the tabs Tools, Workbook, Links and Brand", format!("The pane opens on the right; the search box at its top finds any tool by name (press / to jump to it)")),
        step("Brand tab", "Brand tab", format!("Pick another colour under Palette, add up to ten brand colours with Add color or Paste colors (paste \"Gold #B27E54, Marble #282623, rgb(65, 86, 79)\"), or select cells coloured in your own brand and press Read from selection; after Upload company colors from a logo of yours press Use these colors; then repeat Autocolor selection on {PNL_SHEET}; press Copy palette JSON to copy the palette (Import JSON loads a saved file back), then Reset brand settings (the arrow icon beside Palette, press twice)"), "The model paints in your palette; Primary and Accent are the first two brand colours and the swatches on the Tools tab follow the whole list at once; the palette is on the clipboard as JSON; Reset brand settings brings the pls,fix colours back. Output defaults sit on the same tab"),
        step("Keyboard shortcuts: Apply, Reset all, Printable card", "Brand tab", "Under Keyboard shortcuts type a combination beside an action and press Apply; press Printable card (the keyboard icon beside the heading) to see your keys on one card; press Reset all (press twice) to give every action its pls,fix key back", "The pane says the shortcuts are updated and your key runs that action; the card lists the keys as they stand now, and the keyboard button on the Tools tab (Shortcut card) opens the same card; Reset all brings the defaults back"),
    ]
}

fn audit_steps() -> Vec<Item> {
    let last = last_year_col();
    let planted = planted_cells();
    let cogs = cell(pnl::COGS, year_col(1));
    let opex = cell(pnl::OPEX, year_col(2));
    // Three empty cells below the colour key (J10:K15), one per colour the on-edit paint hands out.
    let spare_row = HEADER_ROW + 8;
    let (typed, formula, link) = (cell(spare_row, last + 2), cell(spare_row, last + 3), cell(spare_row, last + 4));
    let (revenue, growth) = (cell(pnl::REVENUE, FIRST_YEAR_COL), format!("{ASSUMPTIONS_SHEET}!{}", cell(a::GROWTH, FIRST_YEAR_COL)));
    vec![
        step("Autocolor selection", PNL_SHEET, format!("Select {} and press Autocolor selection (Tools tab)", range(HEADER_ROW, LABEL_COL, pnl::CHECK, last)), format!("Inputs blue, formulas black, cross-sheet links green; {} and {} stand out", planted[0], planted[1])),
        step("Autocolor on edit", PNL_SHEET, format!("Tick Autocolor on edit on the Brand tab, then type 100 in {typed}, ={revenue}*2 in {formula} and ={growth} in {link}; untick it when done"), "Each cell is painted as you press Enter: the typed number blue, the formula black, the cross-sheet link green (up to 500 cells at a time)"),
        step("Audit overlay", PNL_SHEET, format!("Select {} and press Audit overlay (press again to clear)", range(pnl::REVENUE, FIRST_YEAR_COL, pnl::NET_MARGIN, last)), format!("Inconsistent formulas striped: {} carries a hardcoded 0.21", planted[0])),
        step("Precedents / Dependents", PNL_SHEET, format!("Click {} (EBITDA 2025E) and press Precedents; on {}!{} press Dependents", cell(pnl::EBITDA, year_col(1)), ASSUMPTIONS_SHEET, cell(a::GROWTH, FIRST_YEAR_COL)), "The cells it reads, then the cells that read it, get selected"),
        step("Select consistent region", PNL_SHEET, format!("Click {cogs} (COGS 2025E) and press Select consistent region; then click {opex} (Opex 2026E) and press it again"), format!("{cogs} grows to the whole forecast row, 5 cells; {opex} stops at 2 cells, because the hardcode in {} breaks the row", planted[0])),
        step("Precedents of selection", PNL_SHEET, format!("Select {} (gross profit, 2025E to 2029E) and press Precedents of selection", range(pnl::GROSS_PROFIT, year_col(1), pnl::GROSS_PROFIT, last)), "A list of what each of the five formulas reads, grouped by cell; the ones on this sheet are selected"),
    ]
}

fn edit_steps() -> Vec<Item> {
    let last = last_year_col();
    vec![
        step("Brand colors: swatches and Edit colors", PNL_SHEET, format!("Select {} (Revenue) and click a Fill swatch under Brand colors on the Tools tab, then a Font swatch; point at a swatch and click one of its five shades; click No fill; press Edit colors to open the Brand tab at the palette", range(pnl::REVENUE, FIRST_YEAR_COL, pnl::REVENUE, last)), "The cells take the brand colour as their fill or their font and the toast says which (Fill #6D7E12); a shade is a lighter or darker step of the swatch you pointed at; No fill clears the fill; Undo last pls,fix action takes any of it back"),
        step("Fill formula right", PNL_SHEET, format!("Select {} and press Fill formula right", range(pnl::PER_MONTH, FIRST_YEAR_COL, pnl::PER_MONTH, last)), "The first cell's formula is filled across the row"),
        step("Fill formula down", REVIEW_SHEET, format!("Click {} (the one Margin formula) and press Fill formula down", review::margin_first_address()), format!("The formula and its percent format fill down the Margin column, {}, sized by the columns beside it", review::margin_rows_words())),
        step("IFERROR guard", REVIEW_SHEET, format!("Select {} (Average price) and press IFERROR guard; press it again to strip the guard", review::price_address()), format!("The #DIV/0! on {} becomes 0 and every formula reads IFERROR(..., 0); the second press takes the guard off", review::error_address())),
        step("Number formats, x1000, /1000, Sign flip, decimals", PNL_SHEET, "Select some numbers and press the buttons on the Tools tab (1,234, 12.3%, ÷ 1,000, × 1,000, ± sign, .00 +/−); Undo last pls,fix action takes the last one back", "Formats cycle; values scale, flip sign or gain a decimal; Undo puts back what the last action overwrote"),
        step("CAGR", PNL_SHEET, format!("Select {} (Revenue) and press CAGR", range(pnl::REVENUE, FIRST_YEAR_COL, pnl::REVENUE, last)), "The 2024A-2029E CAGR written beside the row"),
        step("Insert color key", PNL_SHEET, format!("Click {} (empty, right of the table) and press Insert color key", cell(HEADER_ROW, last + 2)), format!("A six-row legend (Color key, then hardcoded input, formula, cross-sheet link, external file link, partial input) in your palette; rows {} to {}", a1_row(HEADER_ROW), a1_row(HEADER_ROW + 5))),
    ]
}
