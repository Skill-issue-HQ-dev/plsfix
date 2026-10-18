// The formatting and paste groups of the checklist: the Model formatting styles, every format
// cycle and hygiene step on the Format lab and the Data grid, the paint slots, and the whole
// paste suite on the Paste lab.

use crate::layout::{FORMAT_SHEET, PASTE_SHEET};
use crate::sheets::start::{group, step, Item};
use crate::sheets::{data, format_lab as f, paste_lab as p};

pub fn items() -> Vec<Item> {
    let mut items = vec![group("Formatting: Format lab and Data")];
    items.extend(style_steps());
    items.push(group("Copy and paste: Paste lab"));
    items.extend(paste_steps());
    items
}

fn style_steps() -> Vec<Item> {
    let (slot1, slot2) = f::sample_addresses();
    vec![
        step("Cell styles: Title, Header, Input, Formula, Result, Clear", FORMAT_SHEET, format!("Click {} and press Title; select {} and press Header; {} Input; {} Formula; {} Result; then select any styled cell and press Clear", f::title_address(), f::header_address(), f::typed_address(), f::share_address(), f::total_address()), "Each look comes from the Brand palette: ink title, ruled header, blue inputs, black formulas, a result row; Clear strips the formatting and keeps the values"),
        step("Number formats", FORMAT_SHEET, format!("Select {} and press 1,234, then 1,234.0, then € 1,234; select {} and press 12.3%", f::principal_address(), f::rate_address()), "The same numbers in four looks (the € symbol and its side follow the Brand tab); Rate and Share are fractions, so 12.3% reads right"),
        step("Format cycles: number, date, currency, percent, multiple", FORMAT_SHEET, format!("Press Date three times on {}; % on {}; 0.0x on {}; Cur. and 1,234 on {}", f::maturity_address(), f::rate_address(), f::cover_address(), f::principal_address()), "Each press steps to the next variant: 31.12.2026, Dec-26, 2026; 12.3%, 12%, 12.34%; 1.5x, 1.50x; the receipt names the step"),
        step("Row styles: Title, Result, Item", FORMAT_SHEET, format!("Select {} and press Title (row); {} and press Result (row); one item row and press Item (row)", f::header_address(), f::total_address()), "Title steps Filled, Header rule, Accent rule; Result steps Filled, Double rule, Accent block; Item steps Plain, Header tint, Result tint"),
        step("Fill, Font, Border", FORMAT_SHEET, format!("Select {} and press Fill, Font and Border, each more than once", f::table_address()), "Fill and font walk through your Brand palette; borders step bottom rule, total, result, box, grid"),
        step("Row height, Column width", FORMAT_SHEET, format!("Select {} and press Row height; select {} and press Column width", f::header_address(), f::maturity_address()), "Rows step 15, 18, 21, 24, 30 pt; columns 64, 80, 96, 120, 48 pt (outside Undo, the receipt says so)"),
        step("Pinstripes", data::NAME, format!("Select {} and press Pinstripes; press it again; then press Pinstripes ↔", data::grid_address()), "Every second row banded in a light tint of your palette; the second press takes the bands off; the arrows button bands the columns"),
        step("Indent, Align, Underline", FORMAT_SHEET, format!("Select {} and press Indent three times and Align; select {} and press Underline", f::items_address(), f::header_address()), "Indent 1, 2, 3, then 0; alignment left, centre, right, then general; underline single, double, then off"),
        step("Custom styles: Save style, Starter styles from palette", FORMAT_SHEET, format!("Click {slot1}, type Margin in the Custom styles box and press Enter; select {} and click Margin in the list; then press Starter styles from palette and put pls,fix Header on {}", f::principal_address(), f::header_address()), "Margin is now a real style in Excel's Home > Cell Styles gallery and the cells you picked wear it; the eight pls,fix styles take your Brand colors, and pressing the button again after a palette change restyles every cell wearing them; Undo takes an applied style back"),
        step("Format painter: Save 1-3, Use 1-3", FORMAT_SHEET, format!("Click {slot1} and press Save 1; click {slot2} and press Save 2; select {} and press Use 1; select {} and press Use 2", f::principal_address(), f::cover_address()), "Slot 1's bold olive text on a tint lands on the principal column, slot 2's box and 0.0x on the cover column; slots are saved with the workbook"),
    ]
}

fn paste_steps() -> Vec<Item> {
    let (exact, duplicate) = (p::vat_formulas(2), p::vat_formulas(3));
    vec![
        step("Mark, Values, Formats, Transpose", PASTE_SHEET, format!("Select {} and press Mark; click {} and press Values; {} Formats; {} Transpose", p::source_address(), p::target_anchor(0), p::target_anchor(1), p::target_anchor(4)), "The values pad holds the numbers with no formulas; the formats pad the look with no numbers; the transposed pad has rows as columns"),
        step("Exact and Paste: duplicate formulas", PASTE_SHEET, format!("With the block still marked, click {} and press Exact; click {} and press Paste: duplicate formulas", p::target_anchor(2), p::target_anchor(3)), format!("Exact keeps the first VAT formula as {} (it still reads the source row); duplicate rewrites it to {} (the pad's own row) and keeps {} fixed", exact.0, duplicate.1, p::vat_address())),
        step("Paste: number formats", PASTE_SHEET, format!("With the block marked, select {} (the plain block) and press Paste: number formats", p::raw_address()), "The plain numbers take the source's formats (0, 0.00, #,##0.00, €); values, formulas and fills stay as they are"),
        step("Paste: row heights", PASTE_SHEET, format!("With the block marked, click {} and press Paste: row heights", p::raw_corner()), "The plain block's five rows take the source's heights, 30 pt for the header, 21 and 24 pt below (outside Undo, the receipt says so)"),
    ]
}
