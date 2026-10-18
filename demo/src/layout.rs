// Where things sit in the model sheets: the year columns, the Assumptions and
// P&L row maps and the A1 helpers. One source for formulas, charts, notes and
// the checklist, so a moved row changes in exactly one place.

pub const YEARS: [&str; 6] = ["2024A", "2025E", "2026E", "2027E", "2028E", "2029E"];
pub const LABEL_COL: u16 = 0;
pub const UNIT_COL: u16 = 1;
pub const FIRST_YEAR_COL: u16 = 2;

/// Every visible sheet opens with a 7-row guide band (title, up to four
/// tasks, a Commands line, a blank row); every content row constant below is
/// `GUIDE_ROWS + n` so the band and the content it precedes can never
/// collide. The hidden Scratch sheet is the one exception: it carries no
/// band and does not shift.
pub const GUIDE_ROWS: u32 = 7;

pub const TITLE_ROW: u32 = GUIDE_ROWS;
pub const HEADER_ROW: u32 = GUIDE_ROWS + 2;
pub const PNL_SHEET: &str = "P&L";
pub const ASSUMPTIONS_SHEET: &str = "Assumptions";
pub const VARIANCE_SHEET: &str = "Variance";

/// Assumptions rows; the forecast years 2025E-2029E start in FIRST_YEAR_COL.
pub mod assumptions {
    use super::GUIDE_ROWS;

    pub const GROWTH: u32 = GUIDE_ROWS + 3;
    pub const GROSS_MARGIN: u32 = GUIDE_ROWS + 4;
    pub const OPEX_RATIO: u32 = GUIDE_ROWS + 5;
    pub const DA_RATIO: u32 = GUIDE_ROWS + 6;
    pub const INTEREST: u32 = GUIDE_ROWS + 7;
    pub const TAX_RATE: u32 = GUIDE_ROWS + 8;
    pub const REVENUE_2024: u32 = GUIDE_ROWS + 10;
}

/// P&L rows; 2024A sits in FIRST_YEAR_COL, the forecast years follow.
pub mod pnl {
    use super::GUIDE_ROWS;

    pub const REVENUE: u32 = GUIDE_ROWS + 3;
    pub const GROWTH: u32 = GUIDE_ROWS + 4;
    pub const COGS: u32 = GUIDE_ROWS + 5;
    pub const GROSS_PROFIT: u32 = GUIDE_ROWS + 6;
    pub const GROSS_MARGIN: u32 = GUIDE_ROWS + 7;
    pub const OPEX: u32 = GUIDE_ROWS + 8;
    pub const EBITDA: u32 = GUIDE_ROWS + 9;
    pub const EBITDA_MARGIN: u32 = GUIDE_ROWS + 10;
    pub const DA: u32 = GUIDE_ROWS + 11;
    pub const EBIT: u32 = GUIDE_ROWS + 12;
    pub const INTEREST: u32 = GUIDE_ROWS + 13;
    pub const EBT: u32 = GUIDE_ROWS + 14;
    pub const TAX: u32 = GUIDE_ROWS + 15;
    pub const NET_INCOME: u32 = GUIDE_ROWS + 16;
    pub const NET_MARGIN: u32 = GUIDE_ROWS + 17;
    pub const PER_MONTH: u32 = GUIDE_ROWS + 18;
    pub const CHECK: u32 = GUIDE_ROWS + 19;
}

/// Variance rows: the amounts list starts at FIRST_ROW, one row per item,
/// then a blank row before the Target two rows below the last amount.
pub mod variance {
    use super::GUIDE_ROWS;

    pub const FIRST_ROW: u32 = GUIDE_ROWS + 3;
    pub const COUNT: u32 = 14;
    pub const TARGET_ROW: u32 = GUIDE_ROWS + 18;
    pub const AMOUNT_COL: u16 = 1;
}

pub const COMPS_SHEET: &str = "Comps";
pub const TEMPLATES_SHEET: &str = "Templates";
pub const FORMAT_SHEET: &str = "Format lab";
pub const PASTE_SHEET: &str = "Paste lab";
pub const HEADLINES_SHEET: &str = "Headlines";
pub const REVIEW_SHEET: &str = "Review";

/// Comps rows: the peer table Comps stats reads (header, eight peers, then the blank row and the
/// six empty rows it writes into) and, below those, the Method | Low | High block the Football
/// field reads, with three empty columns to its right for the helper block the tool writes.
pub mod comps {
    use super::GUIDE_ROWS;

    pub const HEADER: u32 = GUIDE_ROWS + 2;
    pub const FIRST_PEER: u32 = HEADER + 1;
    pub const PEER_COUNT: u32 = 8;
    pub const LAST_PEER: u32 = FIRST_PEER + PEER_COUNT - 1;
    /// One blank row under the table, then the six statistic rows.
    pub const STATS_FIRST: u32 = LAST_PEER + 2;
    pub const STATS_ROWS: u32 = 6;
    pub const FOOT_TITLE: u32 = STATS_FIRST + STATS_ROWS + 2;
    pub const FOOT_HEADER: u32 = FOOT_TITLE + 1;
    pub const FOOT_FIRST: u32 = FOOT_HEADER + 1;
    pub const FOOT_COUNT: u32 = 6;
    pub const FOOT_COLS: u16 = 3;
    pub const LAST_COL: u16 = 6;
}

/// Templates rows: an index of the eighteen blocks, then one labelled empty pad per block.
pub mod templates {
    use super::GUIDE_ROWS;

    pub const COUNT: u32 = 18;
    pub const INDEX_HEADER: u32 = GUIDE_ROWS + 2;
    pub const INDEX_FIRST: u32 = INDEX_HEADER + 1;
    /// The label row of the first pad.
    pub const FIRST_PAD: u32 = INDEX_FIRST + COUNT + 2;
    /// The column every pad starts in (column A holds the pad number).
    pub const ANCHOR_COL: u16 = 1;
    /// Blank rows under each pad, before the next label row.
    pub const PAD_GAP: u32 = 2;
}

/// Format lab rows: one plain debt table (title, header, four items, total) and two styled
/// samples to the right for the paint slots.
pub mod format_lab {
    use super::GUIDE_ROWS;

    pub const BLOCK_TITLE: u32 = GUIDE_ROWS + 2;
    pub const HEADER: u32 = BLOCK_TITLE + 1;
    pub const FIRST_ITEM: u32 = HEADER + 1;
    pub const ITEM_COUNT: u32 = 4;
    pub const TOTAL: u32 = FIRST_ITEM + ITEM_COUNT;
    pub const SAMPLE_COL: u16 = 7;
}

/// Paste lab rows: the VAT input, the marked source block and five empty targets, then a plain
/// block that only takes formats and row heights.
pub mod paste_lab {
    use super::GUIDE_ROWS;

    pub const VAT_ROW: u32 = GUIDE_ROWS + 2;
    pub const SRC_HEADER: u32 = VAT_ROW + 2;
    pub const SRC_FIRST: u32 = SRC_HEADER + 1;
    pub const SRC_ROWS: u32 = 4;
    /// Header plus items: every target and the raw block share this height.
    pub const BLOCK_ROWS: u32 = SRC_ROWS + 1;
    pub const COLS: u16 = 5;
    pub const TARGETS: u32 = 5;
    pub const FIRST_TARGET_LABEL: u32 = SRC_FIRST + SRC_ROWS + 1;
    /// A label row, the block, a blank row.
    pub const TARGET_STRIDE: u32 = BLOCK_ROWS + 2;
    pub const RAW_LABEL: u32 = FIRST_TARGET_LABEL + TARGET_STRIDE * TARGETS;
}

/// Headlines rows: three sentences built from the P&L, one per row, label left and text right.
pub mod headlines {
    use super::GUIDE_ROWS;

    pub const HEADER: u32 = GUIDE_ROWS + 2;
    pub const FIRST: u32 = HEADER + 1;
    pub const COUNT: u32 = 3;
}

/// Review rows: an error to guard, a formula to fill down, three things a reviewer flags and one
/// formatted blank far down and to the right of the data.
pub mod review {
    use super::GUIDE_ROWS;

    pub const PRICE_LABEL: u32 = GUIDE_ROWS + 2;
    pub const PRICE_HEADER: u32 = PRICE_LABEL + 1;
    pub const PRICE_FIRST: u32 = PRICE_HEADER + 1;
    pub const PRODUCTS: u32 = 5;
    pub const MARGIN_LABEL: u32 = PRICE_FIRST + PRODUCTS + 1;
    pub const MARGIN_HEADER: u32 = MARGIN_LABEL + 1;
    pub const MARGIN_FIRST: u32 = MARGIN_HEADER + 1;
    pub const FLAGS_LABEL: u32 = MARGIN_FIRST + PRODUCTS + 1;
    pub const FLAGS_FIRST: u32 = FLAGS_LABEL + 1;
    pub const FLAGS: u32 = 3;
    /// 0-based row 799 is the sheet's row 800; column 29 is AD.
    pub const STRAY_ROW: u32 = 799;
    pub const STRAY_COL: u16 = 29;
}

/// Column of P&L year `i` (0 = 2024A).
pub const fn year_col(i: usize) -> u16 {
    FIRST_YEAR_COL + i as u16
}

/// Column of the Assumptions entry that drives P&L year `i` (2025E = first).
pub const fn assumption_col(i: usize) -> u16 {
    FIRST_YEAR_COL + i.saturating_sub(1) as u16
}

pub const fn last_year_col() -> u16 {
    FIRST_YEAR_COL + YEARS.len() as u16 - 1
}

/// 0-based column index to Excel letters: 0 -> A, 26 -> AA.
pub fn column(col: u16) -> String {
    let mut n = u32::from(col) + 1;
    let mut letters = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        letters.push(char::from(b'A' + rem as u8));
        n = (n - 1) / 26;
    }
    letters.iter().rev().collect()
}

pub fn a1_row(row: u32) -> String {
    (row + 1).to_string()
}

pub fn cell(row: u32, col: u16) -> String {
    format!("{}{}", column(col), a1_row(row))
}

pub fn cell_abs(row: u32, col: u16) -> String {
    format!("${}${}", column(col), a1_row(row))
}

/// A quoted reference into the P&L sheet, relative or absolute.
pub fn pnl_ref(row: u32, col: u16) -> String {
    format!("'{PNL_SHEET}'!{}", cell(row, col))
}

pub fn pnl_ref_abs(row: u32, col: u16) -> String {
    format!("'{PNL_SHEET}'!{}", cell_abs(row, col))
}

pub fn assumptions_ref_abs(row: u32, col: u16) -> String {
    format!("{ASSUMPTIONS_SHEET}!{}", cell_abs(row, col))
}

/// An A1 range from two corners.
pub fn range(first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> String {
    format!("{}:{}", cell(first_row, first_col), cell(last_row, last_col))
}
