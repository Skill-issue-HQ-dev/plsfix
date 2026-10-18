// The Format lab: one deliberately plain debt table (title, header, four typed rows, a formula
// column, a total) for every style button, number format, cycle and hygiene step, plus two styled
// sample cells for the paint slots. Invariant: every table cell is General with no font, fill or
// border, so the first press of any button is a change the user can see.

use rust_xlsxwriter::{ExcelDateTime, Worksheet, XlsxError};

use crate::layout::{cell, cell_abs, format_lab::*, range, FORMAT_SHEET, GUIDE_ROWS, LABEL_COL, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = FORMAT_SHEET;
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const GUIDE: Guide = Guide {
    title: "a plain table to style, format and step",
    tasks: &[
        "Style it: Title on the heading, Header on the column titles, Input on typed numbers, Formula on Share, Result on the total.",
        "Number formats: 1,234 and 12.3% on the amounts, then the Date, Cur., % and 0.0x cycles on the matching columns.",
        "Indent, Align, Underline, fill, font, border and the sizes step on the selection; Save a sample's style, Use it on a column.",
    ],
    commands: &[
        "Cycle general number format (Ctrl+Shift+1)",
        "Cycle date format (Ctrl+Shift+2)",
        "Cycle percent format (Ctrl+Shift+5)",
        "Cycle multiple format (Ctrl+Shift+6)",
    ],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const ITEM_COL: u16 = 0;
const PRINCIPAL_COL: u16 = 1;
const RATE_COL: u16 = 2;
const MATURITY_COL: u16 = 3;
const COVER_COL: u16 = 4;
const SHARE_COL: u16 = 5;
const ITEM_WIDTH: f64 = 24.0;
const COLUMN_WIDTH: f64 = 12.0;
const SAMPLE_WIDTH: f64 = 34.0;

/// One facility: principal in EUR thousands, rate as a fraction, maturity as (year, month, day)
/// and collateral cover in x.
struct Loan {
    item: &'static str,
    principal: f64,
    rate: f64,
    maturity: (u16, u8, u8),
    cover: f64,
}

const fn loan(item: &'static str, principal: f64, rate: f64, maturity: (u16, u8, u8), cover: f64) -> Loan {
    Loan { item, principal, rate, maturity, cover }
}

const DEBT: [Loan; ITEM_COUNT as usize] = [
    loan("Senior loan", 4_500.0, 0.045, (2027, 12, 31), 1.8),
    loan("Revolver", 1_200.0, 0.062, (2026, 6, 30), 1.0),
    loan("Bond", 6_000.0, 0.0725, (2029, 3, 15), 2.4),
    loan("Leasing", 800.0, 0.051, (2028, 9, 30), 1.2),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "Format lab: a plain table to style", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &note(), &styles.note)?;
    write_table(&mut pen, styles)?;
    write_samples(&mut pen, styles)?;
    let sheet = pen.sheet();
    sheet.set_column_width(ITEM_COL, ITEM_WIDTH)?;
    sheet.set_column_range_width(PRINCIPAL_COL, SHARE_COL, COLUMN_WIDTH)?;
    sheet.set_column_width(SAMPLE_COL, SAMPLE_WIDTH)?;
    tally += pen.tally;
    Ok(tally)
}

fn note() -> String {
    format!(
        "Nothing in {} is styled on purpose: select a part, press a button, press it again to step. Maturity holds date numbers.",
        table_address()
    )
}

fn write_table(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    let plain = &styles.plain;
    pen.text(BLOCK_TITLE, ITEM_COL, "Debt summary (EUR thousands)", plain)?;
    let titles = ["Item", "Principal", "Rate", "Maturity", "Cover", "Share"];
    for (i, title) in titles.iter().enumerate() {
        pen.text(HEADER, i as u16, title, plain)?;
    }
    let total = cell_abs(TOTAL, PRINCIPAL_COL);
    for (i, loan) in DEBT.iter().enumerate() {
        let row = FIRST_ITEM + i as u32;
        let (year, month, day) = loan.maturity;
        pen.text(row, ITEM_COL, loan.item, plain)?;
        pen.number(row, PRINCIPAL_COL, loan.principal, plain)?;
        pen.number(row, RATE_COL, loan.rate, plain)?;
        pen.number(row, MATURITY_COL, ExcelDateTime::from_ymd(year, month, day)?.to_excel(), plain)?;
        pen.number(row, COVER_COL, loan.cover, plain)?;
        pen.formula(row, SHARE_COL, &format!("={}/{total}", cell(row, PRINCIPAL_COL)), plain)?;
    }
    let (first, last) = (FIRST_ITEM, FIRST_ITEM + ITEM_COUNT - 1);
    pen.text(TOTAL, ITEM_COL, "Total debt", plain)?;
    for col in [PRINCIPAL_COL, SHARE_COL] {
        pen.formula(TOTAL, col, &format!("=SUM({})", range(first, col, last, col)), plain)?;
    }
    Ok(())
}

fn write_samples(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(HEADER, SAMPLE_COL, "Samples for the paint slots", &styles.label_bold)?;
    pen.text(FIRST_ITEM, SAMPLE_COL, "Slot 1: bold olive on a tint", &styles.sample_text)?;
    pen.number(FIRST_ITEM + 1, SAMPLE_COL, 2.5, &styles.sample_number)?;
    Ok(())
}

/// A1 range of the whole table, title and total included.
pub fn table_address() -> String {
    range(BLOCK_TITLE, ITEM_COL, TOTAL, SHARE_COL)
}

/// A1 range of the header row.
pub fn header_address() -> String {
    range(HEADER, ITEM_COL, HEADER, SHARE_COL)
}

/// A1 address of the title cell.
pub fn title_address() -> String {
    cell(BLOCK_TITLE, ITEM_COL)
}

/// A1 range of the typed numbers (principal, rate, maturity and cover).
pub fn typed_address() -> String {
    range(FIRST_ITEM, PRINCIPAL_COL, FIRST_ITEM + ITEM_COUNT - 1, COVER_COL)
}

/// A1 range of the formula column (share).
pub fn share_address() -> String {
    range(FIRST_ITEM, SHARE_COL, FIRST_ITEM + ITEM_COUNT - 1, SHARE_COL)
}

/// A1 range of the total row.
pub fn total_address() -> String {
    range(TOTAL, ITEM_COL, TOTAL, SHARE_COL)
}

/// A1 range of the principal amounts.
pub fn principal_address() -> String {
    range(FIRST_ITEM, PRINCIPAL_COL, FIRST_ITEM + ITEM_COUNT - 1, PRINCIPAL_COL)
}

/// A1 range of the rate column (a fraction to show as a percent).
pub fn rate_address() -> String {
    range(FIRST_ITEM, RATE_COL, FIRST_ITEM + ITEM_COUNT - 1, RATE_COL)
}

/// A1 range of the maturity column (date numbers).
pub fn maturity_address() -> String {
    range(FIRST_ITEM, MATURITY_COL, FIRST_ITEM + ITEM_COUNT - 1, MATURITY_COL)
}

/// A1 range of the cover column (multiples).
pub fn cover_address() -> String {
    range(FIRST_ITEM, COVER_COL, FIRST_ITEM + ITEM_COUNT - 1, COVER_COL)
}

/// A1 range of the item names: what the indent, align and underline cycles act on.
pub fn items_address() -> String {
    range(FIRST_ITEM, ITEM_COL, FIRST_ITEM + ITEM_COUNT - 1, ITEM_COL)
}

/// A1 addresses of the two samples: slot 1 text, slot 2 number.
pub fn sample_addresses() -> (String, String) {
    (cell(FIRST_ITEM, SAMPLE_COL), cell(FIRST_ITEM + 1, SAMPLE_COL))
}
