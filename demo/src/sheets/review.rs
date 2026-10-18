// The Review sheet: what a reviewer flags and the tools that answer it: a division by zero (Model
// check lists it, IFERROR guard wraps it), a formula in the first row only (Fill formula down),
// TODAY, OFFSET and a typed 1.15, and one formatted blank at AD800 (Clean past the data).
// Invariant: no chart, shape or table sits here, so the clean-up deletes rows, not just formats.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{cell, range, review::*, GUIDE_ROWS, LABEL_COL, REVIEW_SHEET, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = REVIEW_SHEET;
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const GUIDE: Guide = Guide {
    title: "what a reviewer flags, and the tools that answer it",
    tasks: &[
        "Run model check (Workbook tab): it lists the division by zero, the typed numbers, the volatile functions, the broken name, the hidden sheet.",
        "Select the Average price column and press IFERROR guard; select the one Margin formula and press Fill formula down.",
        "Ctrl+End jumps to row 800 here: press Clean past the data twice and it lands on the corner of the data.",
    ],
    commands: &[
        "Wrap with IFERROR (Ctrl+Shift+I)",
        "Fill formula down (Ctrl+Alt+D)",
        "Undo last pls,fix action (Ctrl+Shift+Z)",
    ],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const PRODUCT_COL: u16 = 0;
const REVENUE_COL: u16 = 1;
const SECOND_COL: u16 = 2;
const RESULT_COL: u16 = 3;
const PRODUCT_WIDTH: f64 = 30.0;
const COLUMN_WIDTH: f64 = 14.0;

/// (product, revenue, units): the fourth product has sold nothing yet, so its average price divides by zero.
const SALES: [(&str, f64, f64); PRODUCTS as usize] = [
    ("Rye bread", 3_480.0, 1_160.0),
    ("Wheat bread", 3_150.0, 1_050.0),
    ("Pastries", 2_420.0, 605.0),
    ("Crispbread", 1_640.0, 0.0),
    ("Cakes", 1_330.0, 190.0),
];

/// (product, revenue, cost) in EUR thousands.
const COSTS: [(&str, f64, f64); PRODUCTS as usize] = [
    ("Rye bread", 3_480.0, 2_610.0),
    ("Wheat bread", 3_150.0, 2_395.0),
    ("Pastries", 2_420.0, 1_815.0),
    ("Crispbread", 1_640.0, 1_230.0),
    ("Cakes", 1_330.0, 905.0),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "Review: what a reviewer flags", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &note(), &styles.note)?;
    write_prices(&mut pen, styles)?;
    write_margins(&mut pen, styles)?;
    write_flags(&mut pen, styles)?;
    pen.blank(STRAY_ROW, STRAY_COL, &styles.stray)?;
    let sheet = pen.sheet();
    sheet.set_column_width(PRODUCT_COL, PRODUCT_WIDTH)?;
    sheet.set_column_range_width(REVENUE_COL, RESULT_COL, COLUMN_WIDTH)?;
    tally += pen.tally;
    Ok(tally)
}

fn note() -> String {
    format!(
        "Average price {} divides by zero on one row; Margin {} holds its formula in the first row only; {} is formatted but empty.",
        price_address(),
        margin_first_address(),
        stray_address()
    )
}

fn write_prices(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(PRICE_LABEL, PRODUCT_COL, "Average price by product (one product has no units yet)", &styles.label_bold)?;
    header(pen, styles, PRICE_HEADER, ["Product", "Revenue", "Units", "Average price"])?;
    for (i, (product, revenue, units)) in SALES.iter().enumerate() {
        let row = PRICE_FIRST + i as u32;
        pen.text(row, PRODUCT_COL, product, &styles.label)?;
        pen.number(row, REVENUE_COL, *revenue, &styles.eur)?;
        pen.number(row, SECOND_COL, *units, &styles.eur)?;
        let formula = format!("={}/{}", cell(row, REVENUE_COL), cell(row, SECOND_COL));
        pen.formula(row, RESULT_COL, &formula, &styles.two_dp)?;
    }
    Ok(())
}

fn write_margins(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(MARGIN_LABEL, PRODUCT_COL, "Margin by product (only the first row has its formula)", &styles.label_bold)?;
    header(pen, styles, MARGIN_HEADER, ["Product", "Revenue", "Cost", "Margin"])?;
    for (i, (product, revenue, cost)) in COSTS.iter().enumerate() {
        let row = MARGIN_FIRST + i as u32;
        pen.text(row, PRODUCT_COL, product, &styles.label)?;
        pen.number(row, REVENUE_COL, *revenue, &styles.eur)?;
        pen.number(row, SECOND_COL, *cost, &styles.eur)?;
    }
    let (revenue, cost) = (cell(MARGIN_FIRST, REVENUE_COL), cell(MARGIN_FIRST, SECOND_COL));
    pen.formula(MARGIN_FIRST, RESULT_COL, &format!("=({revenue}-{cost})/{revenue}"), &styles.pct)?;
    Ok(())
}

fn write_flags(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(FLAGS_LABEL, PRODUCT_COL, "Formulas a reviewer writes down (Model check lists them)", &styles.label_bold)?;
    let first_revenue = cell(MARGIN_FIRST, REVENUE_COL);
    let flags = [
        ("Report date (volatile)", "=TODAY()".to_string(), &styles.date),
        ("Next three months (volatile, typed 3)", format!("=SUM(OFFSET({first_revenue},0,0,3,1))"), &styles.eur),
        ("List price with markup (typed 1.15)", format!("={first_revenue}*1.15"), &styles.eur),
    ];
    for (i, (label, formula, format)) in flags.iter().enumerate() {
        let row = FLAGS_FIRST + i as u32;
        pen.text(row, PRODUCT_COL, label, &styles.label)?;
        pen.formula(row, REVENUE_COL, formula, format)?;
    }
    Ok(())
}

fn header(pen: &mut Pen, styles: &Styles, row: u32, titles: [&str; 4]) -> Result<(), XlsxError> {
    for (i, title) in titles.iter().enumerate() {
        let format = if i == 0 { &styles.header_left } else { &styles.header };
        pen.text(row, i as u16, title, format)?;
    }
    Ok(())
}

/// A1 range of the Average price column (the one holding the division by zero).
pub fn price_address() -> String {
    range(PRICE_FIRST, RESULT_COL, PRICE_FIRST + PRODUCTS - 1, RESULT_COL)
}

/// A1 address of the average price that divides by zero (the product with no units).
pub fn error_address() -> String {
    let zero_units = SALES.iter().position(|(_, _, units)| *units == 0.0).expect("one product has no units");
    cell(PRICE_FIRST + zero_units as u32, RESULT_COL)
}

/// A1 address of the one Margin formula, the cell Fill formula down starts from.
pub fn margin_first_address() -> String {
    cell(MARGIN_FIRST, RESULT_COL)
}

/// A1 address of the formatted blank far from the data.
pub fn stray_address() -> String {
    cell(STRAY_ROW, STRAY_COL)
}

/// A1 address where Ctrl+End lands once the surplus is cleaned: the corner of the data, where the
/// last data row meets the last data column. Nothing sits in that cell itself (the flags below the
/// tables fill two columns only).
pub fn data_corner_address() -> String {
    cell(FLAGS_FIRST + FLAGS - 1, RESULT_COL)
}

/// The sheet row of the stray cell, for the text that says where Ctrl+End lands.
pub fn stray_row_number() -> u32 {
    STRAY_ROW + 1
}

/// "rows N to M": the Margin rows Fill formula down completes.
pub fn margin_rows_words() -> String {
    format!("rows {} to {}", MARGIN_FIRST + 1, MARGIN_FIRST + PRODUCTS)
}
