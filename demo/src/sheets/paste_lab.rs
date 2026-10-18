// The Paste lab: one source block to Mark (header 30 pt, rows of different heights, four number
// formats, formulas that read inside the block and one input outside it), five empty pads for the
// five formula pastes, and a plain block that only takes number formats and row heights.
// Invariant: every pad and the plain block are the source's size, and the pads hold no cell.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{cell, cell_abs, paste_lab::*, range, GUIDE_ROWS, LABEL_COL, PASTE_SHEET, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = PASTE_SHEET;
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const GUIDE: Guide = Guide {
    title: "mark a block, paste it six ways",
    tasks: &[
        "Mark the source block, click a pad's first cell and press its paste button: values, formats, exact, duplicate, transposed.",
        "Compare Paste formulas exactly with Paste: duplicate formulas: the VAT formulas read different rows.",
        "Paste: number formats and Paste: row heights act on the plain block at the bottom.",
    ],
    commands: &[
        "Mark copy source (Ctrl+Shift+C)",
        "Paste values (Ctrl+Shift+V)",
        "Paste formats (Ctrl+Alt+F)",
        "Paste formulas exactly (Ctrl+Alt+P)",
        "Paste transposed (Ctrl+Alt+T)",
    ],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const VAT: f64 = 0.21;
const HEADER_HEIGHT: f64 = 30.0;
/// The item rows' heights, so Paste: row heights has something to copy.
const ROW_HEIGHTS: [f64; SRC_ROWS as usize] = [21.0, 24.0, 21.0, 24.0];
const ITEM_WIDTH: f64 = 18.0;
const COLUMN_WIDTH: f64 = 13.0;
const TARGET_NAMES: [&str; TARGETS as usize] = [
    "Paste values: the numbers, no formulas or formats",
    "Paste formats: the look, no numbers",
    "Paste formulas exactly: the formula text unchanged, so it still reads the source rows",
    "Paste: duplicate formulas: references inside the block move with it, $B$10 outside stays",
    "Paste transposed: rows become columns",
];

/// (item, quantity, price): the invented order lines.
const LINES: [(&str, f64, f64); SRC_ROWS as usize] = [
    ("Rye bread", 120.0, 2.45),
    ("Wheat bread", 80.0, 3.10),
    ("Pastries", 200.0, 0.95),
    ("Cakes", 45.0, 12.50),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "Paste lab: one block, six pastes", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &note(), &styles.note)?;
    pen.text(VAT_ROW, LABEL_COL, "VAT rate", &styles.label_bold)?;
    pen.number(VAT_ROW, 1, VAT, &styles.pct)?;
    pen.text(VAT_ROW, 2, "An input outside the source block: the VAT column reads it with $B$10.", &styles.note)?;
    write_source(&mut pen, styles)?;
    for (k, name) in TARGET_NAMES.iter().enumerate() {
        let label = FIRST_TARGET_LABEL + TARGET_STRIDE * k as u32;
        pen.text(label, LABEL_COL, &format!("{}. {name}. Click {} first.", k + 1, target_anchor(k)), &styles.label_bold)?;
    }
    write_raw(&mut pen, styles)?;
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, ITEM_WIDTH)?;
    sheet.set_column_range_width(1, COLS - 1, COLUMN_WIDTH)?;
    sheet.set_row_height(SRC_HEADER, HEADER_HEIGHT)?;
    for (i, height) in ROW_HEIGHTS.iter().enumerate() {
        sheet.set_row_height(SRC_FIRST + i as u32, *height)?;
    }
    tally += pen.tally;
    Ok(tally)
}

fn note() -> String {
    format!(
        "Mark {}, click the first cell of a pad below ({}, {}, ...) and press its button; {} takes formats and heights.",
        source_address(),
        target_anchor(0),
        target_anchor(1),
        raw_address()
    )
}

fn write_source(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(SRC_HEADER, LABEL_COL, "Item", &styles.header_left)?;
    for (i, title) in ["Qty", "Price", "Net", "VAT"].iter().enumerate() {
        pen.text(SRC_HEADER, 1 + i as u16, title, &styles.header)?;
    }
    let vat = cell_abs(VAT_ROW, 1);
    for (i, (item, quantity, price)) in LINES.iter().enumerate() {
        let row = SRC_FIRST + i as u32;
        pen.text(row, LABEL_COL, item, &styles.label)?;
        pen.number(row, 1, *quantity, &styles.whole)?;
        pen.number(row, 2, *price, &styles.points)?;
        pen.formula(row, 3, &format!("={}*{}", cell(row, 1), cell(row, 2)), &styles.two_dp)?;
        pen.formula(row, 4, &format!("={}*{vat}", cell(row, 3)), &styles.euro_2dp)?;
    }
    Ok(())
}

/// The plain twin of the source: same text, the numbers typed with no format at all.
fn write_raw(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    let label = "Plain numbers: select the whole block for Paste: number formats; click its first cell for Paste: row heights.";
    pen.text(RAW_LABEL, LABEL_COL, label, &styles.label_bold)?;
    let header = RAW_LABEL + 1;
    for (i, title) in ["Item", "Qty", "Price", "Net", "VAT"].iter().enumerate() {
        pen.text(header, i as u16, title, &styles.plain)?;
    }
    for (i, (item, quantity, price)) in LINES.iter().enumerate() {
        let row = header + 1 + i as u32;
        let net = quantity * price;
        pen.text(row, LABEL_COL, item, &styles.plain)?;
        pen.number(row, 1, *quantity, &styles.plain)?;
        pen.number(row, 2, *price, &styles.plain)?;
        pen.number(row, 3, net, &styles.plain)?;
        pen.number(row, 4, net * VAT, &styles.plain)?;
    }
    Ok(())
}

/// A1 range of the source block, header included: what Mark takes.
pub fn source_address() -> String {
    range(SRC_HEADER, LABEL_COL, SRC_FIRST + SRC_ROWS - 1, COLS - 1)
}

/// A1 address of the VAT input the source's last column reads.
pub fn vat_address() -> String {
    cell(VAT_ROW, 1)
}

/// The first cell of pad `k` (0 = values ... 4 = transposed).
pub fn target_anchor(k: usize) -> String {
    cell(FIRST_TARGET_LABEL + TARGET_STRIDE * k as u32 + 1, LABEL_COL)
}

/// A1 address of the plain block's first cell: where Paste: row heights starts.
pub fn raw_corner() -> String {
    cell(RAW_LABEL + 1, LABEL_COL)
}

/// The source's first VAT formula as typed, and the same formula as pad `k` would hold it after
/// a duplicate paste: the cells inside the block move with it, `$B$10` stays.
pub fn vat_formulas(k: usize) -> (String, String) {
    let vat = cell_abs(VAT_ROW, 1);
    let (first_item, _) = target_rows(k);
    (
        format!("={}*{vat}", cell(SRC_FIRST, 3)),
        format!("={}*{vat}", cell(first_item + 1, 3)),
    )
}

/// A1 range of the plain block, header included.
pub fn raw_address() -> String {
    range(RAW_LABEL + 1, LABEL_COL, RAW_LABEL + BLOCK_ROWS, COLS - 1)
}

/// The rectangle pad `k` must keep empty: first row, last row (columns are A to E).
pub fn target_rows(k: usize) -> (u32, u32) {
    let first = FIRST_TARGET_LABEL + TARGET_STRIDE * k as u32 + 1;
    (first, first + BLOCK_ROWS - 1)
}
