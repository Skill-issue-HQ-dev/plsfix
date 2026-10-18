// The Templates sheet: an index of the eighteen calculation blocks the pane's picker inserts, and
// below it one labelled EMPTY pad per block, sized to the block (a block needs an empty rectangle
// at the active cell). Invariant: no pad holds a cell, no two pads touch, and every size here is
// the registry's (test/demo-template-pads.test.ts reads the PADS table below against it).

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{cell, column, range, templates::*, GUIDE_ROWS, LABEL_COL, TEMPLATES_SHEET, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = TEMPLATES_SHEET;
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const GUIDE: Guide = Guide {
    title: "all eighteen calculation blocks, one empty pad each",
    tasks: &[
        "Click a template's name in the list: Excel jumps to the first cell of its pad.",
        "Press its Insert button (Tools tab, Templates): the block lands there, inputs blue, formulas black.",
        "Undo last pls,fix action takes a block back; Football field reads its whole block, Waterfall the rows from Opening to Closing EBITDA (the pad notes name them).",
    ],
    commands: &["Open pls,fix (Ctrl+Shift+M)", "Undo last pls,fix action (Ctrl+Shift+Z)"],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const NUMBER_COL: u16 = 0;
const SIZE_COL: u16 = 4;
const NUMBER_WIDTH: f64 = 5.0;
const NAME_WIDTH: f64 = 38.0;
const BLOCK_WIDTH: f64 = 13.0;
const LAST_BLOCK_COL: u16 = 9;

/// One block the picker inserts: its registry id, name and group, and its size in rows and columns.
pub struct Pad {
    pub id: &'static str,
    pub name: &'static str,
    pub group: &'static str,
    pub rows: u32,
    pub cols: u16,
    /// What to do after the insert, or "". `{rows}` stands for the rows a Waterfall reads in the block.
    pub then: &'static str,
}

/// The registry's order and sizes (src/templates/index.ts).
pub const PADS: [Pad; COUNT as usize] = [
    Pad { id: "dcf", name: "DCF valuation", group: "Valuation", rows: 13, cols: 6, then: "" },
    Pad { id: "wacc", name: "WACC", group: "Valuation", rows: 21, cols: 2, then: "" },
    Pad { id: "trading-comps", name: "Trading comps", group: "Valuation", rows: 25, cols: 9, then: "Change a multiple and watch the statistics and the implied range move." },
    Pad { id: "precedents", name: "Precedent transactions", group: "Valuation", rows: 19, cols: 6, then: "" },
    Pad { id: "football-field", name: "Football field data", group: "Valuation", rows: 6, cols: 4, then: "Then select the whole block and press Football field; the three columns to its right stay empty." },
    Pad { id: "npv-irr", name: "NPV / IRR", group: "Valuation", rows: 17, cols: 3, then: "" },
    Pad { id: "debt-schedule", name: "Debt schedule (annuity)", group: "LBO", rows: 20, cols: 6, then: "" },
    Pad { id: "sources-uses", name: "Sources and uses", group: "LBO", rows: 22, cols: 3, then: "" },
    Pad { id: "debt-paydown", name: "Debt paydown (cash sweep)", group: "LBO", rows: 18, cols: 6, then: "" },
    Pad { id: "returns", name: "Returns (MOIC / IRR)", group: "LBO", rows: 17, cols: 7, then: "" },
    Pad { id: "cap-table", name: "Cap table", group: "LBO", rows: 23, cols: 6, then: "" },
    Pad { id: "working-capital", name: "Working capital days", group: "Operating", rows: 12, cols: 2, then: "" },
    Pad { id: "revenue-build", name: "Revenue build", group: "Operating", rows: 11, cols: 6, then: "" },
    Pad { id: "capex-depreciation", name: "Capex and depreciation", group: "Operating", rows: 15, cols: 6, then: "" },
    Pad { id: "three-statement", name: "Three-statement mini", group: "Operating", rows: 26, cols: 5, then: "" },
    Pad { id: "scenario-switch", name: "Scenario switch", group: "Operating", rows: 16, cols: 5, then: "Type 1, 2 or 3 in the selector and the Live column follows." },
    Pad { id: "sensitivity", name: "Two-way sensitivity", group: "Analysis", rows: 11, cols: 6, then: "" },
    Pad { id: "ebitda-bridge", name: "EBITDA bridge", group: "Analysis", rows: 9, cols: 2, then: "Then select {rows} (Opening to Closing EBITDA) and press Waterfall from selection." },
];
const _: () = assert!(PADS.len() as u32 == COUNT, "the index and the pads list the same blocks");

/// The label row of pad `index`; its block starts on the next row.
pub fn label_row(index: usize) -> u32 {
    FIRST_PAD + PADS[..index].iter().map(|pad| 1 + pad.rows + PAD_GAP).sum::<u32>()
}

/// The first cell of pad `index`'s block: where the user clicks.
pub fn anchor(index: usize) -> String {
    cell(label_row(index) + 1, ANCHOR_COL)
}

/// The rectangle of pad `index`'s block: first row, first column, last row, last column.
pub fn pad_rect(index: usize) -> (u32, u16, u32, u16) {
    let first = label_row(index) + 1;
    (first, ANCHOR_COL, first + PADS[index].rows - 1, ANCHOR_COL + PADS[index].cols - 1)
}

/// The rows a Waterfall reads in pad `index`: the block without its title row (whose second cell is
/// empty, so the chart refuses it) and without its check row, both columns.
fn waterfall_rows(index: usize) -> String {
    let (first_row, first_col, last_row, last_col) = pad_rect(index);
    range(first_row + 1, first_col, last_row - 1, last_col)
}

/// "13 x 6": a block's size as the index prints it.
pub fn size_words(pad: &Pad) -> String {
    format!("{} x {}", pad.rows, pad.cols)
}

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "Templates: eighteen blocks, one empty pad each", &styles.title)?;
    let note = format!("The picker writes a block at the active cell and needs that rectangle empty: every pad below is, and is sized to its block (columns B to {}).", column(LAST_BLOCK_COL));
    pen.text(NOTE_ROW, LABEL_COL, &note, &styles.note)?;
    write_index(&mut pen, styles)?;
    for (index, pad) in PADS.iter().enumerate() {
        write_label(&mut pen, styles, index, pad)?;
    }
    let sheet = pen.sheet();
    sheet.set_column_width(NUMBER_COL, NUMBER_WIDTH)?;
    sheet.set_column_width(ANCHOR_COL, NAME_WIDTH)?;
    sheet.set_column_range_width(ANCHOR_COL + 1, LAST_BLOCK_COL, BLOCK_WIDTH)?;
    tally += pen.tally;
    Ok(tally)
}

fn write_index(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(INDEX_HEADER, NUMBER_COL, "#", &styles.header_left)?;
    pen.text(INDEX_HEADER, ANCHOR_COL, "Template (click to jump to its pad)", &styles.header_left)?;
    pen.text(INDEX_HEADER, ANCHOR_COL + 1, "Group", &styles.header_left)?;
    pen.text(INDEX_HEADER, ANCHOR_COL + 2, "Pad", &styles.header_left)?;
    pen.text(INDEX_HEADER, SIZE_COL, "Rows x columns", &styles.header_left)?;
    for (index, pad) in PADS.iter().enumerate() {
        let row = INDEX_FIRST + index as u32;
        pen.number(row, NUMBER_COL, (index + 1) as f64, &styles.whole)?;
        let target = format!("internal:'{NAME}'!{}", anchor(index));
        pen.link(row, ANCHOR_COL, &target, pad.name, &styles.link)?;
        pen.text(row, ANCHOR_COL + 1, pad.group, &styles.label)?;
        pen.text(row, ANCHOR_COL + 2, &anchor(index), &styles.label)?;
        pen.text(row, SIZE_COL, &size_words(pad), &styles.label)?;
    }
    Ok(())
}

fn write_label(pen: &mut Pen, styles: &Styles, index: usize, pad: &Pad) -> Result<(), XlsxError> {
    let row = label_row(index);
    pen.number(row, NUMBER_COL, (index + 1) as f64, &styles.whole)?;
    pen.text(row, ANCHOR_COL, pad.name, &styles.label_bold)?;
    let (_, _, last_row, last_col) = pad_rect(index);
    let mut hint = format!(
        "Click {} and press Insert on the {} card; the block fills {}.",
        anchor(index),
        pad.name,
        range(row + 1, ANCHOR_COL, last_row, last_col)
    );
    if !pad.then.is_empty() {
        hint.push(' ');
        hint.push_str(&pad.then.replace("{rows}", &waterfall_rows(index)));
    }
    pen.text(row, ANCHOR_COL + 1, &hint, &styles.note)?;
    Ok(())
}
