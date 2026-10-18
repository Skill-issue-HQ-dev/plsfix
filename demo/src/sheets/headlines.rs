// The Headlines sheet: three sentences built from the P&L, the source for Export as text. Each
// cell is a formula, so the text a slide shows moves with the model; the cells are plain text
// results (ROUND and "&", no TEXT format codes, so the sentence reads right in any locale).
// Invariant: the longest sentence stays under the 500 characters a text link carries.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{
    cell, headlines::*, last_year_col, pnl::*, pnl_ref, year_col, FIRST_YEAR_COL, GUIDE_ROWS, HEADLINES_SHEET,
    LABEL_COL, TITLE_ROW,
};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = HEADLINES_SHEET;
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const GUIDE: Guide = Guide {
    title: "slide text that follows the model",
    tasks: &[
        "Select one text cell and press Export as text (Links tab); in PowerPoint insert it on Slide 5, Where = Top left.",
        "Change the first growth rate on Assumptions to 12%, press Copy all (Push all with the relay), paste it into PowerPoint's Inbox box, then Update all in PowerPoint: the sentence redraws.",
    ],
    commands: &["Open pls,fix (Ctrl+Shift+M)"],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const TEXT_COL: u16 = 1;
const LABEL_WIDTH: f64 = 22.0;
const TEXT_WIDTH: f64 = 100.0;
const PARAGRAPH_HEIGHT: f64 = 48.0;
/// The most characters one text link carries (`TEXT_MAX_CHARS` in src/link/model.ts).
pub const TEXT_MAX_CHARS: usize = 500;

/// A rounded P&L figure for a sentence: `scale` turns it into percent points or EUR millions.
fn figure(row: u32, year: usize, scale: &str, decimals: u8) -> String {
    format!("ROUND({}{scale},{decimals})", pnl_ref(row, year_col(year)))
}

/// (label, formula): the three sentences, in the order of the rows.
pub fn sentences() -> [(&'static str, String); COUNT as usize] {
    let last = (last_year_col() - FIRST_YEAR_COL) as usize;
    let title = format!(
        "=\"Revenue grows \"&{}&\"% to EUR \"&{}&\" million in 2025E.\"",
        figure(GROWTH, 1, "*100", 1),
        figure(REVENUE, 1, "/1000", 1)
    );
    let margin = format!(
        "=\"EBITDA margin moves from \"&{}&\"% in 2024A to \"&{}&\"% in 2029E.\"",
        figure(EBITDA_MARGIN, 0, "*100", 1),
        figure(EBITDA_MARGIN, last, "*100", 1)
    );
    let paragraph = format!(
        "=\"DemoCo SIA made EUR \"&{}&\" million of revenue in 2024A and plans EUR \"&{}&\" million in 2025E and EUR \"&{}&\" million in 2029E. EBITDA rises from EUR \"&{}&\" million to EUR \"&{}&\" million, and net income reaches EUR \"&{}&\" million by 2029E.\"",
        figure(REVENUE, 0, "/1000", 1),
        figure(REVENUE, 1, "/1000", 1),
        figure(REVENUE, last, "/1000", 1),
        figure(EBITDA, 0, "/1000", 2),
        figure(EBITDA, last, "/1000", 2),
        figure(NET_INCOME, last, "/1000", 2)
    );
    [("Slide title", title), ("Margin line", margin), ("Summary paragraph", paragraph)]
}

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "Headlines: sentences that follow the P&L", &styles.title)?;
    let note = format!("Select one cell in {} and press Export as text on the Links tab: a text box on the slide, refreshed in place.", text_column());
    pen.text(NOTE_ROW, LABEL_COL, &note, &styles.note)?;
    pen.text(HEADER, LABEL_COL, "Where it goes", &styles.header_left)?;
    pen.text(HEADER, TEXT_COL, "Text (a formula, so it follows the model)", &styles.header_left)?;
    for (i, (label, formula)) in sentences().iter().enumerate() {
        let row = FIRST + i as u32;
        let last = i + 1 == COUNT as usize;
        pen.text(row, LABEL_COL, label, &styles.label)?;
        pen.formula(row, TEXT_COL, formula, if last { &styles.wrap } else { &styles.plain })?;
    }
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, LABEL_WIDTH)?;
    sheet.set_column_width(TEXT_COL, TEXT_WIDTH)?;
    sheet.set_row_height(FIRST + COUNT - 1, PARAGRAPH_HEIGHT)?;
    tally += pen.tally;
    Ok(tally)
}

/// A1 range of the three text cells.
pub fn text_column() -> String {
    format!("{}:{}", cell(FIRST, TEXT_COL), cell(FIRST + COUNT - 1, TEXT_COL))
}

/// A1 address of the text cell `i` (0 = the slide title).
pub fn text_cell(i: u32) -> String {
    cell(FIRST + i, TEXT_COL)
}
