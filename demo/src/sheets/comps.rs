// The Comps sheet: eight invented peers with live multiples for Comps stats (header row, two or
// more data rows, six empty rows under the table), and a Method | Low | High block for Football
// field with its three right-hand columns left empty. Invariant: the rows Comps stats writes into
// and the columns the football helper block lands in hold no cell at all.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{cell, column, comps::*, range, COMPS_SHEET, GUIDE_ROWS, LABEL_COL, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = COMPS_SHEET;
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const GUIDE: Guide = Guide {
    title: "peer multiples and a valuation range",
    tasks: &[
        "Select the peer table with its header row and press Comps stats: six live statistic rows land under it.",
        "Select the Method | Low | High block and press Football field: one floating bar per method.",
    ],
    commands: &["Open pls,fix (Ctrl+Shift+M)", "Undo last pls,fix action (Ctrl+Shift+Z)"],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const REVENUE_COL: u16 = 1;
const EBITDA_COL: u16 = 2;
const EV_COL: u16 = 3;
const EV_REVENUE_COL: u16 = 4;
const EV_EBITDA_COL: u16 = 5;
const MARGIN_COL: u16 = 6;
const LABEL_WIDTH: f64 = 40.0;
const NUMBER_WIDTH: f64 = 14.0;

/// (company, revenue, EBITDA, enterprise value) in EUR millions: invented peers of a Baltic
/// bakery group, EV / EBITDA between 5.1x and 9.0x.
pub const PEERS: [(&str, f64, f64, f64); PEER_COUNT as usize] = [
    ("Nordic Grain Group", 182.0, 27.3, 191.7),
    ("Baltic Bakeries", 96.5, 12.1, 85.5),
    ("Kurzeme Mills", 54.2, 6.5, 33.8),
    ("Rye & Co", 210.8, 33.7, 262.4),
    ("Vidzeme Dairy", 143.0, 18.6, 118.9),
    ("Daugava Foods", 78.4, 8.2, 41.8),
    ("Riga Pastry Group", 61.7, 9.9, 89.1),
    ("Lielupe Catering", 120.3, 14.4, 98.6),
];

/// (method, low, high): DemoCo's enterprise value in EUR millions, around 2.1 of EBITDA.
pub const METHODS: [(&str, f64, f64); FOOT_COUNT as usize] = [
    ("Trading comps, 25th to 75th EV / EBITDA", 12.9, 16.1),
    ("Precedent transactions", 11.8, 17.1),
    ("DCF, WACC 9% to 11%", 11.2, 15.9),
    ("LBO, 20% to 25% IRR", 10.4, 13.8),
    ("Broker targets", 12.0, 14.6),
    ("Management plan, exit at 7x to 9x", 12.4, 15.2),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "DemoCo SIA: peer multiples and a valuation range (EUR millions)", &styles.title)?;
    let note = format!(
        "Select {} and press Comps stats; select {} and press Football field (keep columns {} to {} beside it empty).",
        table_address(),
        football_address(),
        column(FOOT_COLS),
        column(FOOT_COLS + 2)
    );
    pen.text(NOTE_ROW, LABEL_COL, &note, &styles.note)?;
    write_peers(&mut pen, styles)?;
    write_methods(&mut pen, styles)?;
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, LABEL_WIDTH)?;
    sheet.set_column_range_width(REVENUE_COL, LAST_COL, NUMBER_WIDTH)?;
    tally += pen.tally;
    Ok(tally)
}

fn write_peers(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    let titles = ["Company", "Revenue", "EBITDA", "EV", "EV / Revenue", "EV / EBITDA", "EBITDA margin"];
    for (i, title) in titles.iter().enumerate() {
        let format = if i == 0 { &styles.header_left } else { &styles.header };
        pen.text(HEADER, i as u16, title, format)?;
    }
    for (i, (company, revenue, ebitda, ev)) in PEERS.iter().enumerate() {
        let row = FIRST_PEER + i as u32;
        pen.text(row, LABEL_COL, company, &styles.label)?;
        pen.number(row, REVENUE_COL, *revenue, &styles.number_1dp)?;
        pen.number(row, EBITDA_COL, *ebitda, &styles.number_1dp)?;
        pen.number(row, EV_COL, *ev, &styles.number_1dp)?;
        let (rev, ebi, ev_ref) = (cell(row, REVENUE_COL), cell(row, EBITDA_COL), cell(row, EV_COL));
        pen.formula(row, EV_REVENUE_COL, &format!("={ev_ref}/{rev}"), &styles.multiple)?;
        pen.formula(row, EV_EBITDA_COL, &format!("={ev_ref}/{ebi}"), &styles.multiple)?;
        pen.formula(row, MARGIN_COL, &format!("={ebi}/{rev}"), &styles.pct)?;
    }
    Ok(())
}

fn write_methods(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(FOOT_TITLE, LABEL_COL, "Football field input: DemoCo SIA enterprise value by method (EUR millions)", &styles.label_bold)?;
    pen.text(FOOT_HEADER, LABEL_COL, "Method", &styles.header_left)?;
    pen.text(FOOT_HEADER, 1, "Low", &styles.header)?;
    pen.text(FOOT_HEADER, 2, "High", &styles.header)?;
    for (i, (method, low, high)) in METHODS.iter().enumerate() {
        let row = FOOT_FIRST + i as u32;
        pen.text(row, LABEL_COL, method, &styles.label)?;
        pen.number(row, 1, *low, &styles.number_1dp)?;
        pen.number(row, 2, *high, &styles.number_1dp)?;
    }
    Ok(())
}

/// A1 range of the peer table with its header row: the Comps stats input.
pub fn table_address() -> String {
    range(HEADER, LABEL_COL, LAST_PEER, LAST_COL)
}

/// A1 range of the Method | Low | High block with its header row: the Football field input.
pub fn football_address() -> String {
    range(FOOT_HEADER, LABEL_COL, FOOT_FIRST + FOOT_COUNT - 1, FOOT_COLS - 1)
}

/// "rows N to M": the six rows Comps stats writes into (after one blank row under the table).
pub fn stats_rows_words() -> String {
    format!("rows {} to {}", STATS_FIRST + 1, STATS_FIRST + STATS_ROWS)
}

/// "columns B to G": the number columns one Comps stats press writes under (column A holds the names).
pub fn number_columns_words() -> String {
    format!("columns {} to {}", column(LABEL_COL + 1), column(LAST_COL))
}
