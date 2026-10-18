// The Charts sheet: twelve banker charts, one of every kind the link export has to handle, each
// beside its own small DemoCo data block. Owns the geometry (one 18-row band per chart, data in
// columns A-D, the chart from column F), which the tests read back from here.
// Invariant: a chart never covers a written cell and never touches another chart.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{GUIDE_ROWS, LABEL_COL, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::charts_data::{Block, BLOCKS};
use crate::sheets::charts_kinds::{chart, Place};
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "Charts";
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const FIRST_BAND_ROW: u32 = GUIDE_ROWS + 3;
/// Rows from one block's heading to the next: the 15-row scatter block plus its header, its
/// heading and a spare row; the 300-pixel (15-row) charts fit inside.
pub const BAND_ROWS: u32 = 18;
pub const FIRST_SERIES_COL: u16 = 1;
const LAST_SERIES_COL: u16 = 3;
/// Column E stays empty, so no chart leans on a data column.
pub const CHART_COL: u16 = 5;
const LABEL_WIDTH: f64 = 20.0;
const SERIES_WIDTH: f64 = 14.0;
const GAP_WIDTH: f64 = 2.0;
const GUIDE: Guide = Guide {
    title: "twelve chart kinds to send to a deck",
    tasks: &[
        "Click a chart, press Export active chart on the Links tab, then insert it from the PowerPoint Inbox.",
        "Change a number in a block, press Copy all (Push all with the relay), paste it into PowerPoint's Inbox box, then Update all in PowerPoint: the chart redraws in place.",
    ],
    commands: &["Brand-format chart (Ctrl+Alt+G)", "CAGR label (Ctrl+Alt+K)"],
};
const _: () = assert!(
    GUIDE.tasks.len() <= guide::MAX_TASKS,
    "a fifth task would land on the Commands row"
);

// The geometry above holds only while every block fits its columns and its band.
const MAX_SERIES: usize = (LAST_SERIES_COL - FIRST_SERIES_COL + 1) as usize;
const MAX_LABELS: usize = BAND_ROWS as usize - 3;
const _: () = {
    let mut i = 0;
    while i < BLOCKS.len() {
        let block = &BLOCKS[i];
        assert!(
            block.series.len() <= MAX_SERIES,
            "a block has more series than columns B-D"
        );
        assert!(
            block.labels.len() <= MAX_LABELS,
            "a block is taller than its band"
        );
        assert!(
            block.note.is_empty() || block.labels.len() < MAX_LABELS,
            "a note would leave its band"
        );
        let mut s = 0;
        while s < block.series.len() {
            assert!(
                block.series[s].values.len() == block.labels.len(),
                "a series is not as long as its labels"
            );
            s += 1;
        }
        i += 1;
    }
};

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(
        TITLE_ROW,
        LABEL_COL,
        "DemoCo SIA: one chart of every kind (EUR thousands)",
        &styles.title,
    )?;
    pen.text(NOTE_ROW, LABEL_COL, "Each block feeds the chart beside it: click the chart and press Export active chart. DemoCo SIA is invented.", &styles.note)?;
    for (index, block) in BLOCKS.iter().enumerate() {
        write_block(&mut pen, styles, index, block)?;
    }
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, LABEL_WIDTH)?;
    sheet.set_column_range_width(FIRST_SERIES_COL, LAST_SERIES_COL, SERIES_WIDTH)?;
    sheet.set_column_width(LAST_SERIES_COL + 1, GAP_WIDTH)?;
    for (index, block) in BLOCKS.iter().enumerate() {
        sheet.insert_chart(band_row(index), CHART_COL, &chart(block, &place(index)))?;
    }
    tally += pen.tally;
    Ok(tally)
}

/// The heading row of block `index`, which is also the row its chart starts on.
pub fn band_row(index: usize) -> u32 {
    FIRST_BAND_ROW + BAND_ROWS * index as u32
}

/// Where block `index` sits: its header row and its first and last data row.
pub fn place(index: usize) -> Place {
    let first = band_row(index) + 2;
    Place {
        sheet: NAME,
        header: first - 1,
        first,
        last: first + BLOCKS[index].labels.len() as u32 - 1,
        label_col: LABEL_COL,
        series_col: FIRST_SERIES_COL,
    }
}

fn write_block(
    pen: &mut Pen,
    styles: &Styles,
    index: usize,
    block: &Block,
) -> Result<(), XlsxError> {
    let at = place(index);
    pen.text(
        band_row(index),
        LABEL_COL,
        block.heading,
        &styles.label_bold,
    )?;
    pen.text(at.header, LABEL_COL, block.corner, &styles.header_left)?;
    for (s, series) in block.series.iter().enumerate() {
        pen.text(
            at.header,
            FIRST_SERIES_COL + s as u16,
            series.name,
            &styles.header,
        )?;
    }
    for (i, label) in block.labels.iter().enumerate() {
        let row = at.first + i as u32;
        pen.text(row, LABEL_COL, label, &styles.label)?;
        for (s, series) in block.series.iter().enumerate() {
            let value = series.values[i];
            // An empty cell is simply not written: that is the gap the chart shows.
            if value.is_nan() {
                continue;
            }
            let format = if series.percent {
                &styles.pct
            } else {
                &styles.eur
            };
            pen.number(row, FIRST_SERIES_COL + s as u16, value, format)?;
        }
    }
    if !block.note.is_empty() {
        pen.text(at.last + 2, LABEL_COL, block.note, &styles.note)?;
    }
    Ok(())
}
