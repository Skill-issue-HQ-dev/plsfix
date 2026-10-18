// The "Start here" sheet: the play-around checklist, one row per tool with where to click, what to
// do and what should happen, in groups; a sheet named in the Where column is a link to it. Every
// address comes from the layout or the sheet that owns it, so a moved block cannot go stale here.
// The rows live in start_model, start_format, start_blocks and start_links.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{GUIDE_ROWS, LABEL_COL, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::sheets::{scratch, start_blocks, start_format, start_links, start_model, SHEETS};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "Start here";
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const TABLE_HEADER_ROW: u32 = GUIDE_ROWS + 3;
const FIRST_STEP_ROW: u32 = GUIDE_ROWS + 4;
const GUIDE: Guide = Guide {
    title: "orientation and the full checklist",
    tasks: &[
        "Open the pane (Ctrl+Shift+M).",
        "Walk the checklist below, one row per tool, group by group.",
        "Open the demo deck \"pls,fix Demo Deck.pptx\" for the PowerPoint rows (slides 2 to 8).",
    ],
    commands: &["Open pls,fix (Ctrl+Shift+M)"],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const NUMBER_COL: u16 = 0;
const TOOL_COL: u16 = 1;
const PLACE_COL: u16 = 2;
const ACTION_COL: u16 = 3;
const OUTCOME_COL: u16 = 4;
const COLUMNS: [(&str, f64); 5] = [
    ("#", 4.0),
    ("Tool", 30.0),
    ("Where", 28.0),
    ("Do this", 92.0),
    ("You should see", 78.0),
];

/// One checklist row.
pub struct Step {
    pub tool: &'static str,
    pub place: String,
    pub action: String,
    pub outcome: String,
}

/// A row of the table: a group heading or a step.
pub enum Item {
    Group(&'static str),
    Step(Step),
}

pub fn step(tool: &'static str, place: impl Into<String>, action: impl Into<String>, outcome: impl Into<String>) -> Item {
    Item::Step(Step { tool, place: place.into(), action: action.into(), outcome: outcome.into() })
}

pub fn group(title: &'static str) -> Item {
    Item::Group(title)
}

/// Every row of the checklist, in reading order.
pub fn items() -> Vec<Item> {
    let mut items = start_model::items();
    items.extend(start_format::items());
    items.extend(start_blocks::items());
    items.extend(start_links::items());
    items
}

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "pls,fix: play-around workbook", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, "Ribbon tab \"pls,fix\" in Excel and in PowerPoint; the pane loads from the address in its manifest. DemoCo SIA is invented.", &styles.note)?;
    for (i, (title, _)) in COLUMNS.iter().enumerate() {
        pen.text(TABLE_HEADER_ROW, i as u16, title, &styles.header_left)?;
    }
    let mut number = 0;
    for (i, item) in items().iter().enumerate() {
        let row = FIRST_STEP_ROW + i as u32;
        match item {
            Item::Group(title) => {
                pen.text(row, TOOL_COL, title, &styles.label_bold)?;
            }
            Item::Step(step) => {
                number += 1;
                write_step(&mut pen, styles, row, number, step)?;
            }
        }
    }
    let sheet = pen.sheet();
    for (i, (_, width)) in COLUMNS.iter().enumerate() {
        sheet.set_column_width(i as u16, *width)?;
    }
    sheet.set_active(true);
    tally += pen.tally;
    Ok(tally)
}

fn write_step(pen: &mut Pen, styles: &Styles, row: u32, number: u32, step: &Step) -> Result<(), XlsxError> {
    pen.number(row, NUMBER_COL, f64::from(number), &styles.whole)?;
    pen.text(row, TOOL_COL, step.tool, &styles.label_bold)?;
    // A sheet's name in the Where column is a link to that sheet's first cell.
    if SHEETS.iter().any(|(name, _)| *name == step.place && *name != scratch::NAME) {
        pen.link(row, PLACE_COL, &format!("internal:'{}'!A1", step.place), &step.place, &styles.link)?;
    } else {
        pen.text(row, PLACE_COL, &step.place, &styles.label)?;
    }
    pen.text(row, ACTION_COL, &step.action, &styles.label)?;
    pen.text(row, OUTCOME_COL, &step.outcome, &styles.label)?;
    Ok(())
}
