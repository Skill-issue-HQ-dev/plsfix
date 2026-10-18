// Reads the six block sheets back out of the saved workbook and holds each to what its tool needs: empty
// pads (and the rows the bridge pad names for Waterfall), Comps rows, empty paste targets, unformatted
// plain tables, headline sentences that fit a text link, the Review sheet's error, volatiles, stray cell
// and Ctrl+End corner, and no sheet over the scan cap. Pure reads of the saved file, never the builders.

mod common;

use common::{
    any_cell_in, archive, cell_formula, cell_number, cell_text, cell_xml, cells, shared_strings, sheet_xml, values_corner,
};
use smt_demo::layout::{
    cell, comps as cl, format_lab as fl, headlines as hl, paste_lab as pl, review as rl, templates as tl,
};
use smt_demo::sheets::{comps, format_lab, headlines, paste_lab, review, templates, SHEETS};

/// `SHEET_SCAN_CELL_CAP` in src/excel/internal.ts: Find, Model check and Prepare for sharing
/// read a sheet's whole used range only under it.
const SCAN_CAP: u64 = 200_000;

#[test]
fn every_template_pad_is_empty_and_apart() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, templates::NAME);
    assert_eq!(templates::PADS.len() as u32, tl::COUNT);
    for (index, pad) in templates::PADS.iter().enumerate() {
        let (first_row, first_col, last_row, last_col) = templates::pad_rect(index);
        assert!(!any_cell_in(&sheet, first_row, first_col, last_row, last_col), "{}: its pad holds a cell", pad.name);
        if let Some(next) = templates::PADS.get(index + 1) {
            let gap = templates::label_row(index + 1) - last_row - 1;
            assert!(gap >= tl::PAD_GAP, "{} and {}: {gap} blank rows between them", pad.name, next.name);
        }
    }
}

#[test]
fn the_football_pad_keeps_three_empty_columns_to_its_right() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, templates::NAME);
    let index = templates::PADS.iter().position(|pad| pad.id == "football-field").expect("football pad");
    let (first_row, _, last_row, last_col) = templates::pad_rect(index);
    // The helper block is one row taller than the data: a header plus the rows.
    assert!(!any_cell_in(&sheet, first_row, last_col + 1, last_row + 1, last_col + cl::FOOT_COLS), "helper columns");
}

#[test]
fn the_template_sheet_holds_only_its_index_and_labels() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, templates::NAME);
    let mut allowed: Vec<u32> = (0..=tl::INDEX_FIRST + tl::COUNT - 1).collect();
    allowed.extend((0..templates::PADS.len()).map(templates::label_row));
    for (row, col) in cells(&sheet) {
        assert!(allowed.contains(&row), "a cell at row {} column {col} is neither index nor label", row + 1);
    }
}

#[test]
fn each_index_entry_jumps_to_its_pad() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, templates::NAME);
    for (index, pad) in templates::PADS.iter().enumerate() {
        let row = tl::INDEX_FIRST + index as u32 + 1;
        let link = format!(
            "<hyperlink ref=\"B{row}\" location=\"'{}'!{}\" display=\"{}\"/>",
            templates::NAME,
            templates::anchor(index),
            pad.name
        );
        assert!(sheet.contains(&link), "{}: expected {link}", pad.name);
    }
}

#[test]
fn the_bridge_pad_names_the_rows_a_waterfall_reads() {
    let mut archive = archive();
    let strings = shared_strings(&mut archive);
    let sheet = sheet_xml(&mut archive, templates::NAME);
    let index = templates::PADS.iter().position(|pad| pad.id == "ebitda-bridge").expect("bridge pad");
    let (first_row, first_col, last_row, last_col) = templates::pad_rect(index);
    let hint = cell_text(&sheet, &cell(templates::label_row(index), tl::ANCHOR_COL + 1), &strings);
    // The block opens with a title row (its second cell is empty, which the chart refuses) and ends
    // with a check row: the reader selects what lies between.
    let rows = format!("{}:{}", cell(first_row + 1, first_col), cell(last_row - 1, last_col));
    assert!(hint.contains(&format!("select {rows} (Opening to Closing EBITDA)")), "{hint:?} should send the reader to {rows}");
    assert!(!hint.contains(['{', '}']), "{hint:?} leaves a placeholder");
}

#[test]
fn a_formatted_blank_cell_ends_at_its_own_slash() {
    let sheet = r#"<row r="1"><c r="A1" s="2"/><c r="B1" t="s"><v>3</v></c><c r="C1" s="4"/></row>"#;
    assert_eq!(cell_xml(sheet, "A1").as_deref(), Some(r#"<c r="A1" s="2"/>"#), "a blank must not run into B1");
    assert_eq!(cell_xml(sheet, "B1").as_deref(), Some(r#"<c r="B1" t="s"><v>3</v></c>"#));
    assert_eq!(cell_xml(sheet, "C1").as_deref(), Some(r#"<c r="C1" s="4"/>"#), "the last cell of a row");
    assert_eq!(cell_xml(sheet, "D1"), None);
    assert_eq!(values_corner(sheet), (0, 1), "a formatted blank holds no value, so C1 is not the corner");
}

#[test]
fn comps_keeps_the_statistic_rows_and_the_helper_columns_empty() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, comps::NAME);
    // The blank row under the table and the six rows Comps stats writes into.
    assert!(!any_cell_in(&sheet, cl::LAST_PEER + 1, 0, cl::STATS_FIRST + cl::STATS_ROWS - 1, cl::LAST_COL), "stats rows");
    // The Method | Low | Range block the Football field writes beside its three columns.
    let helper = (cl::FOOT_HEADER, cl::FOOT_COLS, cl::FOOT_FIRST + cl::FOOT_COUNT - 1, cl::FOOT_COLS + 2);
    assert!(!any_cell_in(&sheet, helper.0, helper.1, helper.2, helper.3), "helper block columns");
}

// Comps stats need two data rows, and the VAT input has to sit above the paste source, outside it.
const _: () = assert!(cl::PEER_COUNT >= 2, "Comps stats need two data rows");
const _: () = assert!(pl::VAT_ROW < pl::SRC_HEADER, "the VAT input sits above the source block");

#[test]
fn comps_peers_are_a_plausible_table() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, comps::NAME);
    let header = format!("A{}", cl::HEADER + 1);
    assert!(cell_xml(&sheet, &header).expect("header").contains("t=\"s\""), "the label column's title is text");
    for (company, revenue, ebitda, ev) in comps::PEERS {
        let multiple = ev / ebitda;
        assert!((5.0..=9.5).contains(&multiple), "{company}: EV / EBITDA {multiple:.1}x");
        assert!(ebitda < revenue && ebitda > 0.0, "{company}: EBITDA between zero and revenue");
    }
    for (method, low, high) in comps::METHODS {
        assert!(0.0 < low && low < high, "{method}: low {low} high {high}");
    }
    let first = cl::FIRST_PEER + 1;
    assert_eq!(cell_formula(&sheet, &format!("F{first}")), format!("D{first}/C{first}"));
}

#[test]
fn paste_pads_are_empty_and_the_plain_block_has_no_format() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, paste_lab::NAME);
    for k in 0..pl::TARGETS as usize {
        let (first, last) = paste_lab::target_rows(k);
        assert!(!any_cell_in(&sheet, first, 0, last, pl::COLS - 1), "pad {k} holds a cell");
    }
    let raw_header = pl::RAW_LABEL + 1;
    for (row, col) in cells(&sheet) {
        if (raw_header..raw_header + pl::BLOCK_ROWS).contains(&row) {
            let address = format!("{}{}", char::from(b'A' + col as u8), row + 1);
            let xml = cell_xml(&sheet, &address).expect("cell");
            assert!(!xml.contains(" s=\""), "{address} carries a format: {xml}");
        }
    }
}

#[test]
fn the_paste_source_reads_inside_the_block_and_one_input_outside_it() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, paste_lab::NAME);
    let first = pl::SRC_FIRST + 1;
    assert_eq!(cell_formula(&sheet, &format!("D{first}")), format!("B{first}*C{first}"));
    assert_eq!(cell_formula(&sheet, &format!("E{first}")), format!("D{first}*$B${}", pl::VAT_ROW + 1));
    assert_eq!(cell_number(&sheet, &paste_lab::vat_address()), 0.21);
    assert!(row_tag(&sheet, pl::SRC_HEADER + 1).contains("ht=\"30\""), "the header row is 30 pt tall");
    assert!(!row_tag(&sheet, pl::RAW_LABEL + 2).contains("ht="), "the plain block keeps default heights");
}

/// The opening `<row ...>` tag of a 1-based sheet row.
fn row_tag(sheet: &str, row: u32) -> String {
    let start = sheet.find(&format!("<row r=\"{row}\"")).unwrap_or_else(|| panic!("row {row} was never written"));
    sheet[start..start + sheet[start..].find('>').expect("tag end")].to_string()
}

#[test]
fn the_format_lab_table_is_unformatted_and_the_samples_are_not() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, format_lab::NAME);
    for (row, col) in cells(&sheet) {
        if (fl::BLOCK_TITLE..=fl::TOTAL).contains(&row) && col <= 5 {
            let address = format!("{}{}", char::from(b'A' + col as u8), row + 1);
            assert!(!cell_xml(&sheet, &address).expect("cell").contains(" s=\""), "{address} is formatted");
        }
    }
    let (slot1, slot2) = format_lab::sample_addresses();
    for address in [slot1, slot2] {
        assert!(cell_xml(&sheet, &address).expect("sample").contains(" s=\""), "{address} should be styled");
    }
    // The maturities are date serial numbers: 31.12.2027 and 30.06.2026.
    let first = fl::FIRST_ITEM + 1;
    assert_eq!(cell_number(&sheet, &format!("D{first}")), 46_752.0);
    assert_eq!(cell_number(&sheet, &format!("D{}", first + 1)), 46_203.0);
}

/// The most characters one formula of the shape `="text"&ROUND(..)&"text"` can show: literals
/// count as they are, every other piece as a 12-character figure.
fn longest_text(formula: &str) -> usize {
    const FIGURE: usize = 12;
    let (mut quoted, mut depth, mut piece, mut total) = (false, 0, String::new(), 0);
    let close = |piece: &mut String, total: &mut usize| {
        let text = piece.trim();
        *total += if text.starts_with('"') { text.len() - 2 } else { FIGURE };
        piece.clear();
    };
    for c in formula.trim_start_matches('=').chars() {
        match c {
            '"' => quoted = !quoted,
            '(' if !quoted => depth += 1,
            ')' if !quoted => depth -= 1,
            '&' if !quoted && depth == 0 => {
                close(&mut piece, &mut total);
                continue;
            }
            _ => {}
        }
        piece.push(c);
    }
    close(&mut piece, &mut total);
    total
}

#[test]
fn every_headline_fits_a_text_link() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, headlines::NAME);
    let mut longest = 0;
    for (i, (label, formula)) in headlines::sentences().iter().enumerate() {
        let address = headlines::text_cell(i as u32);
        assert_eq!(&cell_formula(&sheet, &address), formula.trim_start_matches('='), "{label}");
        let length = longest_text(formula);
        assert!(length <= headlines::TEXT_MAX_CHARS, "{label}: up to {length} characters");
        longest = longest.max(length);
    }
    assert!(longest > 200, "the paragraph is the long text a text link has to carry, was {longest}");
    assert_eq!(headlines::sentences().len() as u32, hl::COUNT);
}

#[test]
fn review_has_its_error_its_volatiles_and_its_stray_cell() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, review::NAME);
    let error = review::error_address();
    let row = &error[1..];
    assert_eq!(cell_formula(&sheet, &error), format!("B{row}/C{row}"));
    assert_eq!(cell_number(&sheet, &format!("C{row}")), 0.0, "no units: the division fails");
    let flags = rl::FLAGS_FIRST + 1;
    assert_eq!(cell_formula(&sheet, &format!("B{flags}")), "TODAY()");
    assert!(cell_formula(&sheet, &format!("B{}", flags + 1)).contains("OFFSET("));
    assert!(cell_formula(&sheet, &format!("B{}", flags + 2)).contains("*1.15"));
    // Fill formula down needs the formula in the first row only.
    assert!(cell_xml(&sheet, &review::margin_first_address()).expect("margin").contains("<f>"));
    for i in 1..rl::PRODUCTS {
        let below = format!("D{}", rl::MARGIN_FIRST + i + 1);
        assert!(cell_xml(&sheet, &below).is_none(), "{below} should be empty");
    }
}

#[test]
fn review_ends_in_one_formatted_blank_and_holds_no_drawing() {
    let mut archive = archive();
    let sheet = sheet_xml(&mut archive, review::NAME);
    let stray = cell_xml(&sheet, &review::stray_address()).expect("the stray cell");
    assert!(stray.contains(" s=\"") && !stray.contains("<v>") && !stray.contains("<f>"), "{stray}");
    let (last_row, last_col) = cells(&sheet).iter().fold((0, 0), |(r, c), (row, col)| (r.max(*row), c.max(*col)));
    assert_eq!((last_row, last_col), (rl::STRAY_ROW, rl::STRAY_COL), "nothing sits beyond the stray cell");
    for part in ["<drawing", "<legacyDrawing", "<tableParts"] {
        assert!(!sheet.contains(part), "a {part} would make Clean past the data only clear formats");
    }
    // Where Ctrl+End lands once the surplus is gone: the corner of what holds a value in the saved
    // file, not the constant the checklist text is built from.
    let (row, col) = values_corner(&sheet);
    assert_eq!(cell(row, col), review::data_corner_address(), "the corner of the data after Clean past the data");
}

#[test]
fn no_sheet_outgrows_the_scan_cap() {
    let mut archive = archive();
    for (name, _) in SHEETS {
        let sheet = sheet_xml(&mut archive, name);
        let at = sheet.find("<dimension ref=\"").expect("dimension") + "<dimension ref=\"".len();
        let reference = &sheet[at..at + sheet[at..].find('"').expect("quote")];
        let corner = reference.split(':').nth(1).unwrap_or(reference);
        let (row, col) = common::parse_address(corner);
        let used = u64::from(row + 1) * (u64::from(col) + 1);
        assert!(used <= SCAN_CAP, "{name}: {reference} is {used} cells, over the {SCAN_CAP} the scans read");
    }
}
