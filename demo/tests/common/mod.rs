// What the block tests share: the saved workbook as a zip, one sheet's XML by name, and the cells
// of a sheet read back as coordinates, so a test can ask "is anything in this rectangle?" of the
// file itself rather than of the builders. Each test crate uses a part of it.
#![allow(dead_code)]

use std::io::{Cursor, Read};

use smt_demo::sheets::SHEETS;

pub type Archive = zip::ZipArchive<Cursor<Vec<u8>>>;

pub fn archive() -> Archive {
    let (mut workbook, _) = smt_demo::build().expect("build");
    let bytes = workbook.save_to_buffer().expect("save");
    zip::ZipArchive::new(Cursor::new(bytes)).expect("zip")
}

pub fn entry(archive: &mut Archive, name: &str) -> String {
    let mut text = String::new();
    archive.by_name(name).expect(name).read_to_string(&mut text).expect("utf8");
    text
}

pub fn sheet_xml(archive: &mut Archive, name: &str) -> String {
    let index = SHEETS.iter().position(|(sheet, _)| *sheet == name).expect(name);
    entry(archive, &format!("xl/worksheets/sheet{}.xml", index + 1))
}

/// Excel letters to a 0-based column: A -> 0, AD -> 29.
fn column_index(letters: &str) -> u16 {
    letters.bytes().fold(0u32, |n, b| n * 26 + u32::from(b - b'A' + 1)) as u16 - 1
}

/// "AD800" to (799, 29), both 0-based.
pub fn parse_address(address: &str) -> (u32, u16) {
    let split = address.find(|c: char| c.is_ascii_digit()).expect("digits in an address");
    let row: u32 = address[split..].parse().expect("row");
    (row - 1, column_index(&address[..split]))
}

/// Every `<c r="..">` of a sheet as (row, column), 0-based, in file order.
pub fn cells(sheet: &str) -> Vec<(u32, u16)> {
    sheet
        .split("<c r=\"")
        .skip(1)
        .map(|rest| parse_address(&rest[..rest.find('"').expect("closing quote")]))
        .collect()
}

/// True when any cell of the sheet sits inside the rectangle (inclusive, 0-based).
pub fn any_cell_in(sheet: &str, first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> bool {
    cells(sheet)
        .iter()
        .any(|(row, col)| (first_row..=last_row).contains(row) && (first_col..=last_col).contains(col))
}

/// How long the cell element is that `rest` opens (anywhere inside its opening tag): up to its own
/// `/>` when the tag closes itself (a formatted blank), else up to its `</c>`. Never into the next cell.
fn cell_len(rest: &str) -> usize {
    let open = rest.find('>').expect("the end of the cell's opening tag");
    if rest[..open].ends_with('/') {
        return open + 1;
    }
    open + rest[open..].find("</c>").expect("the cell's </c>") + "</c>".len()
}

/// The `<c ...>...</c>` (or self-closing) element of one cell, or None when it was never written.
pub fn cell_xml(sheet: &str, address: &str) -> Option<String> {
    let start = sheet.find(&format!("<c r=\"{address}\""))?;
    let rest = &sheet[start..];
    Some(rest[..cell_len(rest)].to_string())
}

/// The last row and the last column (0-based, each the largest over the cells that hold a value or
/// a formula) of a sheet: where Ctrl+End lands once formatted blanks are gone. The corner cell
/// itself need not hold anything.
pub fn values_corner(sheet: &str) -> (u32, u16) {
    let filled = |cell: &str| ["<v>", "<f>", "<f ", "<is>"].iter().any(|part| cell.contains(part));
    sheet
        .split("<c r=\"")
        .skip(1)
        .filter(|rest| filled(&rest[..cell_len(rest)]))
        .map(|rest| parse_address(&rest[..rest.find('"').expect("closing quote")]))
        .fold((0, 0), |(row, col), (r, c)| (row.max(r), col.max(c)))
}

/// The workbook's shared strings in index order, XML entities undone.
pub fn shared_strings(archive: &mut Archive) -> Vec<String> {
    let xml = entry(archive, "xl/sharedStrings.xml");
    let mut strings = Vec::new();
    let mut rest = xml.as_str();
    while let Some(start) = rest.find("<t") {
        let tag_end = rest[start..].find('>').expect("tag close") + start;
        let text_start = tag_end + 1;
        let text_end = rest[text_start..].find("</t>").expect("text end") + text_start;
        strings.push(unescape(&rest[text_start..text_end]));
        rest = &rest[text_end + "</t>".len()..];
    }
    strings
}

pub fn unescape(text: &str) -> String {
    text.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", "\"").replace("&amp;", "&")
}

/// The text of a shared-string cell, resolved through `strings`.
pub fn cell_text(sheet: &str, address: &str, strings: &[String]) -> String {
    let xml = cell_xml(sheet, address).unwrap_or_else(|| panic!("{address}: no such cell"));
    let start = xml.find("<v>").unwrap_or_else(|| panic!("{address}: not a string cell: {xml}")) + "<v>".len();
    let end = xml[start..].find("</v>").expect("value end") + start;
    strings[xml[start..end].parse::<usize>().expect("shared string index")].clone()
}

/// The formula text of a cell, entities undone.
pub fn cell_formula(sheet: &str, address: &str) -> String {
    let xml = cell_xml(sheet, address).unwrap_or_else(|| panic!("{address}: no such cell"));
    let start = xml.find("<f>").unwrap_or_else(|| panic!("{address}: not a formula: {xml}")) + "<f>".len();
    let end = xml[start..].find("</f>").expect("formula end") + start;
    unescape(&xml[start..end])
}

/// The numeric value stored in a cell.
pub fn cell_number(sheet: &str, address: &str) -> f64 {
    let xml = cell_xml(sheet, address).unwrap_or_else(|| panic!("{address}: no such cell"));
    let start = xml.find("<v>").unwrap_or_else(|| panic!("{address}: no value: {xml}")) + "<v>".len();
    let end = xml[start..].find("</v>").expect("value end") + start;
    xml[start..end].parse().expect("number")
}

/// Every string a sheet shows, in cell order: what a person could read on it.
pub fn sheet_text(archive: &mut Archive, name: &str) -> String {
    let strings = shared_strings(archive);
    let sheet = sheet_xml(archive, name);
    let mut text = String::new();
    for rest in sheet.split("<c r=\"").skip(1) {
        let cell = &rest[..cell_len(rest)];
        if cell.contains("t=\"s\"") {
            let start = cell.find("<v>").expect("value") + 3;
            let end = cell[start..].find("</v>").expect("value end") + start;
            text.push_str(&strings[cell[start..end].parse::<usize>().expect("index")]);
            text.push('\n');
        }
    }
    text
}
