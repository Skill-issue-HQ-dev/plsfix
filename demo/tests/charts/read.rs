// Reads the saved Charts sheet back out of the xlsx: the sheet's XML, each chart's object name,
// its XML and the cells its box covers, and every cell the sheet writes. A small hand-rolled
// walk, because the demo has no XML crate and the files are plain.

use std::io::{Cursor, Read};

use smt_demo::sheets::charts::NAME;
use smt_demo::sheets::SHEETS;

/// One chart as saved: its object name, its XML and the cells its box covers (0-based, inclusive).
pub struct Placed {
    pub name: String,
    pub xml: String,
    pub first: (u16, u32),
    pub last: (u16, u32),
}

type Archive = zip::ZipArchive<Cursor<Vec<u8>>>;

fn entry(archive: &mut Archive, name: &str) -> String {
    let mut text = String::new();
    let mut file = archive
        .by_name(name)
        .unwrap_or_else(|_| panic!("{name} missing"));
    file.read_to_string(&mut text).expect("utf8");
    text
}

/// What sits between each `open` and the next `close` (a small stand-in for an XML walk).
pub fn between<'a>(xml: &'a str, open: &str, close: &str) -> Vec<&'a str> {
    let end = |rest: &'a str| &rest[..rest.find(close).expect(close)];
    xml.split(open).skip(1).map(end).collect()
}

/// Every `<a:srgbClr val="..."/>` colour in `xml`, in document order.
pub fn colours(xml: &str) -> Vec<&str> {
    between(xml, "<a:srgbClr val=\"", "\"")
}

/// The value of the first ` name="..."` in `fragment`.
fn value<'a>(fragment: &'a str, name: &str) -> &'a str {
    let at = fragment
        .find(&format!(" {name}=\""))
        .unwrap_or_else(|| panic!("{name} missing"));
    let start = at + name.len() + 3;
    &fragment[start..start + fragment[start..].find('"').expect("closing quote")]
}

/// A drawing anchor's corner as (column, row); a box that ends exactly on a cell edge does not
/// cover the cell it ends on.
fn corner(anchor: &str, side: &str) -> (u16, u32) {
    let part = between(anchor, &format!("<xdr:{side}>"), &format!("</xdr:{side}>"))[0];
    let number = |tag: &str| -> u32 {
        let found = between(part, &format!("<xdr:{tag}>"), &format!("</xdr:{tag}>"));
        found[0].parse().expect("number")
    };
    let end = |index: u32, offset: u32| match (side, offset) {
        ("to", 0) => index - 1,
        _ => index,
    };
    (
        end(number("col"), number("colOff")) as u16,
        end(number("row"), number("rowOff")),
    )
}

/// The Charts sheet's XML and its charts in drawing order, read through its drawing part.
pub fn charts_sheet() -> (String, Vec<Placed>) {
    let bytes = smt_demo::build().expect("build").0.save_to_buffer();
    let archive = &mut zip::ZipArchive::new(Cursor::new(bytes.expect("save"))).expect("zip");
    let at = SHEETS
        .iter()
        .position(|(name, _)| *name == NAME)
        .expect("registered")
        + 1;
    let sheet = entry(archive, &format!("xl/worksheets/sheet{at}.xml"));
    let rels = entry(archive, &format!("xl/worksheets/_rels/sheet{at}.xml.rels"));
    let file = |target: &str| target.rsplit('/').next().expect("file name").to_string();
    let drawing = file(value(&rels, "Target"));
    let drawing_rels = entry(archive, &format!("xl/drawings/_rels/{drawing}.rels"));
    let anchors = entry(archive, &format!("xl/drawings/{drawing}"));
    let mut placed = Vec::new();
    for anchor in between(&anchors, "<xdr:twoCellAnchor", "</xdr:twoCellAnchor>") {
        let id = format!("Id=\"{}\"", value(anchor, "r:id"));
        let rel = drawing_rels
            .split("<Relationship ")
            .find(|part| part.contains(&id));
        let part = file(value(rel.expect("chart relationship"), "Target"));
        placed.push(Placed {
            name: value(anchor, "name").to_string(),
            xml: entry(archive, &format!("xl/charts/{part}")),
            first: corner(anchor, "from"),
            last: corner(anchor, "to"),
        });
    }
    (sheet, placed)
}

/// Every cell the sheet writes, as (column, row), 0-based.
pub fn written(sheet: &str) -> Vec<(u16, u32)> {
    let address = |text: &str| {
        let letters = text.chars().take_while(char::is_ascii_alphabetic).count();
        let col = text[..letters]
            .bytes()
            .fold(0u16, |sum, b| sum * 26 + u16::from(b - b'A' + 1));
        (col - 1, text[letters..].parse::<u32>().expect("row") - 1)
    };
    between(sheet, "<c r=\"", "\"")
        .into_iter()
        .map(address)
        .collect()
}

/// The chart's own title text (the first title in its XML, before any axis title).
pub fn title_of(xml: &str) -> &str {
    let title = &xml[xml.find("<c:title>").expect("chart title")..];
    between(title, "<a:t>", "</a:t>")[0]
}
