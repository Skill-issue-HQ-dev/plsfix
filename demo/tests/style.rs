// Pins the colours the demo paints with to the pane's Highlighter look (tokens.css, the defaults of
// src/settings.ts, `seriesColor` in src/chart-colors.ts), checks the lettering rule that keeps
// chart labels readable, and reads the saved workbook back for any colour off the palette.

use std::collections::HashSet;
use std::io::{Cursor, Read};

use rust_xlsxwriter::Color;
use smt_demo::style::{text_on, ACCENT, ACCENT_TEXT, GUIDE_TINT, INK, SERIES};

/// `seriesColor(0..12)` at the shipped defaults (primary #111111, accent #6D7E12), worked out
/// with the pane's own `tint` and `shade`.
const PANE_SERIES: [&str; 12] = [
    "6D7E12", "111111", "949494", "AFB87D", "CBCBCB", "D3D8B8", "4C580D", "0C0C0C", "686868",
    "7B8158", "8E8E8E", "949781",
];
/// WCAG's AAA bar for words (7:1) and its AA bar for lettering on a fill (4.5:1).
const WORDS: f64 = 7.0;
const LETTERING: f64 = 4.5;
/// The fills a label sits on: the stacked columns use the first three series, the ring six.
const LABELLED_FILLS: usize = 6;

fn hex(color: Color) -> String {
    match color {
        Color::RGB(rgb) => format!("{rgb:06X}"),
        Color::White => "FFFFFF".to_string(),
        other => panic!("{other:?} is not an RGB colour"),
    }
}

/// WCAG relative luminance of an RGB colour.
fn luminance(color: Color) -> f64 {
    let hex = hex(color);
    let part = |at: usize| {
        let c = f64::from(u8::from_str_radix(&hex[at..at + 2], 16).expect("hex")) / 255.0;
        if c <= 0.03928 {
            c / 12.92
        } else {
            ((c + 0.055) / 1.055).powf(2.4)
        }
    };
    0.2126 * part(0) + 0.7152 * part(2) + 0.0722 * part(4)
}

fn contrast(a: Color, b: Color) -> f64 {
    let (a, b) = (luminance(a), luminance(b));
    (a.max(b) + 0.05) / (a.min(b) + 0.05)
}

#[test]
fn style_names_the_highlighter_palette() {
    assert_eq!(
        [INK, ACCENT, ACCENT_TEXT, GUIDE_TINT].map(hex),
        ["111111", "6D7E12", "4D5B00", "F7FCDB"]
    );
    assert_eq!(SERIES.map(hex), PANE_SERIES);
}

#[test]
fn style_words_read_on_the_grounds_they_sit_on() {
    assert!(contrast(INK, Color::White) >= WORDS, "titles");
    assert!(
        contrast(ACCENT_TEXT, Color::White) >= WORDS,
        "the sheet notes"
    );
    assert!(contrast(INK, GUIDE_TINT) >= WORDS, "the guide band");
}

#[test]
fn style_lettering_reads_on_every_fill_a_label_sits_on() {
    let fills = &SERIES[..LABELLED_FILLS];
    for fill in fills {
        let ratio = contrast(text_on(*fill), *fill);
        assert!(ratio >= LETTERING, "{}: {ratio:.2}:1", hex(*fill));
    }
    let white: Vec<bool> = fills
        .iter()
        .map(|fill| text_on(*fill) == Color::White)
        .collect();
    assert_eq!(white, [true, true, false, false, false, false]);
}

/// Every colour a saved part names: `srgbClr val="RRGGBB"` in a chart, `rgb="FFRRGGBB"` in styles.
fn colours_named(xml: &str) -> Vec<String> {
    let chart = xml
        .split("<a:srgbClr val=\"")
        .skip(1)
        .map(|rest| &rest[..6]);
    let style = xml.split(" rgb=\"FF").skip(1).map(|rest| &rest[..6]);
    chart.chain(style).map(str::to_uppercase).collect()
}

#[test]
fn style_saved_workbook_names_only_palette_colours() {
    let bytes = smt_demo::build()
        .expect("build")
        .0
        .save_to_buffer()
        .expect("save");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let palette: HashSet<String> = [INK, ACCENT, ACCENT_TEXT, GUIDE_TINT, Color::White]
        .into_iter()
        .chain(SERIES)
        .map(hex)
        .collect();
    let mut named = HashSet::new();
    for index in 0..archive.len() {
        let mut part = archive.by_index(index).expect("part");
        let path = part.name().to_string();
        if path != "xl/styles.xml" && !path.starts_with("xl/charts/chart") {
            continue;
        }
        let mut xml = String::new();
        part.read_to_string(&mut xml).expect("utf8");
        for colour in colours_named(&xml) {
            assert!(
                palette.contains(&colour),
                "{path}: {colour} is not on the palette"
            );
            named.insert(colour);
        }
    }
    for colour in [INK, ACCENT, ACCENT_TEXT, GUIDE_TINT] {
        assert!(
            named.contains(&hex(colour)),
            "{} is never painted",
            hex(colour)
        );
    }
}
