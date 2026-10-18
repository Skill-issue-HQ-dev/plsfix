// Reads the tracked demo deck (deck/pls,fix Demo Deck.pptx, built by deck/build_deck.py) and holds it to
// what the workbook promises: eight slides in order, one landing spot per Excel link kind, the half and
// quarter frames at the pane's exact geometry, the shape names of slides 2 to 4 the video capture deletes,
// boxes that stay on slide 7 and apart after Match size, the Highlighter theme, and every "Slide N" named.

mod common;

use std::io::{Cursor, Read};

use common::{archive, shared_strings};

const DECK: &str = "deck/pls,fix Demo Deck.pptx";
/// One point in EMU.
const POINT: u64 = 12_700;

/// The slide titles in order: what the Slide picker lists.
const TITLES: [&str; 8] = [
    "pls,fix demo deck",
    "Choose the slide",
    "Choose the spot",
    "Into a placeholder",
    "Choose the spot: four quarters",
    "Object tools: align and distribute",
    "Object tools: match size, select similar, swap",
    "Object tools: Smart Painter",
];

fn deck() -> zip::ZipArchive<Cursor<Vec<u8>>> {
    let bytes = std::fs::read(DECK).unwrap_or_else(|error| panic!("{DECK}: {error}"));
    zip::ZipArchive::new(Cursor::new(bytes)).expect("a pptx is a zip")
}

fn part(deck: &mut zip::ZipArchive<Cursor<Vec<u8>>>, name: &str) -> String {
    let mut text = String::new();
    deck.by_name(name).unwrap_or_else(|_| panic!("{name} is missing")).read_to_string(&mut text).expect("utf8");
    text
}

fn slide(deck: &mut zip::ZipArchive<Cursor<Vec<u8>>>, number: usize) -> String {
    part(deck, &format!("ppt/slides/slide{number}.xml"))
}

/// The `<a:off x y/>` and `<a:ext cx cy/>` pairs of a slide in points, in file order.
fn boxes(xml: &str) -> Vec<(u64, u64, u64, u64)> {
    let attr = |tag: &str, name: &str| -> u64 {
        let at = tag.find(&format!("{name}=\"")).expect(name) + name.len() + 2;
        tag[at..at + tag[at..].find('"').expect("quote")].parse().expect("emu")
    };
    let offs: Vec<&str> = xml.split("<a:off ").skip(1).map(|rest| &rest[..rest.find('>').unwrap_or(0)]).collect();
    let exts: Vec<&str> = xml.split("<a:ext ").skip(1).map(|rest| &rest[..rest.find('>').unwrap_or(0)]).collect();
    offs.iter()
        .zip(&exts)
        .map(|(off, ext)| (attr(off, "x") / POINT, attr(off, "y") / POINT, attr(ext, "cx") / POINT, attr(ext, "cy") / POINT))
        .collect()
}

/// Every shape of a slide as (name, (left, top, width, height) in points), in file order.
fn shapes(xml: &str) -> Vec<(String, (u64, u64, u64, u64))> {
    xml.split("<p:sp>")
        .skip(1)
        .filter_map(|shape| {
            let at = shape.find("name=\"")? + "name=\"".len();
            let name = shape[at..at + shape[at..].find('"')?].to_string();
            Some((name, *boxes(shape).first()?))
        })
        .collect()
}

/// The dashed frames of a slide: every shape whose outline is a dash.
fn frames(xml: &str) -> Vec<(u64, u64, u64, u64)> {
    xml.split("<p:sp>")
        .skip(1)
        .filter(|shape| shape.contains("<a:prstDash val=\"dash\"/>"))
        .map(|shape| boxes(shape)[0])
        .collect()
}

#[test]
fn the_deck_has_the_eight_slides_in_order() {
    let mut deck = deck();
    for (index, title) in TITLES.iter().enumerate() {
        let xml = slide(&mut deck, index + 1);
        assert!(xml.contains(&format!("<a:t>{title}</a:t>")), "slide {} should be titled {title:?}", index + 1);
    }
    assert!(deck.by_name("ppt/slides/slide9.xml").is_err(), "a ninth slide appeared; add its title to TITLES");
}

#[test]
fn the_overview_names_every_link_kind_and_its_slide() {
    let mut deck = deck();
    let overview = slide(&mut deck, 1);
    for kind in ["table link", "chart link", "picture link", "text link"] {
        assert!(overview.contains(kind), "slide 1 should name the {kind}");
    }
    for number in 2..=8 {
        assert!(overview.contains(&format!("Slide {number}")), "slide 1 never points at slide {number}");
    }
}

#[test]
fn the_halves_and_quarters_are_the_panes_exact_spots() {
    let mut deck = deck();
    // 960 x 540 pt, a 36 pt margin and a 12 pt gap: src/layout.ts spotBox.
    let halves = vec![(36, 36, 438, 468), (486, 36, 438, 468)];
    assert_eq!(frames(&slide(&mut deck, 3)), halves, "slide 3: the left and right half");
    let quarters = vec![(36, 36, 438, 228), (486, 36, 438, 228), (36, 276, 438, 228), (486, 276, 438, 228)];
    assert_eq!(frames(&slide(&mut deck, 5)), quarters, "slide 5: top left, top right, bottom left, bottom right");
}

#[test]
fn the_landing_spots_are_empty_placeholders() {
    let mut deck = deck();
    // Slide 2: its title is the only text. Slide 4: its title and its caption, and no picture yet.
    for (number, texts) in [(2, 1), (4, 2)] {
        let xml = slide(&mut deck, number);
        assert!(xml.matches("<p:ph ").count() >= 2, "slide {number}: a title and the empty placeholder");
        assert_eq!(xml.matches("<a:t>").count(), texts, "slide {number}: text beyond the title and caption");
        assert!(!xml.contains("<p:pic>"), "slide {number}: the placeholder must stay empty");
    }
}

#[test]
fn the_object_tool_slides_hold_named_boxes() {
    let mut deck = deck();
    let names = |xml: &str| -> Vec<String> {
        xml.split("<p:cNvPr id=\"").skip(1).map(|rest| {
            let at = rest.find("name=\"").expect("name") + 6;
            rest[at..at + rest[at..].find('"').expect("quote")].to_string()
        }).collect()
    };
    let six = names(&slide(&mut deck, 6));
    for name in ["Box A", "Box B", "Box C", "Box D", "Box E"] {
        assert!(six.iter().any(|shape| shape == name), "slide 6 lacks {name}");
    }
    let seven = names(&slide(&mut deck, 7));
    for name in ["Box A", "Box B", "Box C", "Box D", "Box E", "Circle 1", "Circle 2"] {
        assert!(seven.iter().any(|shape| shape == name), "slide 7 lacks {name}");
    }
    let eight = names(&slide(&mut deck, 8));
    for name in ["Source", "Target 1", "Target 2", "Target 3"] {
        assert!(eight.iter().any(|shape| shape == name), "slide 8 lacks {name}");
    }
    // Select similar needs A, B and C to be one kind and size; D and E must differ from them.
    let sizes: Vec<(u64, u64, u64, u64)> = boxes(&slide(&mut deck, 7));
    let (a, b, c, d) = (sizes[2], sizes[3], sizes[4], sizes[5]);
    assert!((a.2, a.3) == (b.2, b.3) && (b.2, b.3) == (c.2, c.3), "A, B and C are the same size");
    assert!((d.2, d.3) != (a.2, a.3), "D differs from A");
}

#[test]
fn slides_two_to_four_keep_the_shape_names_the_video_capture_deletes() {
    // The shape order on slides 2 to 4 is fixed on purpose: python-pptx names a shape after its id,
    // and video/capture/ppt-shots.mjs deletes slide 3's two guides and their captions by these names.
    let expected: [(usize, &[&str]); 3] = [
        (2, &["Title 1", "Content Placeholder 2", "Rectangle 3"]),
        (3, &["Title 1", "Rectangle 2", "Rectangle 3", "TextBox 4", "Rectangle 5", "TextBox 6"]),
        (4, &["Title 1", "Picture Placeholder 2", "Text Placeholder 3", "Rectangle 4"]),
    ];
    let mut deck = deck();
    for (number, names) in expected {
        let found: Vec<String> = shapes(&slide(&mut deck, number)).into_iter().map(|(name, _)| name).collect();
        assert_eq!(found, names, "slide {number}: the shape names and their order");
    }
}

#[test]
fn match_size_keeps_slide_seven_on_the_slide_and_its_boxes_apart() {
    // 960 x 540 pt with a 36 pt margin, the caption line at 430 pt, a 12 pt gap: deck_kit.py.
    const LEFT: u64 = 36;
    const RIGHT: u64 = 924;
    const TOP: u64 = 36;
    const CAPTION_TOP: u64 = 430;
    const GAP: u64 = 12;
    let mut deck = deck();
    let mut placed: Vec<(String, (u64, u64, u64, u64))> = shapes(&slide(&mut deck, 7))
        .into_iter()
        .filter(|(name, _)| name.starts_with("Box ") || name.starts_with("Circle "))
        .collect();
    let size_of = |placed: &[(String, (u64, u64, u64, u64))], name: &str| {
        let (_, (_, _, width, height)) = placed.iter().find(|(shape, _)| shape == name).expect(name);
        (*width, *height)
    };
    // Match size, with D selected first: A and E take D's width and height and keep their own corner.
    let (width, height) = size_of(&placed, "Box D");
    for (name, geometry) in &mut placed {
        if name == "Box A" || name == "Box E" {
            geometry.2 = width;
            geometry.3 = height;
        }
    }
    for (name, (left, top, width, height)) in &placed {
        assert!(*left >= LEFT && left + width <= RIGHT, "{name}: {left}..{} pt leaves the content width", left + width);
        assert!(*top >= TOP && top + height <= CAPTION_TOP, "{name}: {top}..{} pt leaves the room above the caption", top + height);
    }
    let apart = |one: &(u64, u64, u64, u64), other: &(u64, u64, u64, u64)| {
        one.0 + one.2 + GAP <= other.0 || other.0 + other.2 + GAP <= one.0 || one.1 + one.3 + GAP <= other.1 || other.1 + other.3 + GAP <= one.1
    };
    for (index, (name, geometry)) in placed.iter().enumerate() {
        for (other_name, other) in &placed[index + 1..] {
            assert!(apart(geometry, other), "{name} and {other_name} touch or overlap after Match size");
        }
    }
}

#[test]
fn the_deck_wears_the_highlighter_theme() {
    let mut deck = deck();
    let theme = part(&mut deck, "ppt/theme/theme1.xml");
    assert!(theme.contains("<a:clrScheme name=\"Highlighter\">"), "the colour scheme is renamed");
    for office_blue in ["4F81BD", "1F497D", "4472C4", "44546A"] {
        assert!(!theme.contains(office_blue), "{office_blue} is Office blue, left in the theme");
    }
    assert!(theme.contains("6D7E12"), "the olive accent is in the theme");
}

#[test]
fn every_slide_the_workbook_names_exists_in_the_deck() {
    let mut workbook = archive();
    let mut deck = deck();
    let slides = (1..).take_while(|n| deck.by_name(&format!("ppt/slides/slide{n}.xml")).is_ok()).count();
    let mut named = Vec::new();
    for text in shared_strings(&mut workbook) {
        for (at, _) in text.match_indices("Slide ") {
            let digits: String = text[at + 6..].chars().take_while(char::is_ascii_digit).collect();
            if let Ok(number) = digits.parse::<usize>() {
                assert!((1..=slides).contains(&number), "{text:?} names slide {number} of {slides}");
                named.push(number);
            }
        }
    }
    for number in 2..=8 {
        assert!(named.contains(&number), "no sheet of the workbook sends the reader to slide {number}");
    }
}
