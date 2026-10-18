// Reads the saved Charts sheet back: twelve charts with the listed names and titles, the plot
// element of each kind, the series ranges, the gap, the markers, the second axis, the labels, the
// series colours and the label lettering, and no chart over a written cell or over another chart.

mod read;

use std::collections::HashSet;

use read::{between, charts_sheet, colours, title_of, written};
use rust_xlsxwriter::Color;
use smt_demo::layout::{cell_abs, LABEL_COL};
use smt_demo::sheets::charts::{self, NAME};
use smt_demo::sheets::charts_data::{Kind, BLOCKS, GAP_INDEX};
use smt_demo::style::{ACCENT, INK, SERIES};

/// The chart names the tours pick, in sheet order.
const NAMES: [&str; 12] = [
    "Column chart",
    "Stacked column chart",
    "100% stacked column chart",
    "Bar chart",
    "Pie chart",
    "Doughnut chart",
    "Line chart",
    "Area chart",
    "Stacked area chart",
    "Scatter chart",
    "Combo chart",
    "Column chart with a gap",
];

/// The chart titles, in sheet order: readable, unique, short enough for one line.
const TITLES: [&str; 12] = [
    "Revenue, gross profit and EBITDA 2025E (EUR thousands)",
    "Revenue split into costs and EBITDA (EUR thousands)",
    "Revenue mix by market (share of revenue)",
    "Operating costs by category (EUR thousands)",
    "Revenue by market, 2025E (EUR thousands)",
    "Revenue by product category, 2025E (EUR thousands)",
    "Sales by product line 2025E (EUR thousands)",
    "Net debt at month end, 2025E (EUR thousands)",
    "Working capital at month end, 2025E (EUR thousands)",
    "Customer revenue against gross margin 2025E",
    "Monthly revenue and EBITDA margin, 2025E",
    "Export sales by month, 2024A (September not reported)",
];

/// The pane refuses a chart title past this many characters (`CHART_TITLE_MAX`).
const TITLE_LIMIT: usize = 80;

/// Every plot element a chart can hold: a chart has one, a combo two.
const PLOTS: [&str; 6] = [
    "<c:barChart>",
    "<c:pieChart>",
    "<c:doughnutChart>",
    "<c:lineChart>",
    "<c:areaChart>",
    "<c:scatterChart>",
];

fn block_of(kind: Kind) -> usize {
    BLOCKS
        .iter()
        .position(|block| block.kind == kind)
        .expect("block")
}

fn count(xml: &str, needle: &str) -> usize {
    xml.matches(needle).count()
}

#[test]
fn charts_sheet_holds_twelve_charts_with_the_listed_names() {
    let charts = charts_sheet().1;
    let names: Vec<&str> = charts.iter().map(|chart| chart.name.as_str()).collect();
    assert_eq!(names, NAMES);
    assert_eq!(
        BLOCKS.map(|block| block.name),
        NAMES,
        "data module and list drifted"
    );
    assert_eq!(
        names.iter().collect::<HashSet<_>>().len(),
        12,
        "names must be unique"
    );
}

#[test]
fn charts_titles_are_unique_and_listed() {
    let charts = charts_sheet().1;
    let titles: Vec<&str> = charts.iter().map(|chart| title_of(&chart.xml)).collect();
    assert_eq!(titles, TITLES);
    assert_eq!(
        BLOCKS.map(|block| block.title),
        TITLES,
        "data module and list drifted"
    );
    assert_eq!(
        titles.iter().collect::<HashSet<_>>().len(),
        12,
        "titles must be unique"
    );
    for title in titles {
        assert!(
            !title.is_empty() && title.chars().count() <= TITLE_LIMIT,
            "{title:?}"
        );
    }
}

/// What a kind must show in its XML, beyond having exactly its own plot element.
fn needles(kind: Kind) -> &'static [&'static str] {
    const COLUMNS: &str = "<c:barDir val=\"col\"/>";
    match kind {
        Kind::Column | Kind::ColumnWithGap => &[COLUMNS, "val=\"clustered\"/>"],
        Kind::StackedColumn => &[COLUMNS, "val=\"stacked\"/>", "<c:overlap val=\"100\"/>"],
        Kind::PercentStackedColumn => &[COLUMNS, "val=\"percentStacked\"/>"],
        Kind::Bar => &["<c:barDir val=\"bar\"/>", "val=\"clustered\"/>"],
        Kind::Pie => &["<c:pieChart>"],
        Kind::Doughnut => &["<c:doughnutChart>"],
        Kind::Line => &["<c:lineChart>", "val=\"standard\"/>"],
        Kind::Area => &["<c:areaChart>", "val=\"standard\"/>"],
        Kind::StackedArea => &["<c:areaChart>", "val=\"stacked\"/>"],
        Kind::Scatter => &["<c:scatterChart>"],
        Kind::Combo => &["<c:barChart>", "<c:lineChart>"],
    }
}

#[test]
fn charts_use_the_plot_element_of_their_kind() {
    for (block, chart) in BLOCKS.iter().zip(charts_sheet().1) {
        let name = block.name;
        for needle in needles(block.kind) {
            assert!(chart.xml.contains(needle), "{name}: {needle} missing");
        }
        let used = PLOTS
            .iter()
            .filter(|plot| chart.xml.contains(**plot))
            .count();
        assert_eq!(
            used,
            if block.kind == Kind::Combo { 2 } else { 1 },
            "{name}: plot elements"
        );
    }
}

#[test]
fn charts_series_read_the_block_beside_them() {
    for (index, (block, chart)) in BLOCKS.iter().zip(charts_sheet().1).enumerate() {
        let (at, xml, name) = (charts::place(index), &chart.xml, block.name);
        let range = |col| {
            format!(
                "{NAME}!{}:{}",
                cell_abs(at.first, col),
                cell_abs(at.last, col)
            )
        };
        let values = |series: usize| range(charts::FIRST_SERIES_COL + series as u16);
        if block.kind == Kind::Scatter {
            assert!(
                xml.contains(&format!("<c:xVal><c:numRef><c:f>{}", values(0))),
                "X values"
            );
            assert!(
                xml.contains(&format!("<c:yVal><c:numRef><c:f>{}", values(1))),
                "Y values"
            );
            continue;
        }
        assert!(xml.contains(&range(LABEL_COL)), "{name}: labels");
        for series in 0..block.series.len() {
            let head = cell_abs(at.header, charts::FIRST_SERIES_COL + series as u16);
            assert!(xml.contains(&values(series)), "{name}: series {series}");
            assert!(
                xml.contains(&format!("{NAME}!{head}")),
                "{name}: series {series} name"
            );
        }
    }
}

#[test]
fn charts_stay_off_the_data_and_each_other() {
    let (sheet, charts) = charts_sheet();
    let cells = written(&sheet);
    assert!(cells.len() > 400, "the data blocks are missing");
    for (index, chart) in charts.iter().enumerate() {
        let (first, last, name) = (chart.first, chart.last, &chart.name);
        let inside = |(col, row): &(u16, u32)| {
            (first.0..=last.0).contains(col) && (first.1..=last.1).contains(row)
        };
        assert!(!cells.iter().any(inside), "{name} covers a written cell");
        assert_eq!(
            first.1,
            charts::band_row(index),
            "{name}: not on its heading row"
        );
        let last_data_col = BLOCKS[index].series.len() as u16 + charts::FIRST_SERIES_COL - 1;
        assert!(first.0 > last_data_col, "{name}: not beside its block");
        for other in &charts[index + 1..] {
            let apart = last.1 < other.first.1 || last.0 < other.first.0;
            assert!(apart, "{name} touches {}", other.name);
        }
    }
}

#[test]
fn charts_gap_is_one_empty_cell_shown_as_a_gap() {
    let (sheet, charts) = charts_sheet();
    let index = block_of(Kind::ColumnWithGap);
    let at = charts::place(index);
    let filled: Vec<u32> = written(&sheet)
        .into_iter()
        .filter(|(col, row)| *col == charts::FIRST_SERIES_COL && (at.first..=at.last).contains(row))
        .map(|(_, row)| row)
        .collect();
    assert_eq!(
        filled.len(),
        11,
        "one of the twelve month cells must stay empty"
    );
    assert!(
        !filled.contains(&(at.first + GAP_INDEX as u32)),
        "September is the empty one"
    );
    let xml = &charts[index].xml;
    assert!(
        xml.contains("<c:dispBlanksAs val=\"gap\"/>"),
        "an empty cell must read as a gap"
    );
    let cache = between(xml, "<c:val>", "</c:val>")[0];
    assert!(cache.contains("<c:ptCount val=\"12\"/>"), "{cache}");
    assert_eq!(count(cache, "<c:pt "), 11, "the gap has no cached point");
    assert!(
        !cache.contains(&format!("<c:pt idx=\"{GAP_INDEX}\">")),
        "the gap has a point"
    );
}

#[test]
fn charts_line_has_markers_on_one_series_only() {
    let xml = &charts_sheet().1[block_of(Kind::Line)].xml;
    let (circle, none) = ("<c:symbol val=\"circle\"/>", "<c:symbol val=\"none\"/>");
    assert_eq!(count(xml, "<c:ser>"), 3);
    assert_eq!((count(xml, circle), count(xml, none)), (1, 2));
    assert!(
        xml.find(circle) < xml.find(none),
        "the first series carries the markers"
    );
}

#[test]
fn charts_combo_puts_the_margin_line_on_a_second_axis() {
    let xml = &charts_sheet().1[block_of(Kind::Combo)].xml;
    let axes = |plot: &str| -> Vec<String> {
        let part = between(xml, &format!("<c:{plot}>"), &format!("</c:{plot}>"))[0];
        between(part, "<c:axId val=\"", "\"")
            .iter()
            .map(|id| id.to_string())
            .collect()
    };
    let (columns, line) = (axes("barChart"), axes("lineChart"));
    assert_eq!((columns.len(), line.len()), (2, 2), "two axes each");
    assert!(
        columns.iter().all(|id| !line.contains(id)),
        "the line shares the columns' axes"
    );
    assert_eq!(count(xml, "<c:valAx>"), 2, "one value axis per side");
    assert!(
        xml.contains("<c:crosses val=\"max\"/>"),
        "the second value axis sits on the right"
    );
    assert!(
        xml.contains("formatCode=\"0%\""),
        "the margin axis reads in percent"
    );
}

// Every chart of the demo labels its points with the value and nothing else, the rule the add-in
// itself follows; the column chart (all three series) and the pie are the two the brief asks for.
#[test]
fn charts_label_the_values_only() {
    let charts = charts_sheet().1;
    for (block, chart) in BLOCKS.iter().zip(&charts) {
        let xml = &chart.xml;
        assert_eq!(
            xml.contains("<c:dLbls>"),
            xml.contains("<c:showVal val=\"1\"/>")
        );
        for part in ["showPercent", "showCatName", "showSerName", "showLegendKey"] {
            assert!(
                !xml.contains(&format!("<c:{part} val=\"1\"/>")),
                "{}: {part}",
                block.name
            );
        }
    }
    assert_eq!(count(&charts[block_of(Kind::Column)].xml, "<c:dLbls>"), 3);
    assert_eq!(count(&charts[block_of(Kind::Pie)].xml, "<c:dLbls>"), 1);
}

/// The `RRGGBB` text of an RGB colour, as the chart XML spells it.
fn rgb(color: Color) -> String {
    match color {
        Color::RGB(value) => format!("{value:06X}"),
        other => panic!("{other:?} is not an RGB colour"),
    }
}

/// The distinct colours a chart draws its series with, in the order they first appear. White is
/// the lettering of a label on a dark fill, never a series colour.
fn drawn_with(xml: &str) -> Vec<String> {
    let mut seen: Vec<String> = Vec::new();
    for colour in colours(xml) {
        if colour != "FFFFFF" && !seen.iter().any(|known| known == colour) {
            seen.push(colour.to_string());
        }
    }
    seen
}

// The pane's series order: one colour per series, one per slice for a pie or a ring, the first
// for a chart with a single series, and the ink for the combo's margin line on its own.
#[test]
fn charts_paint_the_series_colours_in_order() {
    for (block, chart) in BLOCKS.iter().zip(charts_sheet().1) {
        let expected: Vec<Color> = match block.kind {
            Kind::Pie | Kind::Doughnut => SERIES[..block.labels.len()].to_vec(),
            Kind::Area | Kind::Scatter | Kind::ColumnWithGap => vec![ACCENT],
            Kind::Combo => vec![ACCENT, INK],
            _ => SERIES[..block.series.len()].to_vec(),
        };
        let expected: Vec<String> = expected.into_iter().map(rgb).collect();
        assert_eq!(drawn_with(&chart.xml), expected, "{}", block.name);
    }
}

// A label that sits on a segment or a slice is lettered to read on it: white on the olive and the
// ink, ink on the lighter colours. The pie's labels sit outside on white and keep Excel's own.
#[test]
fn charts_letter_stacked_and_ring_labels_to_read_on_their_fill() {
    let charts = charts_sheet().1;
    let lettering = |kind: Kind| -> Vec<String> {
        between(&charts[block_of(kind)].xml, "<c:dLbls>", "</c:dLbls>")
            .into_iter()
            .flat_map(colours)
            .map(String::from)
            .collect()
    };
    assert_eq!(
        lettering(Kind::StackedColumn),
        ["FFFFFF", "FFFFFF", "111111"]
    );
    assert_eq!(
        lettering(Kind::Doughnut),
        ["FFFFFF", "FFFFFF", "111111", "111111", "111111", "111111"]
    );
    assert!(lettering(Kind::Pie).is_empty());
}
