// Holds the demo to the panes: every button and checkbox of taskpane.html and pptpane.html is named by a
// "Start here" row (by its own label, not a tab name other rows repeat) or exempted with its reason, so a
// new control fails until the demo has a row for it. Also: every sheet is mentioned, every step is
// complete, and no shipped string has an em dash or a placeholder.

mod common;

use common::{archive, shared_strings, sheet_text};
use smt_demo::sheets::start::{items, Item, Step};
use smt_demo::sheets::SHEETS;

/// (control key, phrase the checklist must contain). The key is the control's data-action, else its id.
const EXCEL: &[(&str, &str)] = &[
    ("reconcile-find", "Find a combination"),
    ("style-title", "Cell styles"),
    ("style-header", "Cell styles"),
    ("style-input", "Cell styles"),
    ("style-formula", "Cell styles"),
    ("style-result", "Cell styles"),
    ("clear-formats", "Cell styles"),
    ("number-whole", "Number formats"),
    ("number-decimal", "Number formats"),
    ("number-currency", "Number formats"),
    ("number-percent", "Number formats"),
    ("style-save", "Save style"),
    ("styles-starter", "Starter styles from palette"),
    ("paint-capture-1", "Format painter"),
    ("paint-capture-2", "Format painter"),
    ("paint-capture-3", "Format painter"),
    ("paint-apply-1", "Format painter"),
    ("paint-apply-2", "Format painter"),
    ("paint-apply-3", "Format painter"),
    ("cycle-number-general", "Format cycles"),
    ("cycle-number-date", "Format cycles"),
    ("cycle-number-currency", "Format cycles"),
    ("cycle-number-percent", "Format cycles"),
    ("cycle-number-multiple", "Format cycles"),
    ("cycle-row-title", "Row styles"),
    ("cycle-row-result", "Row styles"),
    ("cycle-row-item", "Row styles"),
    ("cycle-fill", "Fill, Font, Border"),
    ("cycle-font", "Fill, Font, Border"),
    ("cycle-border", "Fill, Font, Border"),
    ("cycle-row-height", "Row height, Column width"),
    ("cycle-col-width", "Row height, Column width"),
    ("pinstripes-rows", "Pinstripes"),
    ("pinstripes-columns", "Pinstripes"),
    ("cycle-indent", "Indent, Align, Underline"),
    ("cycle-align", "Indent, Align, Underline"),
    ("cycle-underline", "Indent, Align, Underline"),
    ("undo", "Undo last pls,fix action"),
    ("fill-right", "Fill formula right"),
    ("fill-down", "Fill formula down"),
    ("if-error", "IFERROR guard"),
    ("autocolor", "Autocolor selection"),
    ("insert-color-key", "Insert color key"),
    ("write-rounded", "Consistent rounding"),
    ("unpivot", "Unpivot selection"),
    ("comps-stats", "Comps stats"),
    ("copy-source", "Mark, Values, Formats, Transpose"),
    ("paste-values", "Mark, Values, Formats, Transpose"),
    ("paste-formats", "Mark, Values, Formats, Transpose"),
    ("paste-transpose", "Mark, Values, Formats, Transpose"),
    ("paste-exact", "Exact and Paste: duplicate formulas"),
    ("paste-duplicate", "Paste: duplicate formulas"),
    ("paste-number-formats", "Paste: number formats"),
    ("paste-row-heights", "Paste: row heights"),
    ("divide-1000", "x1000, /1000, Sign flip"),
    ("multiply-1000", "x1000, /1000, Sign flip"),
    ("sign-flip", "x1000, /1000, Sign flip"),
    ("dec-more", "decimals"),
    ("dec-less", "decimals"),
    ("cagr", "CAGR"),
    ("audit-toggle", "Audit overlay"),
    ("trace-precedents", "Precedents / Dependents"),
    ("trace-dependents", "Precedents / Dependents"),
    ("select-consistent", "Select consistent region"),
    ("trace-precedents-all", "Precedents of selection"),
    ("chart-waterfall", "Waterfall from selection"),
    ("tornado", "Tornado from selection"),
    ("chart-football", "Football field"),
    ("chart-format", "Brand-format chart"),
    ("chart-cagr", "CAGR label"),
    ("reset-brand", "Reset brand settings"),
    ("export-brand", "Copy palette JSON"),
    ("edit-brand-colors", "Edit colors"),
    ("read-selection-colors", "Read from selection"),
    ("use-logo-colors", "Use these colors"),
    ("shortcuts-card", "Printable card"),
    ("shortcut-card", "Shortcut card"),
    ("shortcuts-apply", "beside an action and press Apply"),
    ("shortcuts-reset", "Reset all (press twice)"),
    ("setting-autocolor-edit", "Autocolor on edit"),
    ("sheets-very-hidden", "Include very hidden"),
    ("find-case", "Match case"),
    ("find-formulas", "Search formulas"),
    ("links-autopush", "Auto-push on edit"),
    ("links-highlight", "Highlight linked cells"),
    ("sheets-unhide-all", "Sheet explorer"),
    ("sheets-show-only", "Sheet explorer"),
    ("sheets-bury", "Sheet explorer"),
    ("sheets-move-end", "Sheet explorer"),
    ("sheets-move-up", "Sheet explorer"),
    ("sheets-move-down", "Sheet explorer"),
    ("find", "Find in workbook"),
    ("styles-scan", "Scan styles"),
    ("insert-toc", "Insert contents sheet"),
    ("share-prepare", "Prepare for sharing"),
    ("clean-past-data", "Clean past the data"),
    ("run-model-check", "Run model check"),
    ("copy-model-check", "Copy report"),
    ("scan-names", "Scan broken names"),
    ("delete-names", "Delete broken names"),
    ("export-selection", "Export selection"),
    ("export-chart", "Export active chart"),
    ("export-text", "Export as text"),
    ("export-picture", "Export as picture"),
    ("copy-for-powerpoint", "Copy for PowerPoint"),
    ("new-project", "New project"),
    ("move-to-project", "Move to project"),
    ("project-ok", "press Create"),
    ("push-selected", "Push selected"),
    ("push-all", "Push all"),
    ("go-to-source", "Go to source"),
    ("convert-to-table", "Convert to table"),
    ("remove-link", "Remove link"),
    ("generate-key", "Generate"),
    ("copy-key", "Copy and Reveal"),
    ("reveal-key", "Reveal the key"),
    ("forget-key", "press Forget (press twice)"),
];

/// Controls with nothing to try on the demo's data, each with why.
const EXCEL_EXEMPT: &[(&str, &str)] = &[
    ("tab-tools", "tab button"),
    ("tab-workbook", "tab button"),
    ("tab-links", "tab button"),
    ("tab-brand", "tab button"),
    ("first-run-shortcuts", "the one-time New here card"),
    ("first-run-dismiss", "the one-time New here card"),
    ("refresh-selection", "refresh icon"),
    ("refresh-sheets", "refresh icon"),
    ("trace-back", "steps the trace panel back, part of Precedents / Dependents"),
    ("styles-delete", "deletes unused styles, and this file carries none"),
    ("project-cancel", "closes the New project box"),
    ("find-comments", "adds threaded comments and replies to the search, and this workbook holds none (Excel 365 only)"),
];

const PPT: &[(&str, &str)] = &[
    ("update-selected", "Update selected"),
    ("update-slide", "Update this slide"),
    ("update-all", "Update all"),
    ("revert-selected", "Revert last update"),
    ("go-to-slide", "Go to slide"),
    ("change-source", "Change source"),
    ("deck-redraw", "Redraw from this deck"),
    ("change-source-confirm", "press Confirm"),
    ("break-selected", "Break link"),
    ("clear-pasted-links", "Clear pasted links"),
    ("paste-latest-linked", "Paste latest linked"),
    ("align-objects", "Align and Distribute"),
    ("distribute-objects", "Align and Distribute"),
    ("match-size", "Match size"),
    ("select-similar", "Select similar"),
    ("swap-objects", "Swap"),
    ("capture-object-style", "Smart Painter"),
    ("apply-object-style", "Smart Painter"),
    ("save-shape-style", "Shape styles: save, apply, update, delete, starters"),
    ("shape-style-starters", "press Starter styles from palette under Shape styles"),
    ("edit-brand-colors", "press Edit colors to open the Settings tab"),
    ("use-workbook-colors", "press Use the workbook's colors"),
    ("use-deck-theme-colors", "press Use this deck's theme colors"),
    ("put-theme-colors", "press Put in the deck theme"),
    ("restore-theme-colors", "press Restore theme colors"),
    ("save-key", "Save key"),
    ("forget-key", "Forget key"),
];

const PPT_EXEMPT: &[(&str, &str)] = &[
    ("tab-links", "tab button"),
    ("tab-inbox", "tab button"),
    ("tab-tools", "tab button"),
    ("tab-settings", "tab button"),
    ("refresh-links", "refresh icon"),
    ("refresh-inbox", "refresh icon"),
    ("change-source-cancel", "closes the Change source picker"),
    ("first-run-ppt-dismiss", "the one-time New here card"),
];

/// The key of every `<button>` and every `<input type="checkbox">`: its data-action, else its id.
fn control_keys(html: &str) -> Vec<String> {
    let attribute = |tag: &str, name: &str| -> Option<String> {
        let at = tag.find(&format!(" {name}=\""))? + name.len() + 3;
        Some(tag[at..at + tag[at..].find('"')?].to_string())
    };
    let tags = |opening: &str| -> Vec<String> {
        html.split(opening).skip(1).filter_map(|rest| Some(rest[..rest.find('>')?].to_string())).collect()
    };
    let checkboxes = tags("<input").into_iter().filter(|tag| attribute(tag, "type").as_deref() == Some("checkbox"));
    tags("<button")
        .into_iter()
        .chain(checkboxes)
        .filter_map(|tag| attribute(&tag, "data-action").or_else(|| attribute(&tag, "id")))
        .collect()
}

fn read(path: &str) -> String {
    std::fs::read_to_string(path).unwrap_or_else(|error| panic!("{path}: {error}"))
}

fn checklist() -> String {
    let mut archive = archive();
    sheet_text(&mut archive, "Start here")
}

fn check_pane(pane: &str, html: &str, covered: &[(&str, &str)], exempt: &[(&str, &str)], text: &str) {
    let keys = control_keys(html);
    for key in &keys {
        let known = covered.iter().chain(exempt).any(|(name, _)| name == key);
        assert!(known, "{pane}: the control {key:?} has no Start here row; add one or exempt it with a reason");
    }
    for (key, _) in covered.iter().chain(exempt) {
        assert!(keys.iter().any(|name| name == key), "{pane}: {key:?} is no button or checkbox any more; drop its entry");
    }
    for (key, phrase) in covered {
        assert!(text.contains(phrase), "{pane}: {key:?} promises {phrase:?} but Start here does not say it");
    }
}

#[test]
fn start_here_names_every_excel_pane_button() {
    check_pane("Excel pane", &read("../taskpane.html"), EXCEL, EXCEL_EXEMPT, &checklist());
}

#[test]
fn start_here_names_every_powerpoint_pane_button() {
    check_pane("PowerPoint pane", &read("../pptpane.html"), PPT, PPT_EXEMPT, &checklist());
}

#[test]
fn start_here_sends_the_reader_to_every_sheet() {
    let text = checklist();
    for (name, _) in SHEETS {
        if name != "Start here" {
            assert!(text.contains(name), "Start here never mentions the {name} sheet");
        }
    }
}

#[test]
fn start_here_links_every_sheet_it_names_in_the_where_column() {
    let mut archive = archive();
    let sheet = common::sheet_xml(&mut archive, "Start here");
    let mut linked = Vec::new();
    for rest in sheet.split("<hyperlink ").skip(1) {
        let at = rest.find("location=\"'").expect("a location") + "location=\"'".len();
        let name = common::unescape(&rest[at..at + rest[at..].find("'!").expect("closing quote")]);
        assert!(SHEETS.iter().any(|(sheet, _)| *sheet == name), "a Where link points at {name:?}, no sheet");
        linked.push(name);
    }
    assert!(linked.len() >= 25, "only {} Where cells link to a sheet", linked.len());
    assert!(!linked.iter().any(|name| name == "Scratch"), "the hidden sheet must not be linked");
}

#[test]
fn every_step_is_complete_and_named_once() {
    let mut tools = Vec::new();
    for item in items() {
        let Item::Step(step) = item else { continue };
        for (what, value) in [("tool", &step.tool.to_string()), ("place", &step.place), ("action", &step.action), ("outcome", &step.outcome)] {
            assert!(!value.trim().is_empty(), "{}: empty {what}", step.tool);
            assert!(!value.contains(['{', '}']), "{}: a placeholder is left in the {what}", step.tool);
        }
        assert!(!tools.contains(&step.tool), "{}: two rows carry this tool name", step.tool);
        tools.push(step.tool);
    }
    assert!(tools.len() >= 50, "the checklist shrank to {} rows", tools.len());
}

#[test]
fn the_powerpoint_rows_run_in_an_order_a_reader_can_follow() {
    let steps: Vec<_> = items().into_iter().filter_map(|item| if let Item::Step(step) = item { Some(step) } else { None }).collect();
    let at = |tool: &str| steps.iter().position(|step| step.tool == tool).unwrap_or_else(|| panic!("no row is named {tool:?}"));
    // A link already on a slide leaves the Inbox, so each slide takes an export of its own: the first
    // export row makes the table, the Revenue chart and the Segment pie, the four quarters row exports
    // the table and the chart once more, and Paste latest linked takes the newest export last.
    let (exports, insert) = (&steps[at("Export selection (table), Export active chart")], at("Insert and update: table, chart, picture"));
    assert!(exports.action.contains("Segment pie"), "the first export row never exports the Segment pie");
    assert!(steps[at("Text, table and chart in four quarters")].action.contains("once more"), "Slide 5 needs exports of its own");
    // Convert to table acts on the Data picture, so it follows the Slide 4 insert; Revert needs the pasted
    // copies Clear pasted links removes, so that row comes after the Links tab rows.
    assert!(insert < at("Convert to table") && at("Convert to table") < at("Text, table and chart in four quarters"));
    assert!(at("Text, table and chart in four quarters") < at("Paste latest linked"));
    assert!(at("Paste latest linked") < at("Update selected, Update this slide, Revert last update"));
    assert!(at("Links list filters") < at("Clear pasted links"));
}

#[test]
fn no_shipped_string_has_a_placeholder_left() {
    let mut archive = archive();
    for string in shared_strings(&mut archive) {
        assert!(!string.contains(['{', '}']), "a placeholder is left in {string:?}");
    }
}

#[test]
fn no_shipped_string_has_an_em_dash() {
    let mut archive = archive();
    for string in shared_strings(&mut archive) {
        assert!(!string.contains('\u{2014}'), "em dash in {string:?}");
    }
}

/// The two Start here rows of the PowerPoint brand colors, by their tool names.
const BRAND_ROWS: &[&str] = &["Brand colors on shapes: swatches and Edit colors", "Brand colors in the deck: palette and theme"];

fn brand_rows() -> Vec<Step> {
    let rows: Vec<Step> = items().into_iter().filter_map(|item| if let Item::Step(step) = item { Some(step) } else { None }).filter(|step| BRAND_ROWS.contains(&step.tool)).collect();
    assert_eq!(rows.len(), BRAND_ROWS.len(), "the PowerPoint brand rows are not both there");
    rows
}

#[test]
fn the_powerpoint_brand_rows_spell_color_the_us_way() {
    for step in brand_rows() {
        for text in [step.tool, step.place.as_str(), step.action.as_str(), step.outcome.as_str()] {
            assert!(!text.to_lowercase().contains("colour"), "colour in {text:?}: the app speaks US English");
        }
    }
}

#[test]
fn the_powerpoint_brand_rows_promise_no_undo_nothing_has_proven() {
    // PowerPoint's own Undo has not been shown to take back a swatch or the theme on any host; Restore
    // theme colors is the way back for the theme, and the row says so.
    for step in brand_rows() {
        for text in [step.action.as_str(), step.outcome.as_str()] {
            assert!(!text.to_lowercase().contains("undo"), "an Undo promise in {text:?}");
        }
    }
}

/// The Start here row of the shape styles, by its tool name.
const SHAPE_STYLES_ROW: &str = "Shape styles: save, apply, update, delete, starters";

fn shape_styles_row() -> Step {
    items()
        .into_iter()
        .filter_map(|item| if let Item::Step(step) = item { Some(step) } else { None })
        .find(|step| step.tool == SHAPE_STYLES_ROW)
        .expect("the shape styles row is not there")
}

#[test]
fn the_shape_styles_row_follows_smart_painter_on_the_same_slide_and_speaks_us_english() {
    let steps: Vec<_> = items().into_iter().filter_map(|item| if let Item::Step(step) = item { Some(step) } else { None }).collect();
    let at = |tool: &str| steps.iter().position(|step| step.tool == tool).unwrap_or_else(|| panic!("no row is named {tool:?}"));
    // The pane shows Smart Painter, then Shape styles; both work on Slide 8's olive box and grey boxes.
    assert_eq!(at("Smart Painter") + 1, at(SHAPE_STYLES_ROW), "Shape styles belongs right after Smart Painter");
    let row = shape_styles_row();
    assert!(row.action.contains("Slide 8"), "the row never says which slide to use");
    for text in [row.tool, row.place.as_str(), row.action.as_str(), row.outcome.as_str()] {
        assert!(!text.to_lowercase().contains("colour"), "colour in {text:?}: the app speaks US English");
        assert!(!text.to_lowercase().contains("undo"), "an Undo promise in {text:?}: nothing has proven PowerPoint takes it back");
    }
}
