// Brand colours and the handful of cell formats the demo sheets share. Fonts
// keep Excel's default colour on purpose: Autocolor is meant to have work to do.
// The palette is the pane's Highlighter look (src/styles/tokens.css, the defaults in
// src/settings.ts); the modelling colours (blue inputs, green links, ...) are never painted here.

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, FormatUnderline};

pub const INK: Color = Color::RGB(0x11_11_11);
/// The olive accent: the first chart series and every single-series chart (white lettering on it
/// reads 4.5:1).
pub const ACCENT: Color = Color::RGB(0x6D_7E_12);
/// The accent dark enough to read as words on white (7.5:1): the sheet notes.
pub const ACCENT_TEXT: Color = Color::RGB(0x4D_5B_00);
/// The guide band: lime at 20 % on white, with ink text on it (17.9:1).
pub const GUIDE_TINT: Color = Color::RGB(0xF7_FC_DB);
/// Chart series colours in the order the pane paints them: the accent, the ink, then tints of
/// both (the first six), then the same six shaded (series seven to twelve). Mirrors
/// `seriesColor` in src/chart-colors.ts at the shipped defaults (primary #111111, accent
/// #6D7E12). A chart takes one per series, a pie or ring one per slice, a single-series chart the
/// first.
pub const SERIES: [Color; 12] = [
    ACCENT,
    INK,
    Color::RGB(0x94_94_94),
    Color::RGB(0xAF_B8_7D),
    Color::RGB(0xCB_CB_CB),
    Color::RGB(0xD3_D8_B8),
    Color::RGB(0x4C_58_0D),
    Color::RGB(0x0C_0C_0C),
    Color::RGB(0x68_68_68),
    Color::RGB(0x7B_81_58),
    Color::RGB(0x8E_8E_8E),
    Color::RGB(0x94_97_81),
];
pub const EUR_K: &str = "#,##0;(#,##0);\"-\"";
pub const PERCENT: &str = "0.0%";
pub const PERCENT_2: &str = "0.00%";
pub const POINTS: &str = "0.00";
pub const WHOLE: &str = "0";
pub const NUMBER_1DP: &str = "#,##0.0";
pub const MULTIPLE: &str = "0.0\"x\"";
pub const TWO_DP: &str = "#,##0.00";
pub const EURO_2DP: &str = "\"€\"#,##0.00";
pub const DATE_DMY: &str = "dd.mm.yyyy";
const TITLE_SIZE: f64 = 14.0;
/// Brightness (0-255, the weighted channel average) from which a fill counts as light and takes
/// ink lettering: the same line as `LIGHT_FILL_BRIGHTNESS` in src/settings.ts.
const LIGHT_FILL_BRIGHTNESS: u32 = 140;

/// The lettering that reads on `fill`: white on a dark fill, ink on a light one (a colour that is
/// not RGB counts as light). Mirrors `contrastText` in src/settings.ts; chart labels that sit on a
/// segment or a slice use it.
pub fn text_on(fill: Color) -> Color {
    let Color::RGB(rgb) = fill else { return INK };
    let (r, g, b) = ((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
    if 299 * r + 587 * g + 114 * b >= LIGHT_FILL_BRIGHTNESS * 1000 {
        INK
    } else {
        Color::White
    }
}

pub struct Styles {
    pub title: Format,
    pub note: Format,
    pub header: Format,
    pub header_left: Format,
    pub label: Format,
    pub label_bold: Format,
    pub eur: Format,
    pub eur_bold: Format,
    pub pct: Format,
    pub pct2: Format,
    pub points: Format,
    pub whole: Format,
    pub number_1dp: Format,
    pub plain: Format,
    pub multiple: Format,
    pub two_dp: Format,
    pub euro_2dp: Format,
    pub date: Format,
    /// A clickable index entry: the sheet-note green, underlined (no browser blue).
    pub link: Format,
    /// Long text that wraps inside its cell (the Headlines paragraph).
    pub wrap: Format,
    /// The formatted blank far from the data that Clean past the data removes.
    pub stray: Format,
    /// Two pre-styled cells for the Format lab's paint slots.
    pub sample_text: Format,
    pub sample_number: Format,
    pub guide_title: Format,
    pub guide_text: Format,
}

impl Styles {
    pub fn new() -> Styles {
        Styles {
            title: Format::new().set_bold().set_font_size(TITLE_SIZE).set_font_color(INK),
            note: Format::new().set_italic().set_font_color(ACCENT_TEXT),
            header: Format::new()
                .set_bold()
                .set_align(FormatAlign::Right)
                .set_border_bottom(FormatBorder::Thin),
            header_left: Format::new().set_bold().set_border_bottom(FormatBorder::Thin),
            label: Format::new(),
            label_bold: Format::new().set_bold(),
            eur: Format::new().set_num_format(EUR_K),
            eur_bold: Format::new().set_bold().set_num_format(EUR_K),
            pct: Format::new().set_num_format(PERCENT),
            pct2: Format::new().set_num_format(PERCENT_2),
            points: Format::new().set_num_format(POINTS),
            whole: Format::new().set_num_format(WHOLE),
            number_1dp: Format::new().set_num_format(NUMBER_1DP),
            plain: Format::new(),
            multiple: Format::new().set_num_format(MULTIPLE),
            two_dp: Format::new().set_num_format(TWO_DP),
            euro_2dp: Format::new().set_num_format(EURO_2DP),
            date: Format::new().set_num_format(DATE_DMY),
            link: Format::new().set_font_color(ACCENT_TEXT).set_underline(FormatUnderline::Single),
            wrap: Format::new().set_text_wrap().set_align(FormatAlign::Top),
            stray: Format::new().set_background_color(GUIDE_TINT),
            sample_text: Format::new()
                .set_bold()
                .set_font_color(ACCENT_TEXT)
                .set_background_color(GUIDE_TINT)
                .set_border_bottom(FormatBorder::Thin),
            sample_number: Format::new().set_num_format(MULTIPLE).set_border(FormatBorder::Thin),
            // No wrap: column A is as narrow as 4 characters on some sheets
            // (Start here), so wrapped text broke into unreadable slivers.
            // The line overflows across the empty cells to its right instead
            // (rows 1-7 hold column A only, everywhere); set_row_format
            // paints the tint over those cells since they carry no format
            // of their own.
            guide_title: Format::new().set_bold().set_font_color(INK).set_background_color(GUIDE_TINT),
            guide_text: Format::new().set_background_color(GUIDE_TINT),
        }
    }
}

impl Default for Styles {
    fn default() -> Self {
        Self::new()
    }
}
