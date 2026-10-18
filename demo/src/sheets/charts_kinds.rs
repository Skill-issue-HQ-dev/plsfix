// The twelve chart builders of the Charts sheet, one per kind, each reading its block through a
// Place. Owns names, titles, colours, labels and axis scales; charts.rs owns where charts sit.
// Invariant: every label position is one Excel accepts for that chart type (a stacked column
// never gets Outside end, a doughnut and an area get no position at all).

use rust_xlsxwriter::{
    Chart, ChartDataLabel, ChartDataLabelPosition, ChartEmptyCells, ChartFont, ChartFormat,
    ChartLegendPosition, ChartLine, ChartMarker, ChartMarkerType, ChartPoint, ChartSeries,
    ChartSolidFill, ChartType, Color,
};

use crate::sheets::charts_data::{Block, Kind};
use crate::style::{text_on, ACCENT, INK, SERIES};

const CHART_WIDTH: u32 = 520;
const CHART_HEIGHT: u32 = 300;
const TITLE_SIZE: f64 = 11.0;
const LABEL_SIZE: f64 = 8.0;
/// Column labels read bottom to top, so 36 of them fit the 12 slots.
const LABEL_TURN: i16 = -90;
const COLUMN_GAP: u16 = 40;
const STACK_GAP: u16 = 60;
const LINE_WIDTH: f64 = 2.25;
const MARKER_SIZE: u8 = 7;
const DOT_SIZE: u8 = 8;
const HOLE_SIZE: u8 = 55;
const AXIS_EUR: &str = "#,##0";
const AXIS_PERCENT: &str = "0%";
/// The combo's two scales: revenue columns stay in the lower half, the margin line above them,
/// and both axes carry four steps so the gridlines meet the right-hand ticks.
const REVENUE_MAX: f64 = 2400.0;
const REVENUE_STEP: f64 = 600.0;
const MARGIN_MAX: f64 = 0.2;
const MARGIN_STEP: f64 = 0.05;
/// The scatter's margin axis frames the 33-46% spread; the bar and gap charts keep headroom for
/// their outside-end labels.
const MARGIN_FLOOR: f64 = 0.3;
const MARGIN_CEILING: f64 = 0.5;
const COSTS_MAX: f64 = 5000.0;
const COSTS_STEP: f64 = 1000.0;
const EXPORT_MAX: f64 = 400.0;

/// Where a block sits on its sheet (0-based): the rows and columns its chart reads.
pub struct Place {
    pub sheet: &'static str,
    pub header: u32,
    pub first: u32,
    pub last: u32,
    pub label_col: u16,
    pub series_col: u16,
}

impl Place {
    fn labels(&self) -> (&'static str, u32, u16, u32, u16) {
        (
            self.sheet,
            self.first,
            self.label_col,
            self.last,
            self.label_col,
        )
    }

    fn values(&self, index: usize) -> (&'static str, u32, u16, u32, u16) {
        let col = self.series_col + index as u16;
        (self.sheet, self.first, col, self.last, col)
    }

    fn name(&self, index: usize) -> (&'static str, u32, u16) {
        (self.sheet, self.header, self.series_col + index as u16)
    }
}

/// Each of a block's series with its colour: the first palette entries, one per series.
fn palette(block: &Block) -> impl Iterator<Item = (usize, Color)> {
    SERIES.into_iter().enumerate().take(block.series.len())
}

/// The chart for one block, named and titled, at the sheet's one chart size.
pub fn chart(block: &Block, place: &Place) -> Chart {
    match block.kind {
        Kind::Column => column(block, place),
        Kind::StackedColumn => stacked_column(block, place, false),
        Kind::PercentStackedColumn => stacked_column(block, place, true),
        Kind::Bar => bar(block, place),
        Kind::Pie => pie(block, place),
        Kind::Doughnut => doughnut(block, place),
        Kind::Line => line(block, place),
        Kind::Area => area(block, place, false),
        Kind::StackedArea => area(block, place, true),
        Kind::Scatter => scatter(block, place),
        Kind::Combo => combo(block, place),
        Kind::ColumnWithGap => column_with_gap(block, place),
    }
}

fn base(chart_type: ChartType, block: &Block) -> Chart {
    let mut chart = Chart::new(chart_type);
    chart.set_name(block.name);
    chart
        .title()
        .set_name(block.title)
        .set_font(ChartFont::new().set_size(TITLE_SIZE));
    chart.set_width(CHART_WIDTH).set_height(CHART_HEIGHT);
    chart
}

/// A new series over the block's label column and its `index`-th value column, named by the
/// header cell.
fn series<'c>(chart: &'c mut Chart, place: &Place, index: usize) -> &'c mut ChartSeries {
    chart
        .add_series()
        .set_name(place.name(index))
        .set_categories(place.labels())
        .set_values(place.values(index))
}

/// A round marker with fill and border in one colour, so Excel adds no outline of its own.
fn dot(color: Color, size: u8) -> ChartMarker {
    let mut marker = ChartMarker::new();
    marker
        .set_type(ChartMarkerType::Circle)
        .set_size(size)
        .set_format(
            ChartFormat::new()
                .set_solid_fill(ChartSolidFill::new().set_color(color))
                .set_border(ChartLine::new().set_color(color)),
        );
    marker
}

/// Labels that show the value and nothing else, the rule every chart of the demo follows.
fn value_labels() -> ChartDataLabel {
    let mut labels = ChartDataLabel::new();
    labels.show_value();
    labels
}

fn column(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Column, block);
    let mut labels = value_labels();
    labels.set_font(
        ChartFont::new()
            .set_size(LABEL_SIZE)
            .set_rotation(LABEL_TURN),
    );
    for (index, color) in palette(block) {
        series(&mut chart, place, index)
            .set_format(ChartSolidFill::new().set_color(color))
            .set_gap(COLUMN_GAP)
            .set_data_label(&labels);
    }
    chart.y_axis().set_num_format(AXIS_EUR);
    chart.legend().set_position(ChartLegendPosition::Bottom);
    chart
}

fn stacked_column(block: &Block, place: &Place, percent: bool) -> Chart {
    let kind = if percent {
        ChartType::ColumnPercentStacked
    } else {
        ChartType::ColumnStacked
    };
    let mut chart = base(kind, block);
    for (index, color) in palette(block) {
        let one = series(&mut chart, place, index);
        one.set_format(ChartSolidFill::new().set_color(color))
            .set_gap(STACK_GAP);
        if !percent {
            one.set_data_label(&centred_labels(color));
        }
    }
    chart
        .y_axis()
        .set_num_format(if percent { AXIS_PERCENT } else { AXIS_EUR });
    chart.legend().set_position(ChartLegendPosition::Bottom);
    chart
}

/// Value labels in the middle of a stacked segment, lettered to read on the segment's `fill`.
fn centred_labels(fill: Color) -> ChartDataLabel {
    let mut labels = value_labels();
    labels
        .set_position(ChartDataLabelPosition::Center)
        .set_font(ChartFont::new().set_color(text_on(fill)));
    labels
}

fn bar(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Bar, block);
    let labels = value_labels();
    for (index, color) in palette(block) {
        series(&mut chart, place, index)
            .set_format(ChartSolidFill::new().set_color(color))
            .set_gap(COLUMN_GAP)
            .set_data_label(&labels);
    }
    // In a bar chart the horizontal axis is the value axis.
    chart
        .x_axis()
        .set_num_format(AXIS_EUR)
        .set_min(0)
        .set_max(COSTS_MAX)
        .set_major_unit(COSTS_STEP);
    chart.legend().set_position(ChartLegendPosition::Bottom);
    chart
}

fn slices(colors: &[Color]) -> Vec<ChartPoint> {
    colors
        .iter()
        .map(|color| ChartPoint::new().set_format(ChartSolidFill::new().set_color(*color)))
        .collect()
}

fn pie(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Pie, block);
    let mut labels = value_labels();
    labels.set_position(ChartDataLabelPosition::OutsideEnd);
    series(&mut chart, place, 0)
        .set_points(&slices(&SERIES[..block.labels.len()]))
        .set_data_label(&labels);
    chart
}

/// One slice per label in the series colours; each slice's label is lettered to read on its own
/// colour, white on the dark slices and ink on the light ones.
fn doughnut(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Doughnut, block);
    chart.set_hole_size(HOLE_SIZE);
    let colors = &SERIES[..block.labels.len()];
    let labels: Vec<ChartDataLabel> = colors
        .iter()
        .map(|color| {
            value_labels()
                .set_font(ChartFont::new().set_color(text_on(*color)))
                .to_custom()
        })
        .collect();
    series(&mut chart, place, 0)
        .set_points(&slices(colors))
        .set_data_label(&value_labels())
        .set_custom_data_labels(&labels);
    chart
}

fn line(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Line, block);
    for (index, color) in palette(block) {
        // Markers on the first line only: the other two are drawn without.
        let marker = if index == 0 {
            dot(color, MARKER_SIZE)
        } else {
            let mut none = ChartMarker::new();
            none.set_none();
            none
        };
        series(&mut chart, place, index)
            .set_format(ChartLine::new().set_color(color).set_width(LINE_WIDTH))
            .set_marker(&marker);
    }
    chart.y_axis().set_num_format(AXIS_EUR);
    chart.legend().set_position(ChartLegendPosition::Bottom);
    chart
}

/// An area chart starts at zero, whatever Excel would pick for a narrow range.
fn area(block: &Block, place: &Place, stacked: bool) -> Chart {
    let mut chart = base(
        if stacked {
            ChartType::AreaStacked
        } else {
            ChartType::Area
        },
        block,
    );
    for (index, color) in palette(block) {
        let color = if stacked { color } else { ACCENT };
        series(&mut chart, place, index).set_format(ChartSolidFill::new().set_color(color));
    }
    chart.y_axis().set_num_format(AXIS_EUR).set_min(0);
    if stacked {
        chart.legend().set_position(ChartLegendPosition::Bottom);
    } else {
        chart.legend().set_hidden();
    }
    chart
}

/// Series 0 is the X value and series 1 the Y value of every point.
fn scatter(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Scatter, block);
    chart
        .add_series()
        .set_name(place.name(1))
        .set_categories(place.values(0))
        .set_values(place.values(1))
        .set_marker(&dot(ACCENT, DOT_SIZE));
    chart
        .x_axis()
        .set_name("Revenue, EUR thousands")
        .set_num_format(AXIS_EUR);
    chart
        .y_axis()
        .set_name("Gross margin")
        .set_num_format(AXIS_PERCENT)
        .set_min(MARGIN_FLOOR)
        .set_max(MARGIN_CEILING)
        .set_major_unit(MARGIN_STEP);
    chart.legend().set_hidden();
    chart
}

/// Revenue as columns, the EBITDA margin as a labelled line on the secondary axis.
fn combo(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Column, block);
    series(&mut chart, place, 0)
        .set_format(ChartSolidFill::new().set_color(ACCENT))
        .set_gap(COLUMN_GAP);
    let mut margin = Chart::new(ChartType::Line);
    let mut labels = value_labels();
    labels.set_position(ChartDataLabelPosition::Above);
    series(&mut margin, place, 1)
        .set_format(ChartLine::new().set_color(INK).set_width(LINE_WIDTH))
        .set_marker(&dot(INK, MARKER_SIZE))
        .set_data_label(&labels)
        .set_secondary_axis(true);
    chart.combine(&margin);
    chart
        .y_axis()
        .set_num_format(AXIS_EUR)
        .set_min(0)
        .set_max(REVENUE_MAX)
        .set_major_unit(REVENUE_STEP);
    chart
        .y2_axis()
        .set_num_format(AXIS_PERCENT)
        .set_min(0)
        .set_max(MARGIN_MAX)
        .set_major_unit(MARGIN_STEP);
    chart.legend().set_position(ChartLegendPosition::Bottom);
    chart
}

/// One month has no value: the empty cell stays a gap, never a zero.
fn column_with_gap(block: &Block, place: &Place) -> Chart {
    let mut chart = base(ChartType::Column, block);
    chart.show_empty_cells_as(ChartEmptyCells::Gaps);
    series(&mut chart, place, 0)
        .set_format(ChartSolidFill::new().set_color(ACCENT))
        .set_gap(COLUMN_GAP)
        .set_data_label(&value_labels());
    chart
        .y_axis()
        .set_num_format(AXIS_EUR)
        .set_min(0)
        .set_max(EXPORT_MAX);
    chart.legend().set_hidden();
    chart
}
