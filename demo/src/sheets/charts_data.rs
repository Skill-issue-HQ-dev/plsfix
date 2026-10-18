// The data behind the Charts sheet: twelve small DemoCo blocks, one per chart kind, in EUR
// thousands. Pure data, no workbook calls. Invariants: a series is as long as its block's
// labels, f64::NAN marks an empty cell (a gap in the chart), and the 2025E revenue is 13,392
// in every block that carries it (tests/charts_data.rs holds the ties).

use crate::layout::YEARS;

/// The chart kinds the Charts sheet carries, one block each.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Column,
    StackedColumn,
    PercentStackedColumn,
    Bar,
    Pie,
    Doughnut,
    Line,
    Area,
    StackedArea,
    Scatter,
    Combo,
    ColumnWithGap,
}

/// One column of numbers in a block, under its header.
pub struct Series {
    pub name: &'static str,
    pub values: &'static [f64],
    /// Shown as a percentage (0.1284 reads 12.8%) instead of EUR thousands.
    pub percent: bool,
}

/// One data block and the chart that sits beside it.
pub struct Block {
    pub kind: Kind,
    /// The chart's object name: what the tours pick it by.
    pub name: &'static str,
    pub title: &'static str,
    /// The bold line over the block, "<n>. <kind>: <what>".
    pub heading: &'static str,
    /// The header over the label column.
    pub corner: &'static str,
    pub labels: &'static [&'static str],
    pub series: &'static [Series],
    /// One line under the block for what is not obvious, or "".
    pub note: &'static str,
}

const MONTHS: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

// 2025E by month: the three arrays add up to the P&L's 2025E (13,392 / 5,089 / 2,143).
pub const REVENUE: [f64; 12] = [
    1020.0, 985.0, 1075.0, 1090.0, 1120.0, 1130.0, 1085.0, 1100.0, 1150.0, 1190.0, 1215.0, 1232.0,
];
pub const GROSS_PROFIT: [f64; 12] = [
    383.0, 364.0, 403.0, 413.0, 427.0, 432.0, 416.0, 418.0, 438.0, 456.0, 465.0, 474.0,
];
pub const EBITDA: [f64; 12] = [
    131.0, 116.0, 157.0, 169.0, 184.0, 188.0, 171.0, 176.0, 194.0, 211.0, 219.0, 227.0,
];
/// EBITDA over revenue to four places: block 11's line, recomputed by the test.
pub const EBITDA_MARGIN: [f64; 12] = [
    0.1284, 0.1178, 0.1460, 0.1550, 0.1643, 0.1664, 0.1576, 0.1600, 0.1687, 0.1773, 0.1802, 0.1843,
];

// 2024A-2029E: costs and EBITDA add up to the P&L's revenue; the markets add up to it too.
pub const COGS: [f64; 6] = [7688.0, 8303.0, 8812.0, 9265.0, 9729.0, 10131.0];
pub const OPEX: [f64; 6] = [2852.0, 2946.0, 3081.0, 3190.0, 3349.0, 3433.0];
pub const EBITDA_YEARLY: [f64; 6] = [1860.0, 2143.0, 2436.0, 2734.0, 2871.0, 3182.0];
pub const LATVIA: [f64; 6] = [6200.0, 6170.0, 6305.0, 6425.0, 6523.0, 6598.0];
pub const OTHER_BALTICS: [f64; 6] = [3162.0, 3822.0, 4241.0, 4602.0, 4881.0, 5158.0];
pub const EXPORT: [f64; 6] = [3038.0, 3400.0, 3783.0, 4162.0, 4545.0, 4990.0];

// Smallest first: a bar chart draws its first category at the bottom, so the biggest sits on top.
const COST_CATEGORIES: [&str; 8] = [
    "Logistics",
    "Energy",
    "Sales and marketing",
    "Rent and utilities",
    "Packaging",
    "Admin and IT",
    "Direct labour",
    "Raw materials",
];
pub const COSTS_2024: [f64; 8] = [668.0, 760.0, 790.0, 930.0, 880.0, 1132.0, 1520.0, 3860.0];
pub const COSTS_2025: [f64; 8] = [733.0, 790.0, 850.0, 946.0, 955.0, 1150.0, 1640.0, 4185.0];

const MARKETS: [&str; 8] = [
    "Latvia",
    "Lithuania",
    "Estonia",
    "Sweden",
    "Germany",
    "Poland",
    "Finland",
    "Other markets",
];
pub const MARKET_REVENUE: [f64; 8] = [6170.0, 2260.0, 1562.0, 880.0, 820.0, 690.0, 600.0, 410.0];

const PRODUCTS: [&str; 6] = [
    "Rye bread",
    "Wheat bread",
    "Pastries",
    "Crispbread",
    "Cakes",
    "Other products",
];
pub const PRODUCT_REVENUE: [f64; 6] = [3480.0, 3150.0, 2420.0, 1640.0, 1330.0, 1372.0];

// The three product lines by month: each adds up to its slice in the doughnut.
pub const RYE: [f64; 12] = [
    269.0, 257.0, 278.0, 281.0, 287.0, 289.0, 281.0, 287.0, 295.0, 301.0, 304.0, 351.0,
];
pub const WHEAT: [f64; 12] = [
    243.0, 235.0, 253.0, 259.0, 264.0, 267.0, 259.0, 259.0, 270.0, 275.0, 278.0, 288.0,
];
pub const PASTRIES: [f64; 12] = [
    173.0, 164.0, 190.0, 233.0, 199.0, 195.0, 182.0, 177.0, 194.0, 207.0, 225.0, 281.0,
];

const NET_DEBT: [f64; 12] = [
    3410.0, 3395.0, 3340.0, 3302.0, 3248.0, 3205.0, 3236.0, 3170.0, 3111.0, 3052.0, 2994.0, 2931.0,
];
const INVENTORIES: [f64; 12] = [
    820.0, 835.0, 842.0, 850.0, 861.0, 858.0, 870.0, 884.0, 902.0, 915.0, 931.0, 948.0,
];
const RECEIVABLES: [f64; 12] = [
    1180.0, 1172.0, 1205.0, 1226.0, 1240.0, 1252.0, 1238.0, 1249.0, 1281.0, 1305.0, 1330.0, 1361.0,
];

const CUSTOMERS: [&str; 15] = [
    "Customer 1",
    "Customer 2",
    "Customer 3",
    "Customer 4",
    "Customer 5",
    "Customer 6",
    "Customer 7",
    "Customer 8",
    "Customer 9",
    "Customer 10",
    "Customer 11",
    "Customer 12",
    "Customer 13",
    "Customer 14",
    "Customer 15",
];
const CUSTOMER_REVENUE: [f64; 15] = [
    1820.0, 1450.0, 1180.0, 990.0, 840.0, 705.0, 610.0, 520.0, 420.0, 365.0, 310.0, 230.0, 185.0,
    140.0, 92.0,
];
const CUSTOMER_MARGIN: [f64; 15] = [
    0.329, 0.341, 0.362, 0.355, 0.381, 0.372, 0.396, 0.389, 0.417, 0.402, 0.431, 0.425, 0.451,
    0.438, 0.462,
];

/// Where the one empty cell of block 12 sits (September).
pub const GAP_INDEX: usize = 8;
pub const EXPORT_SALES: [f64; 12] = [
    233.0,
    223.0,
    244.0,
    248.0,
    255.0,
    259.0,
    247.0,
    253.0,
    f64::NAN,
    266.0,
    277.0,
    296.0,
];

const fn eur(name: &'static str, values: &'static [f64]) -> Series {
    Series {
        name,
        values,
        percent: false,
    }
}

const fn percent(name: &'static str, values: &'static [f64]) -> Series {
    Series {
        name,
        values,
        percent: true,
    }
}

/// The twelve blocks in sheet order, top to bottom.
pub const BLOCKS: [Block; 12] = [
    Block {
        kind: Kind::Column,
        name: "Column chart",
        title: "Revenue, gross profit and EBITDA 2025E (EUR thousands)",
        heading: "1. Column: revenue, gross profit and EBITDA by month",
        corner: "Month",
        labels: &MONTHS,
        series: &[
            eur("Revenue", &REVENUE),
            eur("Gross profit", &GROSS_PROFIT),
            eur("EBITDA", &EBITDA),
        ],
        note: "",
    },
    Block {
        kind: Kind::StackedColumn,
        name: "Stacked column chart",
        title: "Revenue split into costs and EBITDA (EUR thousands)",
        heading: "2. Stacked column: revenue split into costs and EBITDA",
        corner: "Year",
        labels: &YEARS,
        series: &[
            eur("COGS", &COGS),
            eur("Opex", &OPEX),
            eur("EBITDA", &EBITDA_YEARLY),
        ],
        note: "",
    },
    Block {
        kind: Kind::PercentStackedColumn,
        name: "100% stacked column chart",
        title: "Revenue mix by market (share of revenue)",
        heading: "3. 100% stacked column: revenue mix by market",
        corner: "Year",
        labels: &YEARS,
        series: &[
            eur("Latvia", &LATVIA),
            eur("Other Baltics", &OTHER_BALTICS),
            eur("Export", &EXPORT),
        ],
        note: "EUR thousands here: the chart turns them into shares.",
    },
    Block {
        kind: Kind::Bar,
        name: "Bar chart",
        title: "Operating costs by category (EUR thousands)",
        heading: "4. Bar: operating costs by category",
        corner: "Cost category",
        labels: &COST_CATEGORIES,
        series: &[eur("2024A", &COSTS_2024), eur("2025E", &COSTS_2025)],
        note: "",
    },
    Block {
        kind: Kind::Pie,
        name: "Pie chart",
        title: "Revenue by market, 2025E (EUR thousands)",
        heading: "5. Pie: revenue by market, eight slices",
        corner: "Market",
        labels: &MARKETS,
        series: &[eur("Revenue", &MARKET_REVENUE)],
        note: "",
    },
    Block {
        kind: Kind::Doughnut,
        name: "Doughnut chart",
        title: "Revenue by product category, 2025E (EUR thousands)",
        heading: "6. Doughnut: revenue by product category",
        corner: "Product",
        labels: &PRODUCTS,
        series: &[eur("Revenue", &PRODUCT_REVENUE)],
        note: "",
    },
    Block {
        kind: Kind::Line,
        name: "Line chart",
        title: "Sales by product line 2025E (EUR thousands)",
        heading: "7. Line: sales by product line, markers on one line",
        corner: "Month",
        labels: &MONTHS,
        series: &[
            eur("Rye bread", &RYE),
            eur("Wheat bread", &WHEAT),
            eur("Pastries", &PASTRIES),
        ],
        note: "Markers on Rye bread only; the other two lines have none.",
    },
    Block {
        kind: Kind::Area,
        name: "Area chart",
        title: "Net debt at month end, 2025E (EUR thousands)",
        heading: "8. Area: net debt at month end",
        corner: "Month",
        labels: &MONTHS,
        series: &[eur("Net debt", &NET_DEBT)],
        note: "",
    },
    Block {
        kind: Kind::StackedArea,
        name: "Stacked area chart",
        title: "Working capital at month end, 2025E (EUR thousands)",
        heading: "9. Stacked area: working capital at month end",
        corner: "Month",
        labels: &MONTHS,
        series: &[
            eur("Inventories", &INVENTORIES),
            eur("Receivables", &RECEIVABLES),
        ],
        note: "",
    },
    Block {
        kind: Kind::Scatter,
        name: "Scatter chart",
        title: "Customer revenue against gross margin 2025E",
        heading: "10. Scatter: fifteen customers, revenue against margin",
        corner: "Customer",
        labels: &CUSTOMERS,
        series: &[
            eur("Revenue", &CUSTOMER_REVENUE),
            percent("Gross margin", &CUSTOMER_MARGIN),
        ],
        note: "",
    },
    Block {
        kind: Kind::Combo,
        name: "Combo chart",
        title: "Monthly revenue and EBITDA margin, 2025E",
        heading: "11. Combo: revenue columns, margin line on a second axis",
        corner: "Month",
        labels: &MONTHS,
        series: &[
            eur("Revenue", &REVENUE),
            percent("EBITDA margin", &EBITDA_MARGIN),
        ],
        note: "EBITDA margin = EBITDA / revenue of block 1, on the right-hand axis.",
    },
    Block {
        kind: Kind::ColumnWithGap,
        name: "Column chart with a gap",
        title: "Export sales by month, 2024A (September not reported)",
        heading: "12. Column with a gap: export sales, September empty",
        corner: "Month",
        labels: &MONTHS,
        series: &[eur("Export sales", &EXPORT_SALES)],
        note: "September is empty on purpose: the chart shows a gap, not a zero.",
    },
];
