// The two C0b demo charts, `Doughnut chart` and `Combo chart` of the demo
// workbook's Charts sheet (demo/src/sheets/charts_data.rs), as the Excel reader
// ships them as chart data: the numbers of the blocks, the text each cell shows
// (EUR thousands as "#,##0", the margin as "0.0%"), the brand colours the
// reader's rules give (a doughnut's slices by position, a combo's series one
// each), the secondary axis and the 55 % hole Excel reports. Owns no assertions.

import { pieColors, seriesColor } from "../src/chart-colors";
import type { ChartData } from "../src/link/chart-model";

const BRAND = { primary: "#151C2C", accent: "#B8902D" };

const PRODUCTS = [
  "Rye bread",
  "Wheat bread",
  "Pastries",
  "Crispbread",
  "Cakes",
  "Other products",
];
const PRODUCT_REVENUE = [3480, 3150, 2420, 1640, 1330, 1372];

const MONTHS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];
const REVENUE = [
  1020, 985, 1075, 1090, 1120, 1130, 1085, 1100, 1150, 1190, 1215, 1232,
];
const EBITDA_MARGIN = [
  0.1284, 0.1178, 0.146, 0.155, 0.1643, 0.1664, 0.1576, 0.16, 0.1687, 0.1773,
  0.1802, 0.1843,
];

// "#,##0" as Excel shows it: a comma between thousands.
function thousands(value: number): string {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// "0.0%" as Excel shows it.
function percent(value: number): string {
  return `${(value * 100).toFixed(1)}%`;
}

const FRAME = {
  v: 1 as const,
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: BRAND.primary,
};

export const DEMO_DOUGHNUT: ChartData = {
  ...FRAME,
  kind: "doughnut",
  title: "Revenue by product category, 2025E (EUR thousands)",
  categories: PRODUCTS,
  series: [
    {
      name: "Revenue",
      values: PRODUCT_REVENUE,
      labels: PRODUCT_REVENUE.map(thousands),
      colors: pieColors(PRODUCTS.length, BRAND),
    },
  ],
  hole: 55,
};

export const DEMO_COMBO: ChartData = {
  ...FRAME,
  kind: "combo",
  title: "Monthly revenue and EBITDA margin, 2025E",
  categories: MONTHS,
  series: [
    {
      name: "Revenue",
      type: "column",
      values: REVENUE,
      labels: REVENUE.map(thousands),
      colors: MONTHS.map(() => seriesColor(0, BRAND)),
    },
    {
      name: "EBITDA margin",
      type: "line",
      axis: "secondary",
      values: EBITDA_MARGIN,
      labels: EBITDA_MARGIN.map(percent),
      colors: MONTHS.map(() => seriesColor(1, BRAND)),
    },
  ],
};
