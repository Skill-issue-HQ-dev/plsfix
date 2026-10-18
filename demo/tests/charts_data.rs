// Holds the Charts sheet's data blocks to each other: one 2025E revenue of 13,392 runs through
// every block that carries it, the combo's margin is EBITDA over revenue, and the gap block has
// exactly one empty month.

use smt_demo::sheets::charts_data as data;

/// The P&L's 2025E revenue: 12,400 grown by the 8 % in Assumptions.
const REVENUE_2025E: f64 = 13_392.0;
/// The P&L's 2025E gross profit and EBITDA as the sheet shows them, rounded to the thousand.
const GROSS_PROFIT_2025E: f64 = 5_089.0;
const EBITDA_2025E: f64 = 2_143.0;
/// A margin stored to four places is off by at most half of the last place.
const MARGIN_TOLERANCE: f64 = 0.00005;

fn total(values: &[f64]) -> f64 {
    values.iter().sum()
}

// A banker tracing a number must never meet a second revenue: the months, the markets, the
// products and the 2025E column of both yearly blocks are cuts of the same 13,392.
#[test]
fn charts_data_ties_to_one_revenue_figure() {
    assert_eq!(total(&data::REVENUE), REVENUE_2025E);
    assert_eq!(total(&data::MARKET_REVENUE), REVENUE_2025E);
    assert_eq!(total(&data::PRODUCT_REVENUE), REVENUE_2025E);
    assert_eq!(total(&data::GROSS_PROFIT), GROSS_PROFIT_2025E);
    assert_eq!(total(&data::EBITDA), EBITDA_2025E);
    for year in 0..6 {
        let split = data::COGS[year] + data::OPEX[year] + data::EBITDA_YEARLY[year];
        let mix = data::LATVIA[year] + data::OTHER_BALTICS[year] + data::EXPORT[year];
        assert_eq!(split, mix, "year {year}: two cuts of one revenue");
    }
    let (cogs, opex, ebitda) = (data::COGS[1], data::OPEX[1], data::EBITDA_YEARLY[1]);
    assert_eq!(cogs + opex + ebitda, REVENUE_2025E);
    assert_eq!(total(&data::COSTS_2024), data::COGS[0] + data::OPEX[0]);
    assert_eq!(total(&data::COSTS_2025), cogs + opex);
}

// The pie's eight markets regroup into the 100% stack's three, and each product line's twelve
// months add up to its slice of the doughnut.
#[test]
fn charts_data_regroups_markets_and_products() {
    let markets = data::MARKET_REVENUE;
    assert_eq!(data::LATVIA[1], markets[0]);
    assert_eq!(data::OTHER_BALTICS[1], markets[1] + markets[2]);
    assert_eq!(data::EXPORT[1], total(&markets[3..]));
    for (line, slice) in [(data::RYE, 0), (data::WHEAT, 1), (data::PASTRIES, 2)] {
        assert_eq!(
            total(&line),
            data::PRODUCT_REVENUE[slice],
            "product line {slice}"
        );
    }
}

#[test]
fn charts_data_margin_is_ebitda_over_revenue() {
    for month in 0..12 {
        let exact = data::EBITDA[month] / data::REVENUE[month];
        assert!(
            (data::EBITDA_MARGIN[month] - exact).abs() <= MARGIN_TOLERANCE,
            "month {month}"
        );
    }
}

#[test]
fn charts_data_has_exactly_one_empty_month() {
    let empty: Vec<usize> = (0..12)
        .filter(|month| data::EXPORT_SALES[*month].is_nan())
        .collect();
    assert_eq!(empty, [data::GAP_INDEX]);
}
