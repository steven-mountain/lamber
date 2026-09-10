//! Inclusive quote/limit/markup; tariff and service fee exclusive of tax.
//! Round the exclusive fee half-up to cents FIRST, then tax it and round again.
use super::models::SelectionFeeResult;
use rust_decimal::{prelude::*, RoundingStrategy};
use std::str::FromStr;

const TAX_PERCENT: i64 = 6;
const CENT: Decimal = Decimal::from_parts(1, 0, 0, false, 2);

#[derive(Clone, Copy)]
enum Charge {
    Fixed(Decimal),
    Rate(Decimal),
}
struct Band {
    lower: Decimal,
    lower_closed: bool,
    upper: Option<Decimal>,
    upper_closed: bool,
    charge: Charge,
}

// The only tariff source. All inclusive bounds (also for reverse search) derive from it.
fn bands() -> [Band; 5] {
    let d = Decimal::from;
    [
        Band {
            lower: d(0),
            lower_closed: false,
            upper: Some(d(12100)),
            upper_closed: true,
            charge: Charge::Fixed(d(100)),
        },
        Band {
            lower: d(12100),
            lower_closed: false,
            upper: Some(d(48500)),
            upper_closed: false,
            charge: Charge::Rate(Decimal::new(825, 5)),
        },
        Band {
            lower: d(48500),
            lower_closed: true,
            upper: Some(d(100000)),
            upper_closed: true,
            charge: Charge::Fixed(d(400)),
        },
        Band {
            lower: d(100000),
            lower_closed: false,
            upper: Some(d(1000000)),
            upper_closed: true,
            charge: Charge::Rate(Decimal::new(809985, 8)),
        },
        Band {
            lower: d(1000000),
            lower_closed: false,
            upper: None,
            upper_closed: false,
            charge: Charge::Fixed(Decimal::new(809985, 2)),
        },
    ]
}
fn round(value: Decimal) -> Decimal {
    value.round_dp_with_strategy(2, RoundingStrategy::MidpointAwayFromZero)
}
fn factor(tax: i64) -> Decimal {
    Decimal::ONE + Decimal::new(tax, 2)
}
fn money(value: &str, label: &str, signed: bool) -> Result<Decimal, String> {
    let number = Decimal::from_str(value.trim()).map_err(|_| format!("{label}须为有效金额"))?;
    if (!signed && number < Decimal::ZERO)
        || number.scale() > 2
        || number
            .checked_mul(Decimal::from(100))
            .and_then(|n| n.to_i64())
            .is_none()
    {
        return Err(format!(
            "{label}须为{}最多两位小数、在可计算范围内的金额",
            if signed { "" } else { "非负、" }
        ));
    }
    Ok(number)
}
impl Band {
    fn contains(&self, quote: Decimal, tax: i64) -> bool {
        // Compare before division to avoid repeating decimal division at exact bounds.
        let lower = self.lower * factor(tax);
        (quote > lower || (self.lower_closed && quote == lower))
            && self.upper.map_or(true, |upper| {
                let upper = upper * factor(tax);
                quote < upper || (self.upper_closed && quote == upper)
            })
    }
    fn fee(&self, quote: Decimal, tax: i64) -> (Decimal, Decimal) {
        let excl = round(match self.charge {
            Charge::Fixed(fee) => fee,
            Charge::Rate(rate) => quote / factor(tax) * rate,
        });
        (excl, round(excl * factor(tax)))
    }
    fn cent_bounds(&self, tax: i64, cap: i64) -> (i64, i64) {
        let lower = self.lower * factor(tax) / CENT;
        let lo = if self.lower_closed {
            lower.ceil()
        } else {
            lower.floor() + Decimal::ONE
        };
        let hi = self
            .upper
            .map(|upper| {
                let upper = upper * factor(tax) / CENT;
                if self.upper_closed {
                    upper.floor()
                } else {
                    upper.ceil() - Decimal::ONE
                }
            })
            .and_then(|upper| upper.to_i64())
            .unwrap_or(cap);
        (lo.to_i64().unwrap(), hi.min(cap))
    }
}
fn calculate_at(quote: Decimal, markup: Decimal, tax: i64) -> Result<SelectionFeeResult, String> {
    let (excl, incl) = if quote.is_zero() {
        (Decimal::ZERO, Decimal::ZERO)
    } else {
        bands()
            .iter()
            .find(|band| band.contains(quote, tax))
            .ok_or("报价不在资费表范围内")?
            .fee(quote, tax)
    };
    let actual = quote + incl;
    if actual + markup < Decimal::ZERO {
        return Err("浮动后的含税限价不能为负数".into());
    }
    Ok(SelectionFeeResult {
        selection_fee: format!("{excl:.2}"),
        selection_fee_excl: format!("{excl:.2}"),
        selection_fee_incl: format!("{incl:.2}"),
        quote_excl: format!("{:.4}", quote / factor(tax)),
        quote_candidates: vec![],
        quote: format!("{quote:.2}"),
        actual_cost: format!("{actual:.2}"),
        final_limit: format!("{:.2}", actual + markup),
    })
}
pub fn calculate(quote: &str, markup: &str) -> Result<SelectionFeeResult, String> {
    calculate_at(
        money(quote, "含税报价", false)?,
        money(markup, "含税浮动", true)?,
        TAX_PERCENT,
    )
}
fn reverse_at(limit: Decimal, markup: Decimal, tax: i64) -> Result<SelectionFeeResult, String> {
    let target = limit - markup;
    if target < Decimal::ZERO {
        return Err("含税限价不能小于含税浮动，无对应报价".into());
    }
    if target.is_zero() {
        return calculate_at(Decimal::ZERO, markup, tax);
    }
    let cap = (target / CENT).to_i64().ok_or("含税限价超出可计算范围")?;
    let mut candidates = vec![];
    // Forward is strictly increasing WITHIN each band, but not across all bands.
    // Search integral cents; only return candidates whose rounded forward total is exact.
    for band in bands() {
        let (mut lo, mut hi) = band.cent_bounds(tax, cap);
        while lo <= hi {
            let mid = lo + (hi - lo) / 2;
            let quote = Decimal::from(mid) * CENT;
            let actual = quote + band.fee(quote, tax).1;
            match actual.cmp(&target) {
                std::cmp::Ordering::Less => lo = mid + 1,
                std::cmp::Ordering::Greater => hi = mid - 1,
                std::cmp::Ordering::Equal => {
                    candidates.push(quote);
                    break;
                }
            }
        }
    }
    candidates.sort();
    if let Some(quote) = candidates.first() {
        let mut result = calculate_at(*quote, markup, tax)?;
        result.quote_candidates = candidates.iter().map(|q| format!("{q:.2}")).collect();
        return Ok(result);
    }
    let tariff = bands();
    let edge = tariff[2].upper.unwrap() * factor(tax);
    let below = edge + tariff[2].fee(edge, tax).1;
    let next = edge + CENT;
    let above = next + tariff[3].fee(next, tax).1;
    if target > below && target < above {
        return Err("该限价落在资费表 10 万元档位跳变形成的空档，无对应报价".into());
    }
    Err("该限价按资费表及逐分舍入口径无对应报价，请调整限价或改用报价正算".into())
}
pub fn reverse(limit: &str, markup: &str) -> Result<SelectionFeeResult, String> {
    reverse_at(
        money(limit, "含税限价", false)?,
        money(markup, "含税浮动", true)?,
        TAX_PERCENT,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    fn d(s: &str) -> Decimal {
        Decimal::from_str(s).unwrap()
    }

    #[test]
    fn tariff_midpoints_boundaries_and_discrepancies() {
        // Independent expected values transcribed from the tariff, not generated by bands().
        for (quote_excl, fee) in [
            ("5000", "100.00"),
            ("30000", "247.50"),
            ("70000", "400.00"),
            ("300000", "2429.96"),
            ("500000", "4049.93"),
            ("600000", "4859.91"),
            ("2000000", "8099.85"),
            ("12100", "100.00"),
            ("12100.01", "99.83"),
            ("48500", "400.00"),
            ("100000", "400.00"),
            ("100000.01", "809.99"),
            ("1000000", "8099.85"),
            ("1000000.01", "8099.85"),
        ] {
            let quote = round(d(quote_excl) * d("1.06"));
            let result = calculate(&quote.to_string(), "0").unwrap();
            assert_eq!(
                result.selection_fee_excl, fee,
                "exclusive base {quote_excl}"
            );
            assert_eq!(d(&result.selection_fee_incl), round(d(fee) * d("1.06")));
        }
    }
    #[test]
    fn rounding_order_and_inclusive_example() {
        let r = calculate("300000", "0").unwrap();
        assert_eq!(r.quote_excl, "283018.8679");
        assert_eq!(r.selection_fee_excl, "2292.41");
        assert_eq!(r.selection_fee_incl, "2429.95");
        assert_eq!(r.final_limit, "302429.95");
        assert_eq!(calculate("300000", "50").unwrap().final_limit, "302479.95");
        assert_eq!(calculate("106000", "0").unwrap().final_limit, "106424.00");
        // Task-book's 106858.59 omitted the one cent in quote 106000.01.
        assert_eq!(
            calculate("106000.01", "0").unwrap().final_limit,
            "106858.60"
        );
    }
    #[test]
    fn gaps_are_errors_not_approximate_quotes() {
        for limit in ["106424.01", "106600", "106858.59"] {
            assert!(reverse(limit, "0").unwrap_err().contains("10 万元档位跳变"));
        }
        assert!(reverse("106650", "50")
            .unwrap_err()
            .contains("10 万元档位跳变"));
        assert!(reverse("0.01", "0").is_err());
        assert!(reverse("20", "50").is_err());
        assert_eq!(reverse("0", "0").unwrap().quote, "0.00");
    }
    #[test]
    fn inverse_is_exact_and_discloses_non_unique_quotes() {
        let mut overlap_seen = false;
        // Exhaustive cent neighbourhoods include small downward tariff jumps and penny holes.
        for center in [0, 1282600, 5141000, 10600000, 106000000, 30000000] {
            for cents in (center - 40).max(0)..=center + 40 {
                let quote = Decimal::from(cents) * CENT;
                let forward = calculate(&quote.to_string(), "50").unwrap();
                let reverse_result = reverse(&forward.final_limit, "50").unwrap();
                let again = calculate(&reverse_result.quote, "50").unwrap();
                assert_eq!(forward.final_limit, again.final_limit);
                if !quote.is_zero() {
                    assert!(reverse_result
                        .quote_candidates
                        .contains(&format!("{quote:.2}")));
                    if reverse_result.quote_candidates.len() == 1 {
                        assert_eq!(forward.quote, reverse_result.quote);
                    } else {
                        overlap_seen = true;
                    }
                }
            }
        }
        assert!(
            overlap_seen,
            "downward tariff jumps produce multiple exact solutions"
        );
        for cents in (1..200000000i64).step_by(7919) {
            let quote = Decimal::from(cents) * CENT;
            let f = calculate(&quote.to_string(), "-50").unwrap();
            let r = reverse(&f.final_limit, "-50").unwrap();
            assert!(r.quote_candidates.contains(&f.quote));
        }
    }
    #[test]
    fn tax_drives_all_four_boundaries() {
        // Injectable tax exercises the exact production path without a second implementation.
        for (tax, expected) in [
            (6, [1282600, 5141000, 10600000, 106000000]),
            (9, [1318900, 5286500, 10900000, 109000000]),
        ] {
            let tariff = bands();
            for (i, bound) in expected.into_iter().enumerate() {
                assert_eq!(
                    tariff[i].upper.unwrap() * factor(tax) / CENT,
                    Decimal::from(bound)
                );
                for cents in [bound - 1, bound, bound + 1] {
                    let q = Decimal::from(cents) * CENT;
                    let f = calculate_at(q, d("0"), tax).unwrap();
                    let r = reverse_at(d(&f.final_limit), d("0"), tax).unwrap();
                    assert!(r.quote_candidates.contains(&f.quote));
                }
            }
        }
    }
    #[test]
    fn malformed_inputs_do_not_become_zero() {
        for quote in [
            "",
            "abc",
            "NaN",
            "inf",
            "-1",
            "1.001",
            "79228162514264337593543950335",
        ] {
            assert!(calculate(quote, "0").is_err(), "{quote}");
        }
        assert!(calculate("1", "invalid").is_err());
        assert!(calculate("1", "-1000").is_err());
        assert!(reverse("NaN", "0").is_err());
    }
}
