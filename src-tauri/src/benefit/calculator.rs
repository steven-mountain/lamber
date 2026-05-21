use super::models::{
    CalcInput, CalcResult, IctCashflowRow, IctInput, IctItem, IctResult, SelectionFeeResult,
};
use rust_decimal::prelude::*;
use std::str::FromStr;

fn round_2(val: Decimal) -> Decimal {
    val.round_dp(2)
}

fn round_4(val: Decimal) -> Decimal {
    val.round_dp(4)
}

fn get_excl(item: &IctItem) -> Decimal {
    let incl = Decimal::from_str(&item.incl_tax).unwrap_or(Decimal::ZERO);
    let rate = Decimal::from_str(&item.tax_rate).unwrap_or(Decimal::ZERO) / Decimal::new(100, 0);
    if incl.is_zero() {
        return Decimal::ZERO;
    }
    (incl / (Decimal::ONE + rate)).round_dp(2)
}

fn normalize_distribution(input: &[f64]) -> Vec<Decimal> {
    let default = vec![
        Decimal::ONE,
        Decimal::ZERO,
        Decimal::ZERO,
        Decimal::ZERO,
        Decimal::ZERO,
        Decimal::ZERO,
        Decimal::ZERO,
        Decimal::ZERO,
        Decimal::ZERO,
        Decimal::ZERO,
    ];

    if input.len() != 10 {
        return default;
    }

    let mut dist: Vec<Decimal> = input
        .iter()
        .map(|v| Decimal::from_f64_retain(*v).unwrap_or(Decimal::ZERO))
        .collect();

    for item in dist.iter_mut() {
        if *item < Decimal::ZERO {
            *item = Decimal::ZERO;
        }
    }

    let sum: Decimal = dist.iter().copied().sum();

    if sum <= Decimal::ZERO {
        return default;
    }

    dist.iter().map(|v| (*v / sum).round_dp(8)).collect()
}

fn parse_cashflow_override(input: &Option<Vec<String>>) -> Option<Vec<Decimal>> {
    let values = input.as_ref()?;
    let mut result = vec![Decimal::ZERO; 10];

    for (idx, value) in values.iter().take(10).enumerate() {
        let parsed = Decimal::from_str(value).unwrap_or(Decimal::ZERO);
        result[idx] = if parsed > Decimal::ZERO {
            parsed.round_dp(2)
        } else {
            Decimal::ZERO
        };
    }

    let total: Decimal = result.iter().copied().sum();
    if total > Decimal::ZERO {
        Some(result)
    } else {
        None
    }
}

#[tauri::command]
pub fn calculate_ict_benefit(input: IctInput) -> Result<IctResult, String> {
    let discount_rate = Decimal::from_str(&input.discount_rate).unwrap_or(Decimal::new(55, 3)); // 0.055

    // IT Revenue
    let it_rev = get_excl(&input.rev_it_integration)
        + get_excl(&input.rev_it_maintenance)
        + get_excl(&input.rev_it_device_sales)
        + get_excl(&input.rev_it_device_lease)
        + get_excl(&input.rev_it_other)
        + get_excl(&input.rev_it_cloud);

    // CT Revenue
    let ct_rev = get_excl(&input.rev_ct_line) + get_excl(&input.rev_ct_product);

    // Non-IT Revenue
    let non_it_rev = get_excl(&input.rev_non_it_ct);

    let total_rev = it_rev + ct_rev + non_it_rev;

    // IT Cost
    let it_cost = get_excl(&input.cost_it_device)
        + get_excl(&input.cost_it_construction)
        + get_excl(&input.cost_it_survey)
        + get_excl(&input.cost_it_integration)
        + get_excl(&input.cost_it_other)
        + get_excl(&input.cost_it_maintenance)
        + get_excl(&input.cost_it_running)
        + get_excl(&input.cost_it_bidding)
        + get_excl(&input.cost_it_design_eval)
        + get_excl(&input.cost_it_audit);

    // CT Cost
    let ct_cost = get_excl(&input.cost_ct_construction)
        + get_excl(&input.cost_ct_maintenance)
        + get_excl(&input.cost_ct_other)
        + get_excl(&input.cost_ct_bandwidth)
        + get_excl(&input.cost_ct_renewal);

    // Non-IT Cost & Mix Cost
    let non_it_cost = get_excl(&input.cost_non_it_ct);
    let mix_cost = get_excl(&input.cost_mix_marketing)
        + get_excl(&input.cost_mix_channel)
        + get_excl(&input.cost_mix_other);

    let total_cost = it_cost + ct_cost + non_it_cost + mix_cost;

    let margin_rate = if total_rev.is_zero() {
        Decimal::ZERO
    } else {
        ((total_rev - total_cost) / total_rev).round_dp(4)
    };
    let it_margin_rate = if it_rev.is_zero() {
        Decimal::ZERO
    } else {
        ((it_rev - it_cost) / it_rev).round_dp(4)
    };

    // --- 10 Year Cashflow Simulation ---
    let mut cashflow = Vec::new();

    let mut cum_net_cash = Decimal::ZERO;
    let mut cum_pv = Decimal::ZERO;

    let mut total_pv_in = Decimal::ZERO;
    let mut total_pv_out = Decimal::ZERO;
    let mut total_it_pv_in = Decimal::ZERO;
    let mut total_it_pv_out = Decimal::ZERO;

    let mut dynamic_payback_year = 0;
    let mut payback_found = false;

    let rev_dist = normalize_distribution(&input.rev_distribution);
    let cost_dist = normalize_distribution(&input.cost_distribution);
    let direct_rev_cashflow = parse_cashflow_override(&input.rev_cashflow_excl);
    let direct_cost_cashflow = parse_cashflow_override(&input.cost_cashflow_excl);
    let direct_it_rev_cashflow = parse_cashflow_override(&input.it_rev_cashflow_excl);
    let direct_it_cost_cashflow = parse_cashflow_override(&input.it_cost_cashflow_excl);

    for year in 1..=10 {
        let rev_ratio = rev_dist[year - 1];
        let cost_ratio = cost_dist[year - 1];

        let cash_in = direct_rev_cashflow
            .as_ref()
            .map(|values| values[year - 1])
            .unwrap_or_else(|| (total_rev * rev_ratio).round_dp(2));
        let cash_out = direct_cost_cashflow
            .as_ref()
            .map(|values| values[year - 1])
            .unwrap_or_else(|| (total_cost * cost_ratio).round_dp(2));
        let net_cash = cash_in - cash_out;

        // IT specific breakdown
        let it_cash_in = direct_it_rev_cashflow
            .as_ref()
            .map(|values| values[year - 1])
            .unwrap_or_else(|| (it_rev * rev_ratio).round_dp(2));
        let it_cash_out = direct_it_cost_cashflow
            .as_ref()
            .map(|values| values[year - 1])
            .unwrap_or_else(|| (it_cost * cost_ratio).round_dp(2));

        // Discount factor for the current year: (1 + discount_rate)^year
        let mut pv_factor = Decimal::ONE;
        for _ in 0..year {
            pv_factor *= Decimal::ONE + discount_rate;
        }

        let pv_in = (cash_in / pv_factor).round_dp(2);
        let pv_out = (cash_out / pv_factor).round_dp(2);
        let pv_net = pv_in - pv_out;

        let it_pv_in = (it_cash_in / pv_factor).round_dp(2);
        let it_pv_out = (it_cash_out / pv_factor).round_dp(2);

        total_pv_in += pv_in;
        total_pv_out += pv_out;

        total_it_pv_in += it_pv_in;
        total_it_pv_out += it_pv_out;

        cum_net_cash += net_cash;
        cum_pv += pv_net;

        // Determine dynamic payback period
        if !payback_found && cum_pv >= Decimal::ZERO {
            dynamic_payback_year = year;
            payback_found = true;
        }

        cashflow.push(IctCashflowRow {
            year: year as i32,
            cash_in: cash_in.to_string(),
            cash_out: cash_out.to_string(),
            net_cash: net_cash.to_string(),
            cum_net_cash: cum_net_cash.to_string(),
            pv: pv_net.to_string(),
            cum_pv: cum_pv.to_string(),
        });
    }

    let npv = total_pv_in - total_pv_out;
    let npv_rate = if total_pv_out.is_zero() {
        Decimal::ZERO
    } else {
        (npv / total_pv_out).round_dp(4)
    };

    let it_npv = total_it_pv_in - total_it_pv_out;
    let it_npv_rate = if total_it_pv_out.is_zero() {
        Decimal::ZERO
    } else {
        (it_npv / total_it_pv_out).round_dp(4)
    };

    let dynamic_payback_str = if payback_found {
        dynamic_payback_year.to_string()
    } else {
        ">10".to_string()
    };

    Ok(IctResult {
        npv: npv.to_string(),
        npv_rate: npv_rate.to_string(),
        margin_rate: margin_rate.to_string(),
        dynamic_payback: dynamic_payback_str,
        irr: "--".to_string(),
        it_npv: it_npv.to_string(),
        it_npv_rate: it_npv_rate.to_string(),
        it_margin_rate: it_margin_rate.to_string(),
        cashflow,
    })
}

#[tauri::command]
pub fn reverse_calc_ict_target(
    input: IctInput,
    target_type: String,
    target_value: String,
) -> Result<String, String> {
    let target = Decimal::from_str(&target_value).unwrap_or(Decimal::ZERO);
    let mut low = Decimal::ZERO;
    let mut high = Decimal::new(10_000_000_000, 0); // 10 billion limit
    let mut best_mid = Decimal::ZERO;

    for _ in 0..100 {
        let mid = (low + high) / Decimal::new(2, 0);
        best_mid = mid;

        let mut test_input = input.clone();
        test_input.cost_it_integration.incl_tax = mid.to_string();

        let res = calculate_ict_benefit(test_input)?;

        let current_val = if target_type == "margin" {
            Decimal::from_str(&res.margin_rate).unwrap_or(Decimal::ZERO)
        } else {
            Decimal::from_str(&res.npv_rate).unwrap_or(Decimal::ZERO)
        };

        // As cost increases, margin_rate and npv_rate both decrease.
        if current_val > target {
            // We have a higher metric than target -> need to decrease metric -> need to INCREASE cost
            low = mid;
        } else {
            // We have a lower metric than target -> need to increase metric -> need to DECREASE cost
            high = mid;
        }
    }

    Ok(best_mid.round_dp(2).to_string())
}

#[tauri::command]
pub fn reverse_calc_ict_revenue_target(
    input: IctInput,
    target_type: String,
    target_value: String,
) -> Result<String, String> {
    let target = Decimal::from_str(&target_value).unwrap_or(Decimal::ZERO);
    let mut low = Decimal::ZERO;
    let mut high = Decimal::new(10_000_000_000, 0);
    let mut best_mid = Decimal::ZERO;

    for _ in 0..100 {
        let mid = (low + high) / Decimal::new(2, 0);
        best_mid = mid;

        let mut test_input = input.clone();
        test_input.rev_it_integration.incl_tax = mid.to_string();

        let res = calculate_ict_benefit(test_input)?;

        let current_val = if target_type == "margin" {
            Decimal::from_str(&res.margin_rate).unwrap_or(Decimal::ZERO)
        } else {
            Decimal::from_str(&res.npv_rate).unwrap_or(Decimal::ZERO)
        };

        if current_val < target {
            low = mid;
        } else {
            high = mid;
        }
    }

    Ok(best_mid.round_dp(2).to_string())
}

#[tauri::command]
pub fn calculate_benefit(input: CalcInput) -> Result<CalcResult, String> {
    let d1 = Decimal::ONE;
    let d72 = Decimal::new(72, 0);
    let d0_01 = Decimal::new(1, 2);

    let tax_rate_it = Decimal::from_str(&input.tax_rate_it).unwrap_or(Decimal::new(6, 2));
    let tax_rate_ct = Decimal::from_str(&input.tax_rate_ct).unwrap_or(Decimal::new(6, 2));
    let total_income_incl =
        Decimal::from_str(&input.total_income_incl).map_err(|e| e.to_string())?;
    let target_value = Decimal::from_str(&input.target_value).map_err(|e| e.to_string())?;

    // --- 第一步：含税盘子分配 ---
    let ct_income_incl = if let Some(ct_str) = &input.ct_income_incl_opt {
        if ct_str.trim().is_empty() {
            let ct_income_incl_min = round_2(total_income_incl * d0_01);
            let ceil_multiplier = (ct_income_incl_min / d72).ceil();
            round_2(ceil_multiplier * d72)
        } else {
            Decimal::from_str(ct_str).unwrap_or_else(|_| {
                let ct_income_incl_min = round_2(total_income_incl * d0_01);
                let ceil_multiplier = (ct_income_incl_min / d72).ceil();
                round_2(ceil_multiplier * d72)
            })
        }
    } else {
        let ct_income_incl_min = round_2(total_income_incl * d0_01);
        let ceil_multiplier = (ct_income_incl_min / d72).ceil();
        round_2(ceil_multiplier * d72)
    };

    let ct_cost_incl = ct_income_incl;
    let it_income_incl = round_2(total_income_incl - ct_income_incl);

    // --- 第二步：价税分离 ---
    let it_income_excl = round_2(it_income_incl / (d1 + tax_rate_it));
    let ct_income_excl = round_2(ct_income_incl / (d1 + tax_rate_ct));
    let total_income_excl = it_income_excl + ct_income_excl;
    let ct_cost_excl = round_2(ct_cost_incl / (d1 + tax_rate_ct));

    // --- 第三步：测算投入 ---
    let total_cost_excl;
    let it_cost_excl;
    let it_cost_incl;
    let total_cost_incl;

    if input.calc_mode == "margin" {
        total_cost_excl = round_2(total_income_excl * (d1 - target_value));
        it_cost_excl = total_cost_excl - ct_cost_excl;
        it_cost_incl = round_2(it_cost_excl * (d1 + tax_rate_it));
        total_cost_incl = it_cost_incl + ct_cost_incl;
    } else if input.calc_mode == "npv" {
        total_cost_excl = round_2(total_income_excl / (d1 + target_value));
        it_cost_excl = total_cost_excl - ct_cost_excl;
        it_cost_incl = round_2(it_cost_excl * (d1 + tax_rate_it));
        total_cost_incl = it_cost_incl + ct_cost_incl;
    } else if input.calc_mode == "total_cost" {
        total_cost_incl = target_value;
        it_cost_incl = total_cost_incl - ct_cost_incl;
        it_cost_excl = round_2(it_cost_incl / (d1 + tax_rate_it));
        total_cost_excl = it_cost_excl + ct_cost_excl;
    } else {
        return Err("未知的计算模式".to_string());
    }

    // --- 第四步：效益指标核算 ---
    let margin_rate = if total_income_excl.is_zero() {
        Decimal::ZERO
    } else {
        round_4((total_income_excl - total_cost_excl) / total_income_excl)
    };

    let npv_rate = if total_cost_excl.is_zero() {
        Decimal::ZERO
    } else {
        round_4((total_income_excl - total_cost_excl) / total_cost_excl)
    };

    let it_npv_rate = if it_cost_excl.is_zero() {
        Decimal::ZERO
    } else {
        round_4((it_income_excl - it_cost_excl) / it_cost_excl)
    };

    let mut warnings = Vec::new();
    if it_cost_incl < Decimal::ZERO || it_cost_excl < Decimal::ZERO {
        warnings.push("目标太高或投入太低，IT投入已被穿透为负数".to_string());
    }

    let warning_message = if warnings.is_empty() {
        None
    } else {
        Some(warnings.join(" | "))
    };

    Ok(CalcResult {
        it_income_incl: it_income_incl.to_string(),
        ct_income_incl: ct_income_incl.to_string(),
        total_income_incl: total_income_incl.to_string(),
        it_income_excl: it_income_excl.to_string(),
        ct_income_excl: ct_income_excl.to_string(),
        total_income_excl: total_income_excl.to_string(),
        it_cost_incl: it_cost_incl.to_string(),
        ct_cost_incl: ct_cost_incl.to_string(),
        total_cost_incl: total_cost_incl.to_string(),
        it_cost_excl: it_cost_excl.to_string(),
        ct_cost_excl: ct_cost_excl.to_string(),
        total_cost_excl: total_cost_excl.to_string(),
        margin_rate: margin_rate.to_string(),
        npv_rate: npv_rate.to_string(),
        it_npv_rate: it_npv_rate.to_string(),
        warning_message,
    })
}

#[tauri::command]
pub fn calculate_selection_fee(
    quote: String,
    markup: String,
) -> Result<SelectionFeeResult, String> {
    let quote_val = quote.parse::<f64>().unwrap_or(0.0);
    let markup_val = markup.parse::<f64>().unwrap_or(0.0);

    let mut selection_fee = 0.0;
    if quote_val <= 0.0 {
        selection_fee = 0.0;
    } else if quote_val <= 12100.0 {
        selection_fee = 100.0;
    } else if quote_val <= 48500.0 {
        selection_fee = (quote_val * 0.00825 * 100.0).round() / 100.0;
    } else if quote_val <= 100000.0 {
        selection_fee = 400.0;
    } else if quote_val <= 500000.0 {
        selection_fee = (quote_val * 0.009408 * 100.0).round() / 100.0;
    } else {
        selection_fee = 0.0;
    }

    let actual_cost = quote_val + selection_fee;
    let final_limit = actual_cost + markup_val;

    Ok(SelectionFeeResult {
        selection_fee: format!("{:.2}", selection_fee),
        actual_cost: format!("{:.2}", actual_cost),
        final_limit: format!("{:.2}", final_limit),
        quote: format!("{:.2}", quote_val),
    })
}

#[tauri::command]
pub fn reverse_calculate_selection_fee(
    limit: String,
    markup: String,
) -> Result<SelectionFeeResult, String> {
    let limit_val = limit.parse::<f64>().unwrap_or(0.0);
    let markup_val = markup.parse::<f64>().unwrap_or(0.0);

    let actual_cost_target = limit_val - markup_val;

    if actual_cost_target <= 0.0 {
        return Ok(SelectionFeeResult {
            selection_fee: "0.00".to_string(),
            actual_cost: "0.00".to_string(),
            final_limit: format!("{:.2}", limit_val),
            quote: "0.00".to_string(),
        });
    }

    let mut quote = 0.0;
    let mut selection_fee = 0.0;

    if actual_cost_target <= 12200.0 {
        // 12100 + 100
        selection_fee = 100.0;
        quote = actual_cost_target - 100.0;
    } else if actual_cost_target <= 48900.12 {
        // 48500 + 48500*0.00825(400.125)
        quote = actual_cost_target / 1.00825;
        selection_fee = actual_cost_target - quote;
    } else if actual_cost_target <= 100400.0 {
        // 100000 + 400
        selection_fee = 400.0;
        quote = actual_cost_target - 400.0;
    } else if actual_cost_target <= 504704.0 {
        // 500000 + 500000*0.009408(4704)
        quote = actual_cost_target / 1.009408;
        selection_fee = actual_cost_target - quote;
    } else {
        quote = actual_cost_target;
        selection_fee = 0.0;
    }

    Ok(SelectionFeeResult {
        selection_fee: format!("{:.2}", selection_fee),
        actual_cost: format!("{:.2}", actual_cost_target),
        final_limit: format!("{:.2}", limit_val),
        quote: format!("{:.2}", quote),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn item(incl_tax: &str) -> IctItem {
        IctItem {
            incl_tax: incl_tax.to_string(),
            tax_rate: "0".to_string(),
        }
    }

    fn item_with_tax(incl_tax: &str, tax_rate: &str) -> IctItem {
        IctItem {
            incl_tax: incl_tax.to_string(),
            tax_rate: tax_rate.to_string(),
        }
    }

    fn dist(values: &[f64]) -> Vec<f64> {
        let mut result = vec![0.0; 10];
        for (idx, value) in values.iter().enumerate().take(10) {
            result[idx] = *value;
        }
        result
    }

    fn decimal(value: &str) -> Decimal {
        Decimal::from_str(value).unwrap()
    }

    fn input_with(
        revenue: &str,
        cost: &str,
        rev_distribution: Vec<f64>,
        cost_distribution: Vec<f64>,
    ) -> IctInput {
        let zero = item("0");

        IctInput {
            project_name: "test".to_string(),
            customer_name: None,
            property_rights: "customer".to_string(),
            discount_rate: "0.055".to_string(),
            project_years: None,
            cashflow_model: None,
            cashflow_segment_value_mode: None,
            cashflow_segments: None,
            ignore_tail_difference: None,
            tail_difference_value: None,
            rev_distribution,
            cost_distribution,
            rev_cashflow_excl: None,
            cost_cashflow_excl: None,
            it_rev_cashflow_excl: None,
            it_cost_cashflow_excl: None,
            rev_it_integration: item(revenue),
            rev_it_maintenance: zero.clone(),
            rev_it_device_sales: zero.clone(),
            rev_it_device_lease: zero.clone(),
            rev_it_other: zero.clone(),
            rev_it_cloud: zero.clone(),
            rev_ct_line: zero.clone(),
            rev_ct_product: zero.clone(),
            rev_non_it_ct: zero.clone(),
            cost_it_device: zero.clone(),
            cost_it_construction: zero.clone(),
            cost_it_survey: zero.clone(),
            cost_it_integration: item(cost),
            cost_it_other: zero.clone(),
            cost_it_maintenance: zero.clone(),
            cost_it_running: zero.clone(),
            cost_it_bidding: zero.clone(),
            cost_it_design_eval: zero.clone(),
            cost_it_audit: zero.clone(),
            cost_ct_construction: zero.clone(),
            cost_ct_maintenance: zero.clone(),
            cost_ct_other: zero.clone(),
            cost_ct_bandwidth: zero.clone(),
            cost_ct_renewal: zero.clone(),
            cost_non_it_ct: zero.clone(),
            cost_mix_marketing: zero.clone(),
            cost_mix_channel: zero.clone(),
            cost_mix_other: zero,
        }
    }

    #[test]
    fn normalizes_invalid_distribution_to_year_one() {
        let normalized = normalize_distribution(&vec![0.0; 10]);

        assert_eq!(normalized[0], Decimal::ONE);
        assert!(normalized
            .iter()
            .skip(1)
            .all(|value| *value == Decimal::ZERO));
    }

    #[test]
    fn normalizes_custom_distribution_and_clamps_negative_values() {
        let normalized = normalize_distribution(&dist(&[2.0, 1.0, -1.0]));

        assert_eq!(normalized[0], Decimal::from_str("0.66666667").unwrap());
        assert_eq!(normalized[1], Decimal::from_str("0.33333333").unwrap());
        assert_eq!(normalized[2], Decimal::ZERO);
    }

    #[test]
    fn model_a_cashflow_stays_in_first_year() {
        let result =
            calculate_ict_benefit(input_with("1000000", "600000", dist(&[1.0]), dist(&[1.0])))
                .unwrap();

        assert_eq!(
            decimal(&result.cashflow[0].cash_in),
            Decimal::from_str("1000000.00").unwrap()
        );
        assert_eq!(
            decimal(&result.cashflow[0].cash_out),
            Decimal::from_str("600000.00").unwrap()
        );
        assert!(result.cashflow.iter().skip(1).all(|row| {
            decimal(&row.cash_in) == Decimal::ZERO && decimal(&row.cash_out) == Decimal::ZERO
        }));
        assert_eq!(
            decimal(&result.npv),
            Decimal::from_str("379146.92").unwrap()
        );
    }

    #[test]
    fn model_b_even_distribution_changes_cashflow_and_lowers_npv() {
        let model_a =
            calculate_ict_benefit(input_with("1000000", "600000", dist(&[1.0]), dist(&[1.0])))
                .unwrap();
        let model_b = calculate_ict_benefit(input_with(
            "1000000",
            "600000",
            dist(&[1.0, 1.0, 1.0]),
            dist(&[1.0, 1.0, 1.0]),
        ))
        .unwrap();

        for row in model_b.cashflow.iter().take(3) {
            assert_eq!(
                decimal(&row.cash_in),
                Decimal::from_str("333333.33").unwrap()
            );
            assert_eq!(
                decimal(&row.cash_out),
                Decimal::from_str("200000.00").unwrap()
            );
        }
        assert!(model_b.cashflow.iter().skip(3).all(|row| {
            decimal(&row.cash_in) == Decimal::ZERO && decimal(&row.cash_out) == Decimal::ZERO
        }));
        assert!(decimal(&model_b.npv) < decimal(&model_a.npv));
    }

    #[test]
    fn model_c_splits_first_year_and_final_year() {
        let result = calculate_ict_benefit(input_with(
            "1000000",
            "600000",
            dist(&[0.95, 0.0, 0.05]),
            dist(&[0.95, 0.0, 0.05]),
        ))
        .unwrap();

        assert_eq!(
            decimal(&result.cashflow[0].cash_in),
            Decimal::from_str("950000.00").unwrap()
        );
        assert_eq!(
            decimal(&result.cashflow[0].cash_out),
            Decimal::from_str("570000.00").unwrap()
        );
        assert_eq!(decimal(&result.cashflow[1].cash_in), Decimal::ZERO);
        assert_eq!(
            decimal(&result.cashflow[2].cash_in),
            Decimal::from_str("50000.00").unwrap()
        );
        assert_eq!(
            decimal(&result.cashflow[2].cash_out),
            Decimal::from_str("30000.00").unwrap()
        );
    }

    #[test]
    fn reverse_revenue_uses_distribution_cashflow_for_npv_rate() {
        let year_one_revenue = reverse_calc_ict_revenue_target(
            input_with("0", "600000", dist(&[1.0]), dist(&[1.0])),
            "npv_rate".to_string(),
            "0.15".to_string(),
        )
        .unwrap();
        let delayed_revenue = reverse_calc_ict_revenue_target(
            input_with("0", "600000", dist(&[1.0, 1.0, 1.0]), dist(&[1.0])),
            "npv_rate".to_string(),
            "0.15".to_string(),
        )
        .unwrap();

        assert!(decimal(&delayed_revenue) > decimal(&year_one_revenue));
    }

    #[test]
    fn direct_cashflow_override_preserves_segment_tax_and_custom_schedule() {
        let zero = item("0");
        let mut input = IctInput {
            project_name: "test".to_string(),
            customer_name: None,
            property_rights: "customer".to_string(),
            discount_rate: "0.055".to_string(),
            project_years: None,
            cashflow_model: None,
            cashflow_segment_value_mode: None,
            cashflow_segments: None,
            ignore_tail_difference: None,
            tail_difference_value: None,
            rev_distribution: dist(&[0.5, 0.5]),
            cost_distribution: dist(&[0.5, 0.5]),
            rev_cashflow_excl: Some(vec![
                "42738.71".to_string(),
                "42275.23".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
            ]),
            cost_cashflow_excl: Some(vec![
                "35933.34".to_string(),
                "42275.23".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
                "0".to_string(),
            ]),
            it_rev_cashflow_excl: None,
            it_cost_cashflow_excl: None,
            rev_it_integration: item_with_tax("17850", "6"),
            rev_it_maintenance: zero.clone(),
            rev_it_device_sales: zero.clone(),
            rev_it_device_lease: zero.clone(),
            rev_it_other: zero.clone(),
            rev_it_cloud: zero.clone(),
            rev_ct_line: zero.clone(),
            rev_ct_product: item_with_tax("74310", "9"),
            rev_non_it_ct: zero.clone(),
            cost_it_device: zero.clone(),
            cost_it_construction: zero.clone(),
            cost_it_survey: zero.clone(),
            cost_it_integration: item_with_tax("10636.31", "6"),
            cost_it_other: zero.clone(),
            cost_it_maintenance: zero.clone(),
            cost_it_running: zero.clone(),
            cost_it_bidding: zero.clone(),
            cost_it_design_eval: zero.clone(),
            cost_it_audit: zero.clone(),
            cost_ct_construction: zero.clone(),
            cost_ct_maintenance: zero.clone(),
            cost_ct_other: item_with_tax("74310", "9"),
            cost_ct_bandwidth: zero.clone(),
            cost_ct_renewal: zero.clone(),
            cost_non_it_ct: zero.clone(),
            cost_mix_marketing: zero.clone(),
            cost_mix_channel: zero.clone(),
            cost_mix_other: zero,
        };

        let result = calculate_ict_benefit(input.clone()).unwrap();

        assert_eq!(
            decimal(&result.cashflow[0].cash_in),
            Decimal::from_str("42738.71").unwrap()
        );
        assert_eq!(
            decimal(&result.cashflow[1].cash_in),
            Decimal::from_str("42275.23").unwrap()
        );
        assert_eq!(
            decimal(&result.cashflow[0].cash_out),
            Decimal::from_str("35933.34").unwrap()
        );
        assert_eq!(
            decimal(&result.cashflow[1].cash_out),
            Decimal::from_str("42275.23").unwrap()
        );

        input.rev_cashflow_excl = None;
        input.cost_cashflow_excl = None;
        let distributed = calculate_ict_benefit(input).unwrap();

        assert_eq!(
            decimal(&distributed.cashflow[0].cash_in),
            Decimal::from_str("42506.96").unwrap()
        );
        assert_eq!(
            decimal(&distributed.cashflow[0].cash_out),
            Decimal::from_str("39104.28").unwrap()
        );
    }
}
