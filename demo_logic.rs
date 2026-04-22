use rust_decimal::prelude::*;
use rust_decimal_macros::dec;

fn main() {
    let total_income_incl = dec!(1000000);
    let tax_rate_it = dec!(0.06);
    let tax_rate_ct = dec!(0.06);
    let target_margin = dec!(0.15);

    println!("--- 项目效益测算演示 (CLI 模式) ---");
    println!("输入总收入 (含税): {}", total_income_incl);
    println!("IT 税率: {}, CT 税率: {}", tax_rate_it, tax_rate_ct);
    println!("目标毛利率: {}", target_margin);

    // 逻辑实现
    let d1 = Decimal::ONE;
    let d72 = dec!(72);
    let d0_01 = dec!(0.01);

    let ct_income_incl_min = (total_income_incl * d0_01).round_dp(2);
    let ceil_multiplier = (ct_income_incl_min / d72).ceil();
    let ct_income_incl = (ceil_multiplier * d72).round_dp(2);

    let it_income_incl = total_income_incl - ct_income_incl;

    let it_income_excl = (it_income_incl / (d1 + tax_rate_it)).round_dp(2);
    let ct_income_excl = (ct_income_incl / (d1 + tax_rate_ct)).round_dp(2);
    let total_income_excl = it_income_excl + ct_income_excl;

    // 模式 1.1：已知目标[毛利润率]反推投入
    let total_cost_excl = (total_income_excl * (d1 - target_margin)).round_dp(2);
    let ct_cost_excl = ct_income_excl; // 假设 CT 成本 = CT 不含税收入
    let it_cost_excl = total_cost_excl - ct_cost_excl;
    let it_cost_incl = (it_cost_excl * (d1 + tax_rate_it)).round_dp(2);
    let total_cost_incl = it_cost_incl + ct_income_incl;

    // 指标
    let margin_rate = (total_income_excl - total_cost_excl) / total_income_excl;
    let npv_rate = (total_income_excl - total_cost_excl) / total_cost_excl;

    println!("\n--- 测算结果 ---");
    println!("IT 部分 (含税收入): {}", it_income_incl);
    println!("CT 部分 (含税收入): {}", ct_income_incl);
    println!("总体不含税收入: {}", total_income_excl);
    println!("建议投入 (含税总额): {}", total_cost_incl);
    println!("IT 建议投入 (不含税): {}", it_cost_excl);
    println!("测算毛利率: {:.2}%", margin_rate * dec!(100));
    println!("项目净现值率 (NPV): {:.2}%", npv_rate * dec!(100));
}
