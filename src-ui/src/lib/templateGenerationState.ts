import { templateCatalog } from './templateCompletion/catalog';
/** Defaults shared by rendering, completion and document input. Explicit empty values remain empty. */
export const TEMPLATE_FIELD_DEFAULTS: Readonly<Record<string, string>> = {
  ...Object.fromEntries(templateCatalog.flatMap(template => template.fields).filter(field => 'defaultValue' in field).map(field => [field.key, field.defaultValue])),
  "gen_ppt_report_unit": "沙坪坝分公司AI云数中心",
  "gen_meet_mode": "线上",
  "gen_branch_name": "XXXX",
  "gen_tech_solution": "采用端-管-云架构...",
  "gen_threeization": "本项目不涉及三化方案。",
  "gen_tech_conclusion": "方案可行同时能满足客户需求。",
  "gen_construction_time_req": "合同签定后30天内。",
  "gen_risk_owner": "人员A",
  "gen_review_acc": "是，项目投入收入核算完整，各表填写准确",
  "gen_single_source": "单一来源决策依据：符合单一来源场景...",
  "gen_construction_interface": "本项目采购统一集成单位实施。分公司负责客户侧的协调工作，并协调管理合作伙伴完成交付。",
  "gen_zx_scope": "三级库",
  "gen_zx_industry": "/",
  "gen_zx_method": "竞争性甄选",
  "gen_zx_rule": "标准方案",
  "gen_zx_std_plan": "竞价法",
  "gen_zx_is_sme": "否",
  "gen_is_joint": "否"
};

export function buildTemplateGenerationFields(
  saved: Record<string, string>,
  controlled: Record<string, string> = {},
  defaults: Record<string, string> = {},
): Record<string, string> {
  return { ...TEMPLATE_FIELD_DEFAULTS, ...defaults, ...saved, ...controlled };
}

/** Reject lost/overwritten saved values, including accidental fallback to a default. */
export function assertTemplateGenerationFields(
  expected: Record<string, string>, actual: Record<string, string>,
) {
  const lost = Object.keys(expected).filter(key => expected[key].trim() && actual[key] !== expected[key]);
  if (lost.length) throw new Error(`模板生成输入与已保存字段不一致，已阻止生成：${lost.join('、')}`);
}
