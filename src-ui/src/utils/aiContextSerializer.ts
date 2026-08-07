/**
 * AI Context Serializer
 * Converts JSON business state into semantic Markdown for LLM consumption.
 */

export function serializeAiContext(module: string, data: any): string {
  if (!data) return "当前无可用业务数据。";

  let markdown = `### 当前模块: ${module.toUpperCase()}\n`;

  if (module === 'ict') {
    markdown += serializeIctModule(data);
  } else if (module.startsWith('ict.template.')) {
    markdown += serializeTemplateData(data);
  } else {
    // Generic fallback for other modules in Phase 1
    markdown += serializeGenericData(data);
  }

  return markdown;
}


function serializeTemplateData(data: any): string {
  const lines: string[] = [];
  
  // Templates often have project_name at the top level
  if (data.basic?.proj_name) lines.push(`- **当前填报项目**: ${data.basic.proj_name}`);

  lines.push(`\n**[表单填报内容]**`);

  // Recursively or flatly extract non-empty values from the template payload
  // Template data is usually a flat object of fieldKey: value.
  Object.entries(data).forEach(([key, value]) => {
    // Skip internal structures like 'basic', 'revenue', 'cost' as they are handled by ict module
    if (['basic', 'revenue', 'cost', 'metrics', 'techItems', 'inqVendors'].includes(key)) return;
    
    if (value && typeof value === 'string' && value.trim() !== '') {
      // Improve label mapping: Strip common prefixes used in this project
      let label = key
        .replace(/^gen_/, '')
        .replace(/^demand_/, '')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, l => l.toUpperCase());
      
      // Known project specific mappings for better semantics
      const labelMap: Record<string, string> = {
        'Customer Confirm': '客户确认',
        'Proj Name': '项目名称',
        'Project Name': '项目名称',
        'Construction Interface': '施工界面',
        'Env Require': '环境要求',
        'Service Content': '服务内容',
        'Subject It Cost': 'IT计费科目',
        'Subject Ct Cost': 'CT计费科目'
      };

      if (labelMap[label]) label = labelMap[label];
      
      // If long text, use a section header
      if (value.length > 50) {
        lines.push(`\n#### ${label}\n${value}\n`);
      } else {
        lines.push(`- ${label}: ${value}`);
      }
    }
  });

  return lines.join('\n');
}

function serializeIctModule(data: any): string {
  const lines: string[] = [];

  // Project Overview
  if (data.project_name) lines.push(`- **项目名称**: ${data.project_name}`);
  if (data.customer_name) lines.push(`- **客户名称**: ${data.customer_name}`);
  if (data.project_background) lines.push(`- **项目背景**: ${data.project_background}`);

  // Summary Metrics
  if (data.metrics) {
    const m = data.metrics;
    lines.push(`\n**[核心财务指标]**`);
    if (m.npv !== undefined) lines.push(`- 净现值 (NPV): ¥${m.npv.toLocaleString()}`);
    if (m.margin_rate !== undefined) lines.push(`- 毛利润率: ${(m.margin_rate * 100).toFixed(2)}%`);
    if (m.npv_rate !== undefined) lines.push(`- 净现值率: ${(m.npv_rate * 100).toFixed(2)}%`);
    if (m.dynamic_payback) lines.push(`- 动态回收期: ${m.dynamic_payback}年`);
  }

  // Cost Breakdown (Only if values > 0)
  lines.push(`\n**[支出投入明细]**`);
  const costMap: Record<string, string> = {
    'cost_it_device': 'IT主要设备/材料',
    'cost_it_construction': 'IT施工费',
    'cost_it_integration': 'IT集成服务费',
    'cost_it_maintenance': 'IT维护费',
    'cost_ct_construction': 'CT专线建设',
    'cost_ct_product': 'CT其他产品成本',
    'cost_mix_marketing': '融合营销成本',
    'cost_it_bidding': '中标服务费'
  };

  Object.entries(costMap).forEach(([key, label]) => {
    const val = data[key]?.incl_tax || data[key]?.incl;
    if (val && Number(val) > 0) {
      lines.push(`- ${label}: ¥${Number(val).toLocaleString()}`);
    }
  });

  // Revenue Breakdown
  lines.push(`\n**[收入组成]**`);
  const revMap: Record<string, string> = {
    'rev_it_integration': '系统集成收入',
    'rev_it_cloud': '移动云定制收入',
    'rev_ct_line': '专线收入',
    'rev_ct_product': 'CT产品收入'
  };

  Object.entries(revMap).forEach(([key, label]) => {
    const val = data[key]?.incl_tax || data[key]?.incl;
    if (val && Number(val) > 0) {
      lines.push(`- ${label}: ¥${Number(val).toLocaleString()}`);
    }
  });

  return lines.join('\n');
}

function serializeGenericData(data: any): string {
  // Simple key-value list for unknown modules
  return Object.entries(data)
    .filter(([_, v]) => v !== null && v !== undefined && v !== 0 && v !== '')
    .map(([k, v]) => {
      const val = typeof v === 'object' ? JSON.stringify(v) : v;
      return `- ${k}: ${val}`;
    })
    .join('\n');
}
