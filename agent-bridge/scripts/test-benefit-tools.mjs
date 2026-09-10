import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {readBenefitInputs,simulateBenefitCalculation,calculateSelectionFee,reverseCalculateSelectionFee,runBenefitCalculation,isGatedTool} from '../dsh-tool-lamber/lib/index.js';
const evidence=JSON.parse(readFileSync(new URL('../../docs/verification/ai-benefit-real-model-evidence.json',import.meta.url)));
const exec={agent:{session:{id:'trusted-session'}},signal:new AbortController().signal};
process.env.LAMBER_BRIDGE_URL='http://127.0.0.1:1';process.env.LAMBER_BRIDGE_TOKEN='synthetic';
const originalFetch=globalThis.fetch;
try {
 for(const [tool,suffix,args] of [[readBenefitInputs,'read-benefit-inputs',{}],[runBenefitCalculation,'calculate',{projectId:'bound'}],[simulateBenefitCalculation,'simulate-benefit-calculation',{overrides:[{subject:'rev_it_integration',inclTax:'800000'}]}],[calculateSelectionFee,'calculate-selection-fee',{quote:'300000',markup:'0'}],[reverseCalculateSelectionFee,'reverse-calculate-selection-fee',{limit:'51834',markup:'0'}]]){
  assert.equal(isGatedTool(tool.name),false);await assert.rejects(tool.execute(args,{signal:exec.signal}),/可信/);
  const response=evidence.calls.find(c=>c.route.endsWith('/'+suffix)&&c.status===200)?.response;assert(response,suffix);
  globalThis.fetch=async(_url,options)=>{const body=JSON.parse(options.body);assert.equal(body.sessionId,'trusted-session');if(tool!==runBenefitCalculation)assert.equal(body.injected,'keep-and-reject-on-server');return new Response(JSON.stringify(response),{status:200});};
  const value=await tool.execute({...args,sessionId:'spoof',injected:'keep-and-reject-on-server'},exec);
  const rendered=tool.output.render(args,value).map(v=>v.text).join('\n');
  if(tool===simulateBenefitCalculation){assert.match(rendered,/假设试算 · 未保存/);assert.match(rendered,/显式覆盖/);assert.match(rendered,/联动变更/);}
  if(tool===simulateBenefitCalculation)for(const change of value.linkedChanges){if(change.kind==='annual_cashflow'){const before=JSON.parse(change.before),after=JSON.parse(change.after);for(let i=0;i<before.length;i++)assert(rendered.includes(`第 ${i+1} 年 | ${before[i]} | ${after[i]}`));}else{assert(rendered.includes(change.before));assert(rendered.includes(change.after));}}
  if(tool===runBenefitCalculation)assert.match(rendered,/已保存快照重算/);
  if(tool===simulateBenefitCalculation||tool===runBenefitCalculation){
   const result=tool===simulateBenefitCalculation?value.result:value;
   for(const row of result.cashflow) assert(rendered.includes(`第 ${row.year} 年 | 流入 ${row.cashIn} | 流出 ${row.cashOut} | 净 ${row.netCash}`));
   assert.match(rendered,/逐年保留原始小数/);
  }
  if(tool===calculateSelectionFee||tool===reverseCalculateSelectionFee){assert.equal(value.selection_fee,undefined);assert.match(rendered,/不含税服务费/);assert.match(rendered,/含税服务费/);}
  if(tool===reverseCalculateSelectionFee){assert(value.quote_candidates.length>1);for(const q of value.quote_candidates)assert(rendered.includes(q));assert.match(rendered,/较低报价/);assert.match(rendered,/正算指定/);}
 }
 globalThis.fetch=async()=>new Response(JSON.stringify({error:'该限价落在资费表 10 万元档位跳变形成的空档，无对应报价'}),{status:422});
 await assert.rejects(reverseCalculateSelectionFee.execute({limit:'106600',markup:'0'},exec),/该限价落在资费表 10 万元档位跳变形成的空档，无对应报价/);
}finally{globalThis.fetch=originalFetch;}
console.log('Benefit tools: trusted identity, unknown-field forwarding, ungated read-only policy, basis render, tax bases, all exact candidates and verbatim errors passed.');
