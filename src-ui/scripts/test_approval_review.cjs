const assert = require('node:assert/strict');
const fs = require('node:fs');
const ts = require('typescript');
const mod = {exports:{}};
new Function('module','exports',ts.transpileModule(fs.readFileSync('src/ai/approvalReview.ts','utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS}}).outputText)(mod,mod.exports);
const {remainingSeconds,amendedArguments,readableArguments}=mod.exports;
const prose='第一段：审核长文本。\n\n第二段：保留换行。'.repeat(30);
const prompt={args:{note:prose,immutable:'project-A'},intent:{fields:[{key:'note',proposedValue:prose,previousValue:'原文'}]}};
assert.equal(amendedArguments(prompt,{note:prose}),undefined);
assert.deepEqual(amendedArguments(prompt,{note:'人工修订\n正文'}),{note:'人工修订\n正文',immutable:'project-A'});
assert.equal(prompt.args.note,prose);
assert.equal(amendedArguments({...prompt,intent:null},{note:'越权修改'}),undefined);
assert.deepEqual(readableArguments({note:prose}),[{label:'note',text:prose}]);
const now=Date.now(), deadline=new Date(now+600000).toISOString();
assert.equal(remainingSeconds(deadline,now),600);
assert.equal(remainingSeconds(deadline,now+590000),10); // queue wait never resets the deadline
assert.equal(remainingSeconds(deadline,now+601000),0);
assert.equal(remainingSeconds('invalid',now),0);
console.log('approval review: original preserved, edited payload, prose newlines, absolute queue deadline passed');

const nested={...prompt,toolName:'fill_template_fields',args:{projectId:'p',templateId:'需求导入表.docx',fields:{note:prose}}};
assert.deepEqual(amendedArguments(nested,{note:'修改'}),{projectId:'p',templateId:'需求导入表.docx',fields:{note:'修改'}});
