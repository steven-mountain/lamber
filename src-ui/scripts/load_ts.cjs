const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");
const cache = new Map();
module.exports = function load(file) {
  file = path.resolve(file);
  if (file.endsWith(".json")) return JSON.parse(fs.readFileSync(file,"utf8"));
  if (cache.has(file)) return cache.get(file);
  const mod = {exports:{}};
  cache.set(file,mod.exports);
  const code = ts.transpileModule(fs.readFileSync(file,"utf8"), {compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2020,esModuleInterop:true}}).outputText;
  vm.runInNewContext(code,{module:mod,exports:mod.exports,console,require:name=> name.startsWith(".") ? load(path.resolve(path.dirname(file),name)+(path.extname(name)?"":".ts")) : require(name)}, {filename:file});
  return mod.exports;
};
