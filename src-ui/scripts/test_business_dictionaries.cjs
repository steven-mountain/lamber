const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");
const serviceSource = fs.readFileSync(
  path.join(root, "src/services/businessDictionaryService.ts"),
  "utf8",
);
const selectSource = fs.readFileSync(
  path.join(root, "src/components/business-dictionaries/BusinessDictionarySelect.tsx"),
  "utf8",
);
const managerSource = fs.readFileSync(
  path.join(root, "src/components/business-dictionaries/BusinessDictionaryManager.tsx"),
  "utf8",
);
const templateSource = fs.readFileSync(
  path.join(root, "src/views/TemplateForms.tsx"),
  "utf8",
);
const presetCenterSource = fs.readFileSync(
  path.join(root, "src/views/PresetCenterView.tsx"),
  "utf8",
);

assert.match(serviceSource, /list_business_dictionaries/);
assert.match(serviceSource, /get_business_dictionary_options/);
assert.match(serviceSource, /save_business_dictionary_item/);
assert.match(serviceSource, /reorder_business_dictionary_items/);

assert.match(selectSource, /当前值，已停用或未配置/);
assert.match(selectSource, /fallbackOptions/);
assert.match(selectSource, /businessDictionaryService\.getOptions/);

assert.match(managerSource, /新增字典项/);
assert.match(managerSource, /setItemEnabled/);
assert.match(managerSource, /deleteItem/);
assert.match(managerSource, /reorderItems/);
assert.match(presetCenterSource, /业务字典/);
assert.match(presetCenterSource, /BusinessDictionaryManager/);

for (const dictionaryKey of [
  "business_model",
  "funding_source",
  "procurement_method",
  "yes_no",
]) {
  assert.match(templateSource, new RegExp(`dictionaryKey="${dictionaryKey}"`));
}
assert.match(templateSource, /implementationConstructionInterface/);
assert.match(templateSource, /procurementSingleSourceBasis/);
assert.match(templateSource, /demandSecurityDetail/);

console.log("Business dictionary tests passed.");
