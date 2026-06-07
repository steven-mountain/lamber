const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const read = relative => fs.readFileSync(path.join(root, relative), "utf8");

const registry = read("src/lib/presetFieldKeys.ts");
const fields = read("src/lib/projectPresetFields.ts");
const actions = read("src/components/project-presets/ProjectPresetProjectActions.tsx");
const manager = read("src/components/project-presets/ProjectPresetManager.tsx");
const board = read("src/views/ProjectBoard.tsx");
const templateForms = read("src/views/TemplateForms.tsx");

const expect = (condition, message) => {
  if (!condition) throw new Error(message);
};

expect(fields.includes("definition.presetEligible || Boolean(definition.dictionaryKey)"),
  "project preset eligibility must allow registry presets or dictionary fields only");
expect(registry.includes('"amount"') && registry.includes('"percent"') && registry.includes('"computed"'),
  "registry must retain explicit financial/computed field types");
expect(actions.includes('"fill_empty_only"') && actions.includes('"overwrite_all"') && actions.includes('"selected_fields"'),
  "all project preset application strategies must exist");
expect(actions.includes("当前值") && actions.includes("预设值") && actions.includes("覆盖已有字段内容"),
  "application preview and overwrite confirmation must remain visible");
expect(actions.includes("saveCurrentProject()"),
  "project preset application must use the unified project save path");
expect(manager.includes("field?.label") && manager.includes("field?.templates") && manager.includes("field?.groups"),
  "manager must display business field metadata");
expect(!manager.includes("finance.npv") && !manager.includes("finance.tax_rate"),
  "manager must not hardcode financial fields into project presets");
expect(board.includes("空白项目") && board.includes("newProjectPresetId"),
  "new project flow must support optional project presets");
expect(templateForms.includes("onProjectPresetBindingsChange"),
  "template form values must be exposed through controlled bindings");

console.log("project preset frontend checks passed");
