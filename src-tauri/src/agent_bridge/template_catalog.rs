//! Data-only registry shared with the UI/plugin. Never infer writable keys from completion checks.
use serde::Deserialize;
use serde_json::{json, Value};
use std::sync::OnceLock;
#[derive(Deserialize)]
#[serde(untagged)]
pub enum RequiredCondition { Flag(String), Equals(EqualsCondition) }
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
pub struct EqualsCondition { pub field: String, pub equals: Value }
impl RequiredCondition {
    pub fn field(&self) -> &str { match self { Self::Flag(key) => key, Self::Equals(c) => &c.field } }
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CompletionSource { pub field: String, pub source: Option<String>, pub default_value: Option<String> }
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ValidRow { pub non_empty: Vec<String>, pub positive: Vec<String> }
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Field {
    pub key: String,
    pub label: String,
    pub kind: String,
    pub default_value: Option<String>,
    pub state_key: Option<String>,
    pub required_when: Option<RequiredCondition>,
    pub completion_group: Option<String>,
    pub completion_sources: Option<Vec<CompletionSource>>,
    pub valid_row: Option<ValidRow>,
    pub list_type: Option<String>,
    pub columns: Option<Vec<String>>,
    pub dynamic_default: Option<bool>,
    pub reason: Option<String>,
    pub completion_always: Option<bool>,
}
impl Field {
    pub fn raw<'a>(&self, state: &'a Value) -> Option<&'a str> {
        self.state_key.as_ref().and_then(|key| state.get(key))
            .or_else(|| if self.state_key.is_none() {state.get("formData").and_then(|f| f.get(&self.key))} else {None})
            .and_then(Value::as_str)
    }
    pub fn value(&self, state: &Value) -> Option<String> {
        self.raw(state).map(str::to_owned).or_else(|| self.default_value.clone())
    }
    pub fn write(&self, state: &mut Value, value: Value) -> Result<(), String> {
        let root = state.as_object_mut().ok_or("模板保存态格式损坏，拒绝覆盖")?;
        if let Some(key) = &self.state_key { root.insert(key.clone(), value); }
        else {
            let form = root.entry("formData").or_insert_with(|| json!({})).as_object_mut().ok_or("模板字段格式损坏，拒绝覆盖")?;
            form.insert(self.key.clone(), value);
        }
        Ok(())
    }
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Template {
    pub id: String,
    pub name: String,
    pub suffix: String,
    pub excluded_reason: Option<String>,
    pub fields: Vec<Field>,
}
#[derive(Deserialize)]
struct Catalog { templates: Vec<Template> }
pub fn templates() -> &'static [Template] {
    static CATALOG: OnceLock<Catalog> = OnceLock::new();
    &CATALOG.get_or_init(|| serde_json::from_str(include_str!("../../../src-ui/src/lib/templateCompletion/catalog.json")).expect("shared template catalog")).templates
}
pub fn resolve(name: &str, allow_alias: bool) -> Result<&'static Template, String> {
    if name.len() > 256 || name.contains(['/', '\\', '\0', ':']) { return Err("请提供完整模板名称，不能使用路径".into()); }
    let matches: Vec<_> = templates().iter().filter(|t| t.excluded_reason.is_none() &&
        ((allow_alias && name == t.name) || (name.contains(&t.name) && name.ends_with(&t.suffix)))).collect();
    if matches.len() != 1 { return Err("模板不在支持目录中或名称有歧义，请提供完整模板名".into()); }
    Ok(matches[0])
}
