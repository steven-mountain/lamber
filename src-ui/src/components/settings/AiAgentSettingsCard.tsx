import { useEffect, useState } from "react";
import {
  getAiAgentSettings,
  saveAiAgentSettings,
  type AiAgentSettings,
} from "../../ai/agentSettings";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../ui/card";
import { Input } from "../ui/input";
import { Label } from "../ui/label";

const OFFICIAL_BASE_URL = "https://api.deepseek.com";

export default function AiAgentSettingsCard({ onSaved, showDiagnostics = true }: { onSaved?: (settings: AiAgentSettings) => void; showDiagnostics?: boolean } = {}) {
  const [settings, setSettings] = useState<AiAgentSettings | null>(null);
  const [model, setModel] = useState("");
  const [baseUrl, setBaseUrl] = useState(OFFICIAL_BASE_URL);
  const [apiKey, setApiKey] = useState("");
  const [clearApiKey, setClearApiKey] = useState(false);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState<{ kind: "success" | "error"; text: string } | null>(null);

  useEffect(() => {
    let active = true;
    getAiAgentSettings()
      .then((loaded) => {
        if (!active) return;
        setSettings(loaded);
        setModel(loaded.model);
        setBaseUrl(loaded.baseUrl);
      })
      .catch((error) => {
        if (active) setMessage({ kind: "error", text: String(error) });
      })
      .finally(() => {
        if (active) setLoading(false);
      });
    return () => {
      active = false;
    };
  }, []);

  const save = async () => {
    if (!model) return;
    setSaving(true);
    setMessage(null);
    try {
      const updated = await saveAiAgentSettings({
        model,
        baseUrl,
        apiKey: apiKey.trim() || undefined,
        clearApiKey,
      });
      setSettings(updated);
      onSaved?.(updated);
      setModel(updated.model);
      setBaseUrl(updated.baseUrl);
      setApiKey("");
      setClearApiKey(false);
      setMessage({ kind: "success", text: "AI 配置已保存，AI 窗口会重新连接。已有会话的模型请在聊天输入框切换。" });
    } catch (error) {
      setMessage({ kind: "error", text: String(error) });
    } finally {
      setSaving(false);
    }
  };

  const selectedModel = settings?.models.find((candidate) => candidate.id === model);
  const customEndpoint = baseUrl.trim().replace(/\/+$/, "") !== OFFICIAL_BASE_URL;
  const openAgentLab = () => {
    window.location.hash = "#/agent-lab";
    window.location.reload();
  };

  return (
    <Card className="shadow-sm">
      <CardHeader>
        <CardTitle className="text-section-title">AI 模型与服务</CardTitle>
        <CardDescription className="text-caption">
          设置新会话的默认模型和服务连接。已有会话保留自己的模型选择；密钥仅保存在本机，不会回显。
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        {loading ? (
          <div className="rounded-xl bg-muted/40 p-4 text-caption text-secondary-foreground">正在读取 AI 配置…</div>
        ) : (
          <>
            <div className="grid gap-4 sm:grid-cols-2">
              <div className="space-y-1.5">
                <Label htmlFor="ai-agent-model" className="text-label">新会话默认模型</Label>
                <select
                  id="ai-agent-model"
                  value={model}
                  onChange={(event) => setModel(event.target.value)}
                  className="flex h-9 w-full rounded-md bg-background px-3 py-1 text-sm shadow-sm outline-none ring-1 ring-input focus-visible:ring-2 focus-visible:ring-ring/30"
                >
                  {settings?.models.map((candidate) => (
                    <option key={candidate.id} value={candidate.id}>
                      {candidate.name}{candidate.supportsImages ? "（支持图片）" : ""}
                    </option>
                  ))}
                </select>
                <p className="text-caption text-secondary-foreground">
                  {selectedModel?.supportsImages
                    ? "该模型支持图片输入。"
                    : "当前模型不支持图片；如需图片输入，请选择视觉模型。"}
                </p>
              </div>

              <div className="space-y-1.5">
                <Label htmlFor="ai-agent-key" className="text-label">API Key</Label>
                <Input
                  id="ai-agent-key"
                  type="password"
                  autoComplete="off"
                  value={apiKey}
                  onChange={(event) => {
                    setApiKey(event.target.value);
                    setClearApiKey(false);
                  }}
                  placeholder={settings?.hasApiKey ? "已保存；留空表示不更改" : "请输入 DeepSeek API Key"}
                />
                <div className="flex items-center justify-between gap-3 text-caption text-secondary-foreground">
                  <span>{clearApiKey ? "保存后将清除密钥" : settings?.hasApiKey ? "本机已保存密钥" : "尚未保存密钥"}</span>
                  {(settings?.hasApiKey || clearApiKey) && (
                    <button
                      type="button"
                      className="rounded-md bg-muted px-2 py-1 text-destructive transition-colors hover:bg-destructive-soft"
                      onClick={() => {
                        setApiKey("");
                        setClearApiKey(!clearApiKey);
                      }}
                    >
                      {clearApiKey ? "撤销清除" : "清除密钥"}
                    </button>
                  )}
                </div>
              </div>
            </div>

            <div className="space-y-1.5">
              <Label htmlFor="ai-agent-base-url" className="text-label">服务地址（baseURL）</Label>
              <Input
                id="ai-agent-base-url"
                type="url"
                value={baseUrl}
                onChange={(event) => setBaseUrl(event.target.value)}
                placeholder={OFFICIAL_BASE_URL}
              />
              <p className={`rounded-lg p-3 text-caption ${customEndpoint ? "bg-warning-soft text-warning-foreground" : "bg-muted/40 text-secondary-foreground"}`}>
                {customEndpoint
                  ? "当前版本只验证过 DeepSeek 官方端点。普通 OpenAI 兼容服务可能拒绝 dsh 的 thinking、reasoning_effort 或插件扩展字段。"
                  : "DeepSeek 官方端点是当前已验证配置。修改模型、密钥或服务地址后，保存会停止旧 dsh 子进程。"}
              </p>
            </div>

            {message && (
              <div
                role="status"
                className={`flex items-start gap-2 rounded-lg p-3 text-caption ${
                  message.kind === "success"
                    ? "bg-success-soft text-success-foreground"
                    : "bg-destructive-soft text-destructive"
                }`}
              >
                <AppIcon name={message.kind === "success" ? "check" : "warning"} size={14} className="mt-0.5 shrink-0" />
                <span>{message.text}</span>
              </div>
            )}

            <div className="flex flex-wrap justify-end gap-2">
              {showDiagnostics && <Button type="button" variant="secondary" onClick={openAgentLab}>
                打开 dsh 联调台
              </Button>}
              <Button onClick={save} disabled={saving || !model || !baseUrl.trim()}>
                {saving ? "正在保存…" : "保存 AI 配置"}
              </Button>
            </div>
          </>
        )}
      </CardContent>
    </Card>
  );
}
