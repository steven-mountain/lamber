import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export function readPluginContract(pluginRoot) {
  // Read the emitted JavaScript, not a sidecar that can be refreshed separately.
  const source = readFileSync(resolve(pluginRoot, 'lib/contract.generated.js'), 'utf8');
  const match = /export const BRIDGE_CONTRACT = (\{[^\n]+\});/.exec(source);
  if (!match) throw new Error('AI 插件缺少编译后的契约，请重新构建。');
  return JSON.parse(match[1]);
}
export function readBinaryContract(binary) {
  const bytes = readFileSync(binary);
  const start = Buffer.from('LAMBER_BRIDGE_CONTRACT:');
  const end = Buffer.from(':END_LAMBER_BRIDGE_CONTRACT');
  const offset = bytes.indexOf(start);
  const stop = bytes.indexOf(end, offset + start.length);
  if (offset < 0 || stop < 0) throw new Error('Rust 程序缺少桥接契约，请重新构建。');
  return JSON.parse(bytes.subarray(offset + start.length, stop).toString('utf8'));
}
export function assertContractsEqual(plugin, binary) {
  if (plugin.version !== binary.version || JSON.stringify([...plugin.routes].sort()) !== JSON.stringify([...binary.routes].sort())) {
    throw new Error('AI 插件与 Rust 程序契约不匹配，打包已停止。请完整重新构建。');
  }
}
export function assertPluginContract(pluginRoot, expected) {
  assertContractsEqual(readPluginContract(pluginRoot), expected);
}
export function assertBinaryContract(pluginRoot, binary) {
  assertContractsEqual(readPluginContract(pluginRoot), readBinaryContract(binary));
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  assertPluginContract(process.argv[2], JSON.parse(readFileSync(process.argv[3], 'utf8')));
}
