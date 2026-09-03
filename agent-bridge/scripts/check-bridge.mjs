#!/usr/bin/env node
/**
 * Exercise the `run_benefit_calculation` tool body against a live lamber bridge,
 * without booting dsh or spending an LLM call.
 *
 * This isolates the one hop the no-API-key dsh test cannot cover: plugin
 * `execute()` → HTTP → Rust bridge → `benefit::calculator`. When the full agent
 * loop misbehaves, run this first to tell a tool/transport fault apart from a
 * model or harness fault.
 *
 * Usage:
 *   LAMBER_BRIDGE_URL=http://127.0.0.1:PORT LAMBER_BRIDGE_TOKEN=… \
 *     node scripts/check-bridge.mjs <projectId> [scenario]
 *
 * Prints the tool's canonical JSON value on stdout; exits non-zero on failure.
 */
import { runBenefitCalculation } from '../dsh-tool-lamber/lib/index.js';

const [projectId, scenario] = process.argv.slice(2);
if (!projectId) {
  console.error('usage: check-bridge.mjs <projectId> [scenario]');
  process.exit(2);
}

const controller = new AbortController();
try {
  const value = await runBenefitCalculation.execute(
    { projectId, ...(scenario ? { scenario } : {}) },
    { signal: controller.signal },
  );
  process.stdout.write(`${JSON.stringify(value)}\n`);
} catch (error) {
  console.error(`[check-bridge] ${error?.message ?? error}`);
  process.exit(1);
}
