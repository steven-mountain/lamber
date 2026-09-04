#!/usr/bin/env node
/**
 * Report which tools this plugin gates behind human approval.
 *
 * The guard's policy is a pure function of the tool name, so it can be asserted
 * without an agent run: this prints one `<name>=<true|false>` line per argument.
 * Used to prove that read-only tools stay ungated after approval work lands.
 *
 * Usage: node scripts/check-gating.mjs run_benefit_calculation write_test_marker
 */
import { isGatedTool } from '../dsh-tool-lamber/lib/index.js';

const names = process.argv.slice(2);
if (names.length === 0) {
  console.error('usage: check-gating.mjs <toolName...>');
  process.exit(2);
}
for (const name of names) {
  process.stdout.write(`${name}=${isGatedTool(name)}\n`);
}
