#!/usr/bin/env node
/**
 * Exercise the approval answerer against a live lamber bridge, without booting
 * dsh or spending an LLM call.
 *
 * This isolates the hop the no-API-key dsh tests cannot cover on their own:
 * answerer → HTTP → Rust gate → (frontend decision) → outcome back in-process.
 * Prints the resolved `ApprovalOutcome` on stdout.
 *
 * Usage:
 *   LAMBER_BRIDGE_URL=http://127.0.0.1:PORT LAMBER_BRIDGE_TOKEN=… \
 *     node scripts/check-approval.mjs [toolName] [reason]
 */
import { askLamber } from '../dsh-tool-lamber/lib/index.js';

const [toolName = 'write_test_marker', reason = '测试审批通道'] = process.argv.slice(2);

const outcome = await askLamber(
  toolName,
  `test-call-${Date.now()}`,
  reason,
  undefined,
);
process.stdout.write(`${outcome}\n`);
