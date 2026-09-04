#!/usr/bin/env node
/**
 * Provision the dsh `acp` profile used by lamber's agent bridge.
 *
 * dsh resolves a patched plugin's package name relative to
 * `$DSH_HOME/profiles/<profile>/`, not relative to the launcher's cwd. So the
 * local `dsh-tool-lamber` package must be linked into that profile directory
 * before `dsh --patch patch.yml` can find it. This script is idempotent: run it
 * once per machine, or whenever $DSH_HOME is reset.
 *
 * Usage: node scripts/provision-profile.mjs [--profile acp] [--home <dir>]
 */
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const agentBridgeDir = resolve(here, '..');
const pluginDir = join(agentBridgeDir, 'dsh-tool-lamber');

function arg(flag, fallback) {
  const i = process.argv.indexOf(flag);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}

const profile = arg('--profile', 'acp');
const dshHome = resolve(
  arg('--home', process.env.DSH_HOME ?? join(agentBridgeDir, '.dsh-home')),
);

function run(cmd, args, opts = {}) {
  const binDir = join(agentBridgeDir, 'node_modules', '.bin');
  const result = spawnSync(cmd, args, {
    stdio: 'inherit',
    cwd: agentBridgeDir,
    ...opts,
    env: {
      ...process.env,
      PATH: `${binDir}:${process.env.PATH ?? ''}`,
      DSH_HOME: dshHome,
      DSH_TELEMETRY_MODE: process.env.DSH_TELEMETRY_MODE ?? 'DISABLED',
      ...opts.env,
    },
  });
  if (result.status !== 0) {
    throw new Error(`${cmd} ${args.join(' ')} exited with ${result.status}`);
  }
}

if (!existsSync(join(pluginDir, 'lib', 'index.js'))) {
  console.log('[provision] building dsh-tool-lamber…');
  run('npm', ['run', 'build'], { cwd: pluginDir });
}

mkdirSync(dshHome, { recursive: true });
console.log(`[provision] DSH_HOME=${dshHome}`);
console.log(`[provision] linking ${pluginDir} into profile "${profile}"…`);
run('dsh', ['plugin', '--profile', profile, 'add', pluginDir]);

console.log('[provision] done. Launch with:');
console.log(
  `  DSH_HOME=${dshHome} DSH_TELEMETRY_MODE=DISABLED LAMBER_BRIDGE_URL=http://127.0.0.1:<port> \\`,
);
console.log(
  `    ${join(agentBridgeDir, 'node_modules', '.bin', 'dsh')} --profile ${profile} --patch ${join(agentBridgeDir, 'patch.yml')}`,
);
