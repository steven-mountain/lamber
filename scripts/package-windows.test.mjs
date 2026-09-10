import assert from "node:assert/strict";
import test from "node:test";

import {
  assertAgentPackageMetadata,
  bumpVersion,
  replaceCargoLockVersion,
  replaceCargoTomlVersion,
  replaceJsonVersion,
} from "./package-windows.mjs";

test("increments semantic versions by release type", () => {
  assert.equal(bumpVersion("1.1.0"), "1.1.1");
  assert.equal(bumpVersion("1.1.9", "minor"), "1.2.0");
  assert.equal(bumpVersion("1.9.9", "major"), "2.0.0");
});

test("rejects unsupported versions and release types", () => {
  assert.throws(() => bumpVersion("1.0"), /Expected MAJOR\.MINOR\.PATCH/);
  assert.throws(() => bumpVersion("1.0.0", "build"), /Use patch, minor, or major/);
});

test("updates only the application package version in Cargo manifests", () => {
  const cargoToml = `[package]\nname = "benefit-calculator"\nversion = "1.1.0"\n\n[dependencies]\nserde = "1.0"\n`;
  const cargoLock = `[[package]]\nname = "benefit-calculator"\nversion = "1.1.0"\n\n[[package]]\nname = "serde"\nversion = "1.0.0"\n`;

  assert.match(replaceCargoTomlVersion(cargoToml, "1.1.1"), /version = "1\.1\.1"/);
  assert.match(replaceCargoLockVersion(cargoLock, "1.1.1"), /version = "1\.1\.1"/);
  assert.match(replaceCargoLockVersion(cargoLock, "1.1.1"), /serde"\nversion = "1\.0\.0"/);
});

test("updates JSON versions without reformatting the document", () => {
  const tauriConfig = `{\n  "productName": "云数中心工具集",\n  "version": "1.1.0",\n  "bundle": {\n    "targets": ["nsis"]\n  }\n}\n`;

  const updated = replaceJsonVersion(tauriConfig, "1.1.1", "tauri.conf.json");

  assert.match(updated, /"version": "1\.1\.1"/);
  assert.match(updated, /"targets": \["nsis"\]/);
});

test("requires dsh to be pinned as a production dependency", () => {
  assert.doesNotThrow(() =>
    assertAgentPackageMetadata(
      { dependencies: { "@deepseek-ai/dsh": "0.1.2-alpha.5" }, devDependencies: { pnpm: "^10" } },
      { packages: { "": { dependencies: { "@deepseek-ai/dsh": "0.1.2-alpha.5" } } } },
    ),
  );
  assert.throws(
    () =>
      assertAgentPackageMetadata(
        { devDependencies: { "@deepseek-ai/dsh": "0.1.2-alpha.5" } },
        { packages: { "": { devDependencies: { "@deepseek-ai/dsh": "0.1.2-alpha.5" } } } },
      ),
    /production dependency/,
  );
});

import { mkdtempSync, mkdirSync, writeFileSync, rmSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { assertBinaryContract, assertPluginContract } from './bridge-contract.mjs';

test('packaging rejects old and mismatched compiled binaries and stale plugin output', () => {
  const root = mkdtempSync(join(tmpdir(), 'lamber-contract-packaging-'));
  try {
    const contract = JSON.parse(readFileSync(new URL('../agent-bridge/bridge-contract.json', import.meta.url)));
    mkdirSync(join(root, 'lib'));
    writeFileSync(join(root, 'lib/contract.generated.js'), `export const BRIDGE_CONTRACT = ${JSON.stringify(contract)};\n`);
    const binary = join(root, 'app');
    const writeBinary = value => writeFileSync(binary, Buffer.concat([Buffer.from([0, 128, 255]), Buffer.from(`LAMBER_BRIDGE_CONTRACT:${JSON.stringify(value)}:END_LAMBER_BRIDGE_CONTRACT`)]));
    writeBinary(contract);
    assert.doesNotThrow(() => assertBinaryContract(root, binary));
    assert.doesNotThrow(() => assertPluginContract(root, contract));
    writeBinary({ ...contract, version: contract.version + 1 });
    assert.throws(() => assertBinaryContract(root, binary), /打包已停止/);
    writeBinary({ ...contract, routes: contract.routes.slice(1) });
    assert.throws(() => assertBinaryContract(root, binary), /打包已停止/);
    writeFileSync(binary, 'legacy executable');
    assert.throws(() => assertBinaryContract(root, binary), /缺少桥接契约/);
    assert.throws(() => assertPluginContract(root, { ...contract, version: 99 }), /打包已停止/);
  } finally { rmSync(root, { recursive: true, force: true }); }
});
