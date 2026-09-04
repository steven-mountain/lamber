import assert from "node:assert/strict";
import test from "node:test";

import {
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
