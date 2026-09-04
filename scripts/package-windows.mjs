import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import {
  existsSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const VERSION_FILES = [
  "package.json",
  "package-lock.json",
  "src-ui/package.json",
  "src-ui/package-lock.json",
  "src-tauri/tauri.conf.json",
  "src-tauri/Cargo.toml",
  "src-tauri/Cargo.lock",
];

export function bumpVersion(version, releaseType = "patch") {
  const match = /^(\d+)\.(\d+)\.(\d+)$/.exec(version);
  if (!match) {
    throw new Error(`Unsupported version "${version}". Expected MAJOR.MINOR.PATCH.`);
  }

  let [, major, minor, patch] = match.map(Number);
  if (releaseType === "major") {
    major += 1;
    minor = 0;
    patch = 0;
  } else if (releaseType === "minor") {
    minor += 1;
    patch = 0;
  } else if (releaseType === "patch") {
    patch += 1;
  } else {
    throw new Error(`Unsupported release type "${releaseType}". Use patch, minor, or major.`);
  }

  return `${major}.${minor}.${patch}`;
}

export function replaceCargoTomlVersion(content, nextVersion) {
  const pattern = /(\[package\][\s\S]*?\r?\nversion\s*=\s*")([^"]+)(")/;
  if (!pattern.test(content)) {
    throw new Error("Could not find [package] version in src-tauri/Cargo.toml.");
  }
  return content.replace(pattern, `$1${nextVersion}$3`);
}

export function replaceCargoLockVersion(content, nextVersion) {
  const pattern =
    /(\[\[package\]\]\r?\nname = "benefit-calculator"\r?\nversion = ")([^"]+)(")/;
  if (!pattern.test(content)) {
    throw new Error("Could not find benefit-calculator version in src-tauri/Cargo.lock.");
  }
  return content.replace(pattern, `$1${nextVersion}$3`);
}

export function replaceJsonVersion(content, nextVersion, label) {
  const pattern = /("version"\s*:\s*")([^"]+)(")/;
  if (!pattern.test(content)) {
    throw new Error(`Could not find version in ${label}.`);
  }
  return content.replace(pattern, `$1${nextVersion}$3`);
}

function readJson(relativePath) {
  return JSON.parse(readFileSync(join(ROOT, relativePath), "utf8"));
}

function packageVersionFromToml(content) {
  const match = /\[package\][\s\S]*?\r?\nversion\s*=\s*"([^"]+)"/.exec(content);
  if (!match) {
    throw new Error("Could not read [package] version from src-tauri/Cargo.toml.");
  }
  return match[1];
}

function packageVersionFromCargoLock(content) {
  const match =
    /\[\[package\]\]\r?\nname = "benefit-calculator"\r?\nversion = "([^"]+)"/.exec(
      content,
    );
  if (!match) {
    throw new Error("Could not read benefit-calculator version from src-tauri/Cargo.lock.");
  }
  return match[1];
}

function assertVersionsAreSynchronized(snapshot) {
  const rootPackage = JSON.parse(snapshot.get("package.json"));
  const rootLock = JSON.parse(snapshot.get("package-lock.json"));
  const uiPackage = JSON.parse(snapshot.get("src-ui/package.json"));
  const uiLock = JSON.parse(snapshot.get("src-ui/package-lock.json"));
  const tauriConfig = JSON.parse(snapshot.get("src-tauri/tauri.conf.json"));
  const versions = new Map([
    ["package.json", rootPackage.version],
    ["package-lock.json", rootLock.version],
    ['package-lock.json packages[""]', rootLock.packages?.[""]?.version],
    ["src-ui/package.json", uiPackage.version],
    ["src-ui/package-lock.json", uiLock.version],
    ['src-ui/package-lock.json packages[""]', uiLock.packages?.[""]?.version],
    ["src-tauri/tauri.conf.json", tauriConfig.version],
    [
      "src-tauri/Cargo.toml",
      packageVersionFromToml(snapshot.get("src-tauri/Cargo.toml")),
    ],
    [
      "src-tauri/Cargo.lock",
      packageVersionFromCargoLock(snapshot.get("src-tauri/Cargo.lock")),
    ],
  ]);
  const uniqueVersions = new Set(versions.values());

  if (uniqueVersions.size !== 1 || uniqueVersions.has(undefined)) {
    const details = [...versions].map(([file, version]) => `${file}: ${version}`).join("\n");
    throw new Error(`Version files are not synchronized:\n${details}`);
  }

  return rootPackage.version;
}

function stringifyJson(data, original) {
  const newline = original.includes("\r\n") ? "\r\n" : "\n";
  return `${JSON.stringify(data, null, 2).replaceAll("\n", newline)}${newline}`;
}

function updateVersionFiles(snapshot, nextVersion) {
  const rootPackage = JSON.parse(snapshot.get("package.json"));
  rootPackage.version = nextVersion;

  const rootLock = JSON.parse(snapshot.get("package-lock.json"));
  rootLock.version = nextVersion;
  rootLock.packages[""].version = nextVersion;

  const uiPackage = JSON.parse(snapshot.get("src-ui/package.json"));
  uiPackage.version = nextVersion;

  const uiLock = JSON.parse(snapshot.get("src-ui/package-lock.json"));
  uiLock.version = nextVersion;
  uiLock.packages[""].version = nextVersion;

  const updates = new Map([
    ["package.json", stringifyJson(rootPackage, snapshot.get("package.json"))],
    ["package-lock.json", stringifyJson(rootLock, snapshot.get("package-lock.json"))],
    ["src-ui/package.json", stringifyJson(uiPackage, snapshot.get("src-ui/package.json"))],
    [
      "src-ui/package-lock.json",
      stringifyJson(uiLock, snapshot.get("src-ui/package-lock.json")),
    ],
    [
      "src-tauri/tauri.conf.json",
      replaceJsonVersion(
        snapshot.get("src-tauri/tauri.conf.json"),
        nextVersion,
        "src-tauri/tauri.conf.json",
      ),
    ],
    [
      "src-tauri/Cargo.toml",
      replaceCargoTomlVersion(snapshot.get("src-tauri/Cargo.toml"), nextVersion),
    ],
    [
      "src-tauri/Cargo.lock",
      replaceCargoLockVersion(snapshot.get("src-tauri/Cargo.lock"), nextVersion),
    ],
  ]);

  for (const [relativePath, content] of updates) {
    writeFileSync(join(ROOT, relativePath), content, "utf8");
  }
}

function restoreVersionFiles(snapshot) {
  for (const [relativePath, content] of snapshot) {
    writeFileSync(join(ROOT, relativePath), content, "utf8");
  }
}

function collectFiles(directory) {
  if (!existsSync(directory)) {
    return [];
  }

  return readdirSync(directory).flatMap((name) => {
    const path = join(directory, name);
    return statSync(path).isDirectory() ? collectFiles(path) : [path];
  });
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex").toUpperCase();
}

function printArtifacts(version) {
  const bundleDirectory = join(ROOT, "src-tauri", "target", "release", "bundle");
  const artifacts = collectFiles(bundleDirectory).filter((path) =>
    path.includes(`_${version}_`),
  );

  if (artifacts.length === 0) {
    console.warn(`Build succeeded, but no versioned bundle was found for ${version}.`);
    return;
  }

  console.log("\nRelease artifacts:");
  for (const path of artifacts) {
    const sizeMb = (statSync(path).size / 1024 / 1024).toFixed(2);
    console.log(`- ${relative(ROOT, path)} (${sizeMb} MB)`);
    console.log(`  SHA-256: ${sha256(path)}`);
  }
}

function printHelp() {
  console.log(`Usage: npm run package:windows -- [patch|minor|major]

Default: patch
Examples:
  npm run package:windows          1.1.0 -> 1.1.1
  npm run package:windows -- minor 1.1.1 -> 1.2.0
  npm run package:windows -- major 1.2.0 -> 2.0.0`);
}

export function main(args = process.argv.slice(2)) {
  if (args.includes("--help") || args.includes("-h")) {
    printHelp();
    return 0;
  }
  if (process.platform !== "win32") {
    throw new Error("The Windows packaging command must run on Windows.");
  }

  const releaseType = args[0] ?? "patch";
  const snapshot = new Map(
    VERSION_FILES.map((relativePath) => [
      relativePath,
      readFileSync(join(ROOT, relativePath), "utf8"),
    ]),
  );
  const currentVersion = assertVersionsAreSynchronized(snapshot);
  const nextVersion = bumpVersion(currentVersion, releaseType);

  console.log(`Packaging Windows release ${currentVersion} -> ${nextVersion} (${releaseType})`);
  updateVersionFiles(snapshot, nextVersion);

  const result = spawnSync("npm run tauri -- build", {
    cwd: ROOT,
    shell: true,
    stdio: "inherit",
  });

  if (result.status !== 0) {
    restoreVersionFiles(snapshot);
    if (result.error) {
      console.error(`Build command failed to start: ${result.error.message}`);
    }
    console.error(`Packaging failed. Version files were restored to ${currentVersion}.`);
    return result.status ?? 1;
  }

  printArtifacts(nextVersion);
  console.log(`\nPackaging completed. Current project version is ${nextVersion}.`);
  return 0;
}

const invokedDirectly =
  process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href;
if (invokedDirectly) {
  try {
    process.exitCode = main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  }
}
