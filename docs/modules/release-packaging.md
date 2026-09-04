# Windows Release Packaging

## Standard command

Run packaging from the repository root:

```powershell
npm run package:windows
```

This command increments the patch version before building. For example, `1.1.0`
becomes `1.1.1`.

## Version policy

- `patch` is the default for routine fixes and packaging iterations:
  `npm run package:windows`
- `minor` is used for backward-compatible features:
  `npm run package:windows:minor`
- `major` is reserved for incompatible product or data-contract changes:
  `npm run package:windows:major`

The release script requires all version sources to match before it starts. It
updates:

- root `package.json` and `package-lock.json`
- `src-ui/package.json` and `src-ui/package-lock.json`
- `src-tauri/tauri.conf.json`
- `src-tauri/Cargo.toml` and the application package entry in `Cargo.lock`

If the Tauri build fails, all version files are restored to their original
contents. A successful build keeps the new version and prints each generated
bundle path, size, and SHA-256 checksum.

## Output

The Windows NSIS installer is generated under:

```text
src-tauri/target/release/bundle/nsis/
```

The installer filename contains the incremented version. Build outputs remain
ignored by Git; commit the synchronized version files with the release changes.

## Validation

The versioning rules have a dependency-free Node test:

```powershell
npm run test:packaging
```
