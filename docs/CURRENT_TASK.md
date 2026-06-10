# macOS 1.0.1 发布

- **Status:** Done
- **Objective:** 将净现值率成本反算修复打包为 Apple Silicon macOS `1.0.1` 版本，并生成可分发 DMG。

## Progress

1. [x] 将根 npm 包、Tauri 配置和 Rust 包版本统一从 `1.0.0` / `0.1.0` 更新为 `1.0.1`。
2. [x] 运行发布前测试与前端生产构建。
3. [x] 使用显式 ad hoc 身份对完整应用包签名，并执行 Tauri macOS DMG 打包。
4. [x] 校验应用版本、CPU 架构、代码签名和 DMG 完整性。
5. [x] 生成发布产物：
   - `src-tauri/target/release/bundle/dmg/云数中心工具集_1.0.1_aarch64.dmg`
   - SHA-256: `8c036dda1111d43a967fd3e3820820f699c9f17bf37f07eca17996a9ae47f8c9`

## Validation

- ICT reverse-search and all subject-funding scripts: passed.
- `npx tsc --noEmit`: passed.
- `cargo fmt -- --check`: passed.
- `cargo test benefit::calculator::tests`: passed, 9 tests.
- Tauri release build: passed.
- DMG checksum verification: passed.
- Bundle versions: `CFBundleShortVersionString = 1.0.1`, `CFBundleVersion = 1.0.1`.
- Executable architecture: Apple Silicon `arm64`.
- `codesign --verify --deep --strict`: passed with ad hoc signature.

## Distribution Note

- 当前构建环境没有 Apple Developer 代码签名身份。
- 本次产物为 Apple Silicon `aarch64` DMG，使用临时签名，不包含 Developer ID 签名或 Apple 公证。
