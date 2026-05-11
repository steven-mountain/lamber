#!/bin/bash
# Navigate to the script's directory
cd "$(dirname "$0")"

echo "正在启动云数中心工具集..."
npm run tauri dev
