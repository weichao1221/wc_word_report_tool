#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

rm -rf dist build src/*.egg-info(N)
python3 -m build
python3 -m twine check dist/*

echo
echo "构建校验通过。PyPI 发布由 GitHub Actions 完成："
echo "  1. 提交版本变更并推送 main"
echo "  2. git tag v0.7.6"
echo "  3. git push origin v0.7.6"
