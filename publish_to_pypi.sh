#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

VERSION="$(sed -n 's/^version = "\(.*\)"/\1/p' pyproject.toml)"

rm -rf dist build
rm -rf src/*.egg-info 2>/dev/null || true
python3 -m build
python3 -m twine check dist/*

echo
echo "构建校验通过，当前版本 v${VERSION}。"
echo "PyPI 发布有两条路："
echo "  A) 手动上传（当前可用，需要本机 ~/.pypirc 里有 API token）："
echo "       python3 -m twine upload dist/*"
echo "  B) 推标签触发 GitHub Actions（需先在 PyPI 配好 Trusted Publisher）："
echo "       git tag v${VERSION} && git push origin v${VERSION}"
echo "     未配置时会报 invalid-publisher —— v0.7.6 / v0.8.0 / v0.8.3 都是这么失败的。"
