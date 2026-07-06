#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

rm -rf dist build src/*.egg-info(N)
python3 -m build
python3 -m twine check dist/*

echo
echo "如果上面检查通过，执行下面命令正式上传："
echo "python3 -m twine upload dist/*"
echo
echo "如果你想先传测试仓库："
echo "python3 -m twine upload --repository testpypi dist/*"
