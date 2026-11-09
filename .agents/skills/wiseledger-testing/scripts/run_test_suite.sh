#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
FINANCE_ROOT="$(cd "$SCRIPT_DIR/../../../.." && pwd)"

VENV_PYTHON="$FINANCE_ROOT/.venv/bin/python"
if [ ! -f "$VENV_PYTHON" ]; then
    VENV_PYTHON="python3"
fi

echo "============================================================"
echo "🧪 [Skill: wiseledger-testing] 启动自动化综合测试套件"
echo "工作目录: $FINANCE_ROOT"
echo "Python 解释器: $VENV_PYTHON"
echo "============================================================"

"$VENV_PYTHON" "$FINANCE_ROOT/tests/run_tests.py"
