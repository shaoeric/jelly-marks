#!/bin/bash
# 唛头生成工具启动器(macOS 双击运行)
# 用 uv 管理环境:优先 uv run,找不到 uv 时回退到项目内的 .venv。

cd "$(dirname "$0")" || exit 1

if command -v uv >/dev/null 2>&1; then
    exec uv run gui_marks.py
fi

if [ -x ".venv/bin/python" ]; then
    exec .venv/bin/python gui_marks.py
fi

echo "未找到 uv,也没有可用的 .venv。"
echo "请先安装 uv: https://docs.astral.sh/uv/"
echo "或手动执行: uv sync"
echo "按任意键关闭窗口…"
read -n 1
