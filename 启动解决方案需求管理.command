#!/bin/zsh
set -e

SOLUTION_LAUNCHER_DIR="${0:A:h}"
SOLUTION_APP_PATH="$SOLUTION_LAUNCHER_DIR/release/解决方案需求管理-v0.7.0.app"

if [[ ! -d "$SOLUTION_APP_PATH" ]]; then
  echo "未找到解决方案需求管理：$SOLUTION_APP_PATH"
  echo "请确认 release 文件夹没有被移动或删除。"
  read -k 1 "?按任意键关闭…"
  exit 1
fi

/usr/bin/open "$SOLUTION_APP_PATH"
