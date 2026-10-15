#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
echo "启动志特AI（WPS 文字）... 按 Ctrl-C 退出"
npm run dev:wps
