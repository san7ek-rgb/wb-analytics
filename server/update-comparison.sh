#!/usr/bin/env bash
set -euo pipefail
if [[ "$(id -u)" -ne 0 ]]; then
  echo "Запустите обновление от root." >&2
  exit 1
fi
WB_UPDATE_REV=${1:-main}
if [[ ! "$WB_UPDATE_REV" =~ ^([a-f0-9]{7,40}|main)$ ]]; then
  echo "Некорректная версия." >&2
  exit 1
fi
WB_UPDATE_DIR=/home/reportdeploy/wb-automation
WB_UPDATE_RAW="https://raw.githubusercontent.com/san7ek-rgb/wb-analytics/$WB_UPDATE_REV/server"
if [[ ! -s "$WB_UPDATE_DIR/.env" || ! -d "$WB_UPDATE_DIR/profile" || ! -x "$WB_UPDATE_DIR/venv/bin/python" ]]; then
  echo "Не найдена существующая установка WB. Секреты не изменены." >&2
  exit 1
fi
WB_UPDATE_TEMP=$(mktemp -d)
trap 'rm -rf "$WB_UPDATE_TEMP"' EXIT
for file in wb_funnel_collector.py wb-funnel.service wb-comparison.service wb-comparison.timer; do
  curl -fsSL "$WB_UPDATE_RAW/$file" -o "$WB_UPDATE_TEMP/$file"
done
install -d -o reportdeploy -g reportdeploy -m 700 /home/reportdeploy/.cache/selenium
install -o reportdeploy -g reportdeploy -m 600 "$WB_UPDATE_TEMP/wb_funnel_collector.py" "$WB_UPDATE_DIR/wb_funnel_collector.py.new"
mv "$WB_UPDATE_DIR/wb_funnel_collector.py.new" "$WB_UPDATE_DIR/wb_funnel_collector.py"
for file in wb-funnel.service wb-comparison.service wb-comparison.timer; do
  install -o root -g root -m 644 "$WB_UPDATE_TEMP/$file" "/etc/systemd/system/$file"
done
systemctl daemon-reload
systemctl enable --now wb-comparison.timer
systemctl start --no-block wb-comparison.service
echo "Проверка запросов сравнения включена раз в минуту. Существующий токен и ежедневный таймер сохранены."
systemctl list-timers wb-comparison.timer wb-funnel.timer --no-pager
