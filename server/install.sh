#!/usr/bin/env bash
set -euo pipefail

if [[ "$(id -u)" -ne 0 ]]; then
  echo "Запустите установщик от root через sudo или SSH." >&2
  exit 1
fi

APP_DIR=/home/reportdeploy/wb-automation
PROFILE_DIR="$APP_DIR/profile"
RAW=https://raw.githubusercontent.com/san7ek-rgb/wb-analytics/main/server

if [[ ! -d "$PROFILE_DIR" ]]; then
  echo "Не найден профиль $PROFILE_DIR. Сначала скопируйте профиль Firefox по инструкции." >&2
  exit 1
fi

read -r -s -p "Введите токен WB_COST_RECEIVER_TOKEN из свойств Apps Script (ввод скрыт): " WB_RECEIVER_TOKEN
printf '\n'
if [[ ! "$WB_RECEIVER_TOKEN" =~ ^[A-Za-z0-9_-]{32,128}$ ]]; then
  unset WB_RECEIVER_TOKEN
  echo "Токен должен содержать от 32 до 128 символов: латинские буквы, цифры, дефис или подчёркивание." >&2
  exit 1
fi

install -d -o reportdeploy -g reportdeploy -m 700 "$APP_DIR"
TMP_DIR=$(mktemp -d)
trap 'rm -rf "$TMP_DIR"; unset WB_RECEIVER_TOKEN' EXIT
for file in wb_funnel_collector.py wb-funnel.service wb-funnel.timer; do
  curl -fsSL "$RAW/$file" -o "$TMP_DIR/$file"
done
install -o reportdeploy -g reportdeploy -m 600 "$TMP_DIR/wb_funnel_collector.py" "$APP_DIR/wb_funnel_collector.py"
printf 'WB_COST_RECEIVER_TOKEN=%s\nWB_FIREFOX_PROFILE=%s\n' "$WB_RECEIVER_TOKEN" "$PROFILE_DIR" > "$APP_DIR/.env"
chown reportdeploy:reportdeploy "$APP_DIR/.env"
chmod 600 "$APP_DIR/.env"
unset WB_RECEIVER_TOKEN

runuser -u reportdeploy -- python3 -m venv "$APP_DIR/venv"
runuser -u reportdeploy -- "$APP_DIR/venv/bin/pip" install --quiet selenium
install -o root -g root -m 644 "$TMP_DIR/wb-funnel.service" /etc/systemd/system/wb-funnel.service
install -o root -g root -m 644 "$TMP_DIR/wb-funnel.timer" /etc/systemd/system/wb-funnel.timer
systemctl daemon-reload
systemctl enable --now wb-funnel.timer
systemctl start --no-block wb-funnel.service
echo "Установка завершена. Таймер включён; первый сбор запущен."
systemctl list-timers wb-funnel.timer --no-pager
