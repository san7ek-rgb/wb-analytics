#!/usr/bin/env bash
set -euo pipefail
if [[ "$(id -u)" -ne 0 ]]; then
  echo "Установку нужно выполнить от root." >&2
  exit 1
fi
WB_AUTO_REV=${1:-main}
if [[ ! "$WB_AUTO_REV" =~ ^([a-f0-9]{7,40}|main)$ ]]; then
  echo "Некорректная версия." >&2
  exit 1
fi
WB_AUTO_RAW="https://raw.githubusercontent.com/san7ek-rgb/wb-analytics/$WB_AUTO_REV/server"
WB_AUTO_APP=/home/reportdeploy/wb-automation
if [[ ! -s "$WB_AUTO_APP/.env" || ! -d "$WB_AUTO_APP/profile" || ! -x "$WB_AUTO_APP/venv/bin/python" ]]; then
  echo "Не найдена существующая установка WB. Сначала нужна исходная установка." >&2
  exit 1
fi
WB_AUTO_TEMP=$(mktemp -d)
trap 'rm -rf "$WB_AUTO_TEMP"' EXIT
for file in auto_update.py wb-auto-update.service wb-auto-update.timer; do
  curl -fsSL "$WB_AUTO_RAW/$file" -o "$WB_AUTO_TEMP/$file"
done
/usr/bin/python3 -m py_compile "$WB_AUTO_TEMP/auto_update.py"
install -d -o root -g root -m 755 /usr/local/lib/wb-autoupdate
install -d -o root -g root -m 700 /var/lib/wb-autoupdate
install -d -o reportdeploy -g reportdeploy -m 700 /home/reportdeploy/.cache/selenium
install -o root -g root -m 644 "$WB_AUTO_TEMP/auto_update.py" /usr/local/lib/wb-autoupdate/auto_update.py
install -o root -g root -m 644 "$WB_AUTO_TEMP/wb-auto-update.service" /etc/systemd/system/wb-auto-update.service
install -o root -g root -m 644 "$WB_AUTO_TEMP/wb-auto-update.timer" /etc/systemd/system/wb-auto-update.timer
systemctl daemon-reload
touch /var/lib/wb-autoupdate/bootstrap-pending
systemctl enable --now wb-auto-update.timer
/usr/bin/python3 /usr/local/lib/wb-autoupdate/auto_update.py
if [[ ! -f /etc/systemd/system/wb-comparison.timer ]]; then
  echo "Обновление отложено до завершения текущего сбора; модуль повторит попытку автоматически."
fi
echo "Автообновление кода включено раз в две минуты. Ключ и профиль WB сохранены."
systemctl list-timers wb-auto-update.timer wb-comparison.timer wb-funnel.timer --no-pager
