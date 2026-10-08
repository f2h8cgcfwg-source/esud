#!/bin/bash
# Установка всього потрібного на Mac. Запускати подвійним клацанням (або: bash vstanovyty.command).
cd "$(dirname "$0")" || exit 1

zupynka() { echo; echo "❌ $1"; echo; read -r -p "Натисніть Enter, щоб закрити вікно..."; exit 1; }

echo "=== Установка програми для Е-суду ==="
echo "Папка: $(pwd)"
echo

# 1. Python: шукаємо версію з python.org, потім будь-яку python3.
PY=""
for kandydat in /Library/Frameworks/Python.framework/Versions/3.1[0-9]/bin/python3 /usr/local/bin/python3 /opt/homebrew/bin/python3; do
  [ -x "$kandydat" ] && PY="$kandydat"
done
if [ -z "$PY" ] && xcode-select -p >/dev/null 2>&1; then
  PY="$(command -v python3)"
fi
if [ -z "$PY" ]; then
  open "https://www.python.org/downloads/macos/"
  zupynka "Python не знайдено. Я відкрив сторінку завантаження: встановіть Python 3.13 (macOS 64-bit universal2 installer), а потім знову запустіть цей файл."
fi
"$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' \
  || zupynka "Python занадто старий ($("$PY" --version)). Встановіть Python 3.13 з https://www.python.org/downloads/macos/"
echo "✅ Python: $("$PY" --version)"

# 2. Окреме середовище для бібліотек.
if [ ! -x .venv/bin/python ]; then
  echo "Створюю окреме середовище (.venv)..."
  "$PY" -m venv .venv || zupynka "Не вдалося створити середовище .venv."
fi
source .venv/bin/activate

# 3. Бібліотеки та браузер.
echo "Встановлюю бібліотеки (1–3 хвилини)..."
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet || zupynka "Не вдалося встановити бібліотеки. Перевірте інтернет."
echo "✅ Бібліотеки встановлено"
echo "Завантажую браузер Chromium для програми (~150 МБ)..."
python -m playwright install chromium || zupynka "Не вдалося завантажити Chromium. Перевірте інтернет."
echo "✅ Chromium встановлено"

# 4. Перевірка на зразках.
echo
echo "Перевіряю читач заяв на зразках..."
python chytach.py samples/zrazky.pdf || zupynka "Читач заяв не спрацював. Надішліть Claude текст із цього вікна."
open rezultat/zrazky.html
echo
echo "🎉 Готово! У браузері відкрилася сторінка перевірки трьох зразків."
echo "Наступний крок — запис форми Е-суду: файл zapys.command (див. KROK3_ZAPYS.md)."
read -r -p "Натисніть Enter, щоб закрити вікно..."
