#!/bin/bash
# Крок 3: запис одного проходу створення чернетки в Е-суді.
# Спершу вхід (без запису), потім запис дій уже в кабінеті — щоб пароль ключа НЕ потрапив у запис.
cd "$(dirname "$0")" || exit 1
[ -x .venv/bin/python ] || { echo "Спершу запустіть vstanovyty.command"; read -r; exit 1; }
source .venv/bin/activate
mkdir -p zapys

echo "=== Частина 1 з 2: вхід у кабінет (нічого не записується) ==="
python uviity.py || { echo "Вхід не вдався. Надішліть Claude текст із цього вікна."; read -r; exit 1; }

FAIL="zapys/codegen_$(date +%Y-%m-%d_%H-%M).py"
echo
echo "=== Частина 2 з 2: запис ==="
echo "Відкриються два вікна: браузер (уже з вашим кабінетом) і вікно Playwright Inspector із кодом."
echo "Створіть ОДНУ тестову чернетку за інструкцією KROK3_ZAPYS.md."
echo "НЕ натискайте «Підписати» чи «Надіслати». Наприкінці просто закрийте вікно браузера."
echo
python -m playwright codegen --target python --load-storage=auth.json --save-storage=auth.json \
  --lang=uk-UA --viewport-size=1400,900 -o "$FAIL" https://id.court.gov.ua

echo
if [ -s "$FAIL" ]; then
  echo "✅ Запис збережено у файл: $FAIL"
  open -R "$FAIL"
  echo "Надішліть цей файл Claude (перетягніть його у вікно чату)."
else
  echo "⚠️ Запис порожній. Можливо, браузер закрили до початку дій. Спробуйте ще раз."
fi
read -r -p "Натисніть Enter, щоб закрити вікно..."
