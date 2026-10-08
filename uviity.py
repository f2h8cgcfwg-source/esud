"""
Відкриває браузер програми, щоб юрист сам увійшов у кабінет Е-суду (КЕП / Дія.Підпис).
Після входу зберігає сесію у файл auth.json і в папку profil/.

Програма НЕ бачить і НЕ записує ваш ключ чи пароль: ви вводите їх самі,
а зберігається лише «перепустка» (cookies) уже відкритої сесії.

Запуск: python uviity.py
"""

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

PAPKA = Path(__file__).parent
PROFIL = PAPKA / "profil"
AUTH = PAPKA / "auth.json"
ADRESA = "https://id.court.gov.ua"


def main():
    with sync_playwright() as p:
        kontekst = p.chromium.launch_persistent_context(
            PROFIL, headless=False, locale="uk-UA", viewport={"width": 1400, "height": 900})
        storinka = kontekst.pages[0] if kontekst.pages else kontekst.new_page()
        storinka.goto(ADRESA)
        print("\nУ вікні браузера увійдіть у свій кабінет Е-суду, як зазвичай.")
        print("Коли побачите свій кабінет, поверніться сюди й натисніть Enter.")
        input("> ")
        kontekst.storage_state(path=AUTH)
        os.chmod(AUTH, 0o600)  # файл читає лише ваш користувач Mac
        kontekst.close()
    print(f"\n✅ Вхід збережено ({AUTH.name}). Нікому не пересилайте цей файл.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
