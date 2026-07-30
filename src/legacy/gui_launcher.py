import sys
import os
from pathlib import Path

# Додаємо кореневу папку проекту до шляху Python для імпорту модулів
project_root = Path(__file__).parent.parent.resolve()
if str(project_root) not in sys.path:
    sys.path.append(str(project_root))

from python_src.main_processor import main_pdf_to_text_to_sheets, write_log, setup_logging

def display_menu():
    print("\n--- Меню Автоматичної Обробки Документів ---")
    print("1. Запустити обробку PDF/текстових файлів")
    print("2. Вийти")
    print("------------------------------------------")

def main_gui_launcher():
    setup_logging() # Ініціалізація логування для launcher
    write_log("Запущено gui_launcher.py", level='debug')

    while True:
        display_menu()
        choice = input("Виберіть опцію (1-2): ")

        if choice == '1':
            write_log("Користувач обрав Запустити обробку.", level='user')
            write_log("Виклик main_pdf_to_text_to_sheets()", level='debug')
            try:
                main_pdf_to_text_to_sheets()
            except Exception as e:
                write_log(f"Помилка при запуску обробки: {e}", level='user')
                write_log(f"Помилка при запуску обробки: {e}", level='debug')
            print("\nОбробка завершена. Перевірте лог-файли.")
        elif choice == '2':
            write_log("Користувач обрав Вийти.", level='user')
            print("Вихід з програми.")
            break
        else:
            print("Невірний вибір. Будь ласка, введіть 1 або 2.")

if __name__ == "__main__":
    main_gui_launcher()
