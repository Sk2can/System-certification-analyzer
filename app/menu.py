from pathlib import Path
from app.services.category_manager import run_category_manager
from app.services.read_file import BaseParser, HardwareParser, SoftwareParser
from app.services.word_report import WordTableBuilder


def show_menu() -> str:
    """Выводит главное меню и возвращает выбор пользователя."""

    print("\n" + "=" * 40)
    print("   SCA — System-certification-analyzer")
    print("=" * 40)
    print("1. Отчёт по оборудованию")
    print("2. Отчёт по программному обеспечению")
    print("3. Оба отчёта")
    print("4. Настройки")
    print("0. Выход")
    print("=" * 40)

    return input("Выберите действие: ").strip()

def run_menu():
    """Основной цикл меню."""
    try:
        while True:
            choice = show_menu()

            if choice == "1":
                print("Формирование отчёта по оборудованию...")
                create_hardware_report()

            elif choice == "2":
                print("Формирование отчёта по ПО...")
                create_software_report()

            elif choice == "3":
                print("Формирование обоих отчётов...")
                create_combined_report()

            elif choice == "4":
                print("Настройки программы...")
                run_category_manager()

            elif choice == "0":
                print("Завершение работы.")
                break

            else:
                print("Ошибка: выберите пункт от 0 до 4.")
                continue

            input("\nНажмите Enter, чтобы вернуться в меню.")

    except (KeyboardInterrupt, EOFError):
        print("\nРабота программы завершена.")

def create_hardware_report():
    input_dir = Path(str(input("Введите путь к директории с отчетами: ")).strip())
    result_dir = Path(str(input("Введите путь к директории с результатами: ")).strip())
    hardware_file_name = str(input("Введите название таблицы оборудования (Enter - по умолчанию): "))

    parser = HardwareParser(input_dir)
    hardware_data = parser.parse(location=str(input("Введите местонахождение оборудования: ")))
    if not hardware_data:
        print("Нет доступных отчетов об оборудовании для обработки.")
        return

    # Создаем директорию, если ее нет
    result_dir.mkdir(parents=True, exist_ok=True)
    
    builder = WordTableBuilder(
            config_path=BaseParser.get_project_path("app\\settings\\hardware_table.json"),
            output_path=WordTableBuilder.input_filename(result_dir, hardware_file_name, "hardware_report.docx"),
        )
    builder.build(hardware_data)

def create_software_report():
    input_dir = Path(str(input("Введите путь к директории с отчетами: ")).strip())
    result_dir = Path(str(input("Введите путь к директории с результатами: ")).strip())
    software_file_name = str(input("Введите название таблицы ПО (Enter - по умолчанию): "))

    parser = SoftwareParser(dir_path=input_dir)
    records = parser.read_software_from_files()
    if not records:
        print("Нет доступных отчетов о программном обеспечении для обработки.")
        return

    # Создаем директорию, если ее нет
    result_dir.mkdir(parents=True, exist_ok=True)
    
    builder = WordTableBuilder(
        config_path= BaseParser.get_project_path("app\\settings\\software_table.json"),
        output_path=WordTableBuilder.input_filename(result_dir, software_file_name, "software_report.docx"),
    )
    builder.build(records)

def create_combined_report():
    input_dir = Path(str(input("Введите путь к директории с отчетами: ")).strip())
    result_dir = Path(str(input("Введите путь к директории с результатами: ")).strip())
    software_file_name = str(input("Введите название таблицы ПО (Enter - по умолчанию): "))
    hardware_file_name = str(input("Введите название таблицы оборудования (Enter - по умолчанию): "))

    parser = SoftwareParser(dir_path=input_dir)
    records = parser.read_software_from_files()
    if not records:
        print("Нет доступных отчетов о программном обеспечении для обработки.")
        return
    builder = WordTableBuilder(
        config_path= BaseParser.get_project_path("app\\settings\\software_table.json"),
        output_path=WordTableBuilder.input_filename(result_dir, software_file_name, "software_report.docx"),
    )
    builder.build(records)

    parser = HardwareParser(input_dir)
    hardware_data = parser.parse(location=str(input("Введите местонахождение оборудования: ")))
    if not hardware_data:
        print("Нет доступных отчетов об оборудовании для обработки.")
        return

    # Создаем директорию, если ее нет
    result_dir.mkdir(parents=True, exist_ok=True)

    builder = WordTableBuilder(
            config_path=BaseParser.get_project_path("app\\settings\\hardware_table.json"),
            output_path=WordTableBuilder.input_filename(result_dir, hardware_file_name, "hardware_report.docx"),
        )
    builder.build(hardware_data)