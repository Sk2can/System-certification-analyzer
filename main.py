from pathlib import Path
from app.services.word_report import WordTableBuilder
from app.services.read_file import BaseParser, SoftwareParser
from app.services.read_file import HardwareParser


def main():
    input_dir = Path(str(input("Введите путь к директории с отчетами: ")).strip())
    result_dir = Path(str(input("Введите путь к директории с результатами: ")).strip())

    # Создаем директории, если их нет
    result_dir.mkdir(parents=True, exist_ok=True)
    input_dir.mkdir(parents=True, exist_ok=True)

    parser = SoftwareParser(dir_path=input_dir)
    records = parser.read_software_from_files()
    builder = WordTableBuilder(
        config_path= BaseParser.get_project_path("app\\settings\\software_table.json"),
        output_path=f"{result_dir}/software.docx",
    )
    builder.build(records)


    parser = HardwareParser(input_dir)
    hardware_data = parser.parse(location=str(input("Введите местонахождение оборудования: ")))
    builder = WordTableBuilder(
            config_path=BaseParser.get_project_path("app\\settings\\hardware_table.json"),
            output_path=f"{result_dir}/hardware.docx",
        )
    builder.build(hardware_data)

if __name__ == "__main__":
    try:
        main()
    except PermissionError as error:
        print(f"\nОШИБКА, файл занят другим процессом: {error}")
        raise SystemExit(1)
    finally:
        input("\nНажмите Enter для выхода...")