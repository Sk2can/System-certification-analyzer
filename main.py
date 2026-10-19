from pathlib import Path
from app.services.word_report import WordTableBuilder
from app.services.read_file import BaseParser, SoftwareParser
from app.services.read_file import HardwareParser
from dotenv import load_dotenv
import os   


def main():
    # Загружаем переменные окружения из .env
    load_dotenv("./app/settings/.env")
    #input_dir = Path(os.getenv("input_dir", "."))
    input_dir = Path(str(input("Введите путь к директории с файлами: ")).strip())
    result_dir = Path(os.getenv("result_dir", "results"))

    # Создаем директории, если их нет
    result_dir.mkdir(parents=True, exist_ok=True)
    input_dir.mkdir(parents=True, exist_ok=True)



    parser = SoftwareParser(dir_path=input_dir)
    records = parser.read_software_from_files(BaseParser.search_txt_reports(input_dir))

    builder = WordTableBuilder(
        config_path="app/settings/software_table.json",
        output_path=f"{result_dir}/software.docx",
    )

    builder.build(records)


    parser = HardwareParser(input_dir)
    hardware_data = parser.parse(location="Кабинет 101")
    
    builder = WordTableBuilder(
            config_path="app/settings/hardware_table.json",
            output_path=f"{result_dir}/hardware.docx",
        )
    
    builder.build(hardware_data)
    

if __name__ == "__main__":
    try:
        main()
    except PermissionError as error:
        print(f"\nОШИБКА, файл занят другим процессом: {error}")
        raise SystemExit(1)