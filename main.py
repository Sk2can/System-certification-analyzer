from pathlib import Path
from app.services.word_report import WordTableBuilder
from app.services.read_file import Parser
from dotenv import load_dotenv
import os   


def main():
    # Загружаем переменные окружения из .env
    load_dotenv("./app/settings/.env")
    input_dir = Path(os.getenv("input_dir", "."))
    result_dir = Path(os.getenv("result_dir", "results"))

    # Создаем директории, если их нет
    result_dir.mkdir(parents=True, exist_ok=True)
    input_dir.mkdir(parents=True, exist_ok=True)



    parser = Parser(dir_path=input_dir)
    records = parser.read_software_from_files(parser.search_txt_reports())

    builder = WordTableBuilder(
        config_path="app/settings/software_table.json",
        output_path="reports/software.docx",
    )

    builder.build(records)
    

if __name__ == "__main__":
    main()