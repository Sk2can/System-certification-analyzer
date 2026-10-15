from pathlib import Path
from app.services.word_report import SoftwareTableBuilder
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

    # Создаем экземпляр word файла
    builder = SoftwareTableBuilder(
        input_dir=input_dir,
        output_file=f"{result_dir}/software.docx",
    )

    builder.build()
    

if __name__ == "__main__":
    main()