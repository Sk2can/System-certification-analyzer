import json
from pathlib import Path


class Parser:
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)

    def read_json(self) -> dict:
        """Читает конфигурационный JSON."""
        with self.file_path.open("r", encoding="utf-8") as file:
            return json.load(file)

    def read_software_block(self) -> list[dict]:
        """Читает блок 'software' из файла отчета."""
        path = Path(self.file_path)

        with path.open("r", encoding="utf-8") as file:
            lines = file.readlines()

        packages = []
        inside_block = False

        for line in lines:
            line = line.strip()

            # Находим начало нужного блока
            if line == "Установленные пакеты":
                inside_block = True
                continue

            # Если встретили разделитель после начала блока — завершаем чтение
            if inside_block and line and set(line) == {"*"}:
                break

            if not inside_block or not line:
                continue

            # Пропускаем заголовки таблицы, если они есть
            if line.lower().startswith(("название", "пакет")):
                continue

            # Разделяем строку по двум и более пробелам
            parts = line.split(maxsplit=1)

            if len(parts) < 2:
                continue

            name = parts[0]
            remaining = parts[1].split(maxsplit=1)

            if len(remaining) < 2:
                continue

            version = remaining[0]
            repository = remaining[1]

            packages.append({
                "name": name,
                "version": version,
                "repository": repository
            })

        return packages

    def read_equipment_id(self) -> str:
        """Читает идентификатор оборудования из первой строки TXT-файла."""
        file_path = self.file_path
        with open(file_path, "r", encoding="utf-8-sig") as file:
            equipment_id = file.readline().strip()

        if not equipment_id:
            raise ValueError(
                f"В первой строке файла {file_path} "
                "не указан идентификатор оборудования"
            )
        return equipment_id