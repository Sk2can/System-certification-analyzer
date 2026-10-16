import json
from pathlib import Path


class Parser:
    def __init__(self, dir_path: str):
        self.path = Path(dir_path)
        self.reports_list = self.search_txt_reports()

    def read_json(self) -> dict:
        """Читает конфигурационный JSON."""
        with self.path.open("r", encoding="utf-8") as file:
            return json.load(file)

    def read_software_block(self, path: str | Path) -> list[dict]:
        """Читает блок 'software' из файла отчета."""
        path = Path(path)

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

    def read_software_from_files(
        self,
        paths: list[str | Path]
    ) -> dict[str, list[dict]]:
        """Читает блок установленного ПО из нескольких файлов."""
        result = {}

        for path in paths:
            path = Path(path)

            packages = self.read_software_block(path)
            result[path.stem] = packages

        return result

    def search_txt_reports(self) -> list[Path]:
        """Возвращает список TXT-файлов в каталоге."""

        if not self.path.is_dir():
            raise NotADirectoryError(
                f"Каталог не найден: {self.path}"
            )

        files = sorted(
            self.path.glob("*.txt"),
            key=lambda path: path.stat().st_mtime,
        )

        if not files:
            raise FileNotFoundError(
                f"В каталоге нет TXT-файлов: {self.path}"
            )

        return files

    def read_device_name(self) -> str:
        """Читает идентификатор оборудования из первой строки TXT-файла."""
        file_path = self.path
        with open(file_path, "r", encoding="utf-8-sig") as file:
            equipment_id = file.readline().strip()

        if not equipment_id:
            raise ValueError(
                f"В первой строке файла {file_path} "
                "не указан идентификатор оборудования"
            )
        return equipment_id


if __name__ == "__main__":
    path = Path("C:\\Users\\labadmin\\Desktop\\SCA\\Info\\report_kgauitc-22-05.txt")
    parser = Parser(path)
    equipment_id = parser.read_device_name()
    software_packages = parser.read_software_block()
    print(f"Equipment ID: {equipment_id}")
    print("Software Packages:")
    for package in software_packages:
        print(package)
