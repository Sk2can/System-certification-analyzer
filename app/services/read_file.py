import json
import os
import re
from pathlib import Path
import sys
from typing import Any

    
class BaseParser:
    """Общие методы для работы с файлами."""

    @staticmethod
    def get_project_path(relation_path: str) -> Path:
        """Возвращает путь к файлу в проекте, независимо от того,
        где находится директория проекта.

        Args:
            relation_path: путь к файлу относительно корня проекта (например app/config/software_table.json)"""

        if getattr(sys, "frozen", False):
            base_path = Path(sys._MEIPASS)
        else:
            base_path = Path(__file__).resolve().parents[2]

        return Path(base_path / relation_path)

    @staticmethod
    def search_txt_reports(path: str | Path) -> list[Path]:
        """Возвращает список TXT-файлов в каталоге.
        
        Args:
            path: Абсолютный путь к каталогу.

        Returns:
            Список абсолютных путей к TXT-файлам.

        Raises:
            NotADirectoryError: Если указанный путь не является каталогом.
            FileNotFoundError: Если в каталоге нет TXT-файлов."""

        if not Path(path).is_dir():
            raise NotADirectoryError(
                f"Каталог не найден: {path}"
            )

        files = sorted(
            Path(path).glob("*.txt"),
            key=lambda path: path.stat().st_mtime,
        )

        if not files:
            raise FileNotFoundError(
                f"В каталоге нет TXT-файлов: {path}"
            )


        return files

    @staticmethod
    def get_computer_name(path: Path | str) -> str:
        """Возвращает имя компьютера и серийный номер."""

        # Получаем серийный номер из отчёта и название файла
        serial_number = ""
        computer_name = Path(path).stem.removeprefix("report_")

        with Path(path).open("r", encoding="utf-8-sig") as file:
            for line in file:
                if line.strip().lower().startswith("serial number:"):
                    serial_number = line.split(":", 1)[1].strip()
                    if " " in serial_number:
                        serial_number = None
                    break

        # Добавляем серийный номер к имени компьютера
        if serial_number:
            computer_name = f"{computer_name}\n[{serial_number}]"

        return computer_name

    @staticmethod
    def load_json(file_path: str | Path) -> Any:
            """Загружает данные из JSON-файла."""

            path = Path(file_path)

            with path.open("r", encoding="utf-8") as file:
                return json.load(file)


class SoftwareParser:
    """Парсер информации об установленном ПО."""

    def __init__(self, dir_path: str):
        self.dir_path = Path(dir_path)
        try:
            self.reports_list = sorted(BaseParser.search_txt_reports(dir_path))
        except (NotADirectoryError, FileNotFoundError) as error:
            print(f"Ошибка: {error}")
            self.reports_list = []

    @staticmethod
    def read_software_block(path: str | Path) -> list[dict]:
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

    def read_software_from_files(self) -> dict[str, list[dict]]:
        """Читает ПО из файлов и определяет категорию и назначение."""

        if not self.reports_list:
            print("Нет доступных отчетов для обработки.")
            return {}
        
        config_path = (
            Path(os.environ["APPDATA"])
            / "SCA"
            / "software_categories.json"
        )
        categories = BaseParser.load_json(config_path)

        result = {}

        for path in self.reports_list:
            path = Path(path)

            packages = self.read_software_block(path)

            for package in packages:
                package_name = str(package.get("name", "")).casefold()

                # Значения по умолчанию для ПО, которого нет в JSON.
                package["category"] = "Дополнительное ПО"
                package["purpose"] = ""

                for item in categories:
                    matches = item.get("match", [])

                    if any(
                        str(keyword).casefold() in package_name
                        for keyword in matches
                    ):
                        package["category"] = item["category"]
                        package["purpose"] = item.get("purpose", "")
                        break

            computer_name = BaseParser.get_computer_name(path)
            result[computer_name] = packages

        return result

    @staticmethod
    def get_software_category(
        package_name: str,
        categories: dict[str, list[str]],
    ) -> str | None:
        for category, names in categories.items():
            for name in names:
                if name.casefold() in package_name.casefold():
                    return category

        return "Дополнительное ПО"


class HardwareParser:
    """Парсер аппаратных характеристик одного или нескольких АРМ."""

    def __init__(self, path: str | Path):
        self.path = Path(path)

    def parse(self, location: str = "") -> list[dict]:
        """
        Обрабатывает один TXT-файл или все TXT-файлы в директории.

        Всегда возвращает список словарей.
        """
        if self.path.is_file():
            return [self.parse_file(self.path, location)]

        if self.path.is_dir():
            return self.parse_directory(location)

        raise FileNotFoundError(
            f"Файл или директория не найдены: {self.path}"
        )

    def parse_directory(self, location: str = "") -> list[dict]:
        """Последовательно обрабатывает TXT-файлы директории."""
        results = []

        for file_path in sorted(self.path.glob("*.txt")):
            try:
                result = self.parse_file(file_path, location)
                results.append(result)
            except (OSError, UnicodeError) as error:
                print(f"Не удалось прочитать {file_path}: {error}")

        return results

    def parse_file(
        self,
        file_path: str | Path,
        location: str = "",
    ) -> dict:
        """Извлекает характеристики из одного TXT-отчёта."""
        file_path = Path(file_path)

        if not file_path.is_file():
            raise FileNotFoundError(
                f"Файл не найден: {file_path}"
            )

        text = file_path.read_text(
            encoding="utf-8",
            errors="replace",
        )

        lines = text.splitlines()

        # Идентификатор оборудования находится в первой строке.
        equipment_id = BaseParser.get_computer_name(file_path)

        # Определяем ОС и извлекаем технические характеристики.
        os_name = self._find_os(text)

        motherboard = self._find_component(
            lines, "core", "product"
        )
        cpu = self._find_component(
            lines, "cpu", "product"
        )
        ram = self._find_component(
            lines, "memory", "size"
        )

        disks = self._find_disks(lines)

        technical_specs = []

        if motherboard:
            technical_specs.append(
                f"Материнская плата: {motherboard}"
            )

        if cpu:
            technical_specs.append(
                f"Процессор: {cpu}"
            )

        if ram:
            technical_specs.append(
                f"Оперативная память: {ram}"
            )

        if disks:
            technical_specs.append(
                f"Накопители: {', '.join(disks)}"
            )

        return {
            "equipment_id": equipment_id,
            "location": location,
            "technical_specs": ".\n".join(technical_specs) +".",
            "operating_system": os_name,
        }

    @staticmethod
    def _find_os(text: str) -> str:
        """Извлекает название ОС и редакцию из отчёта."""

        pretty_name = re.search(
            r'^\s*PRETTY_NAME\s*=\s*["\']?(.*?)["\']?\s*$',
            text,
            re.MULTILINE,
        )

        edition = re.search(
            r'^\s*EDITION\s*=\s*["\']?(.*?)["\']?\s*$',
            text,
            re.MULTILINE,
        )

        os_name = (
            pretty_name.group(1).strip().strip("\"'")
            if pretty_name
            else ""
        )

        os_edition = (
            edition.group(1).strip().strip("\"'")
            if edition
            else ""
        )

        return " ".join(filter(None, [os_name, os_edition]))

    @staticmethod
    def _marker_info(line: str) -> tuple[int, str] | None:
        """Возвращает отступ и имя компонента lshw."""
        match = re.match(
            r"^(\s*)\*-([A-Za-z0-9_.:-]+)",
            line,
        )

        if not match:
            return None

        return len(match.group(1)), match.group(2)

    @classmethod
    def _component_lines(
        cls,
        lines: list[str],
        component: str,
    ) -> list[str]:
        """
        Находит компонент и возвращает его строки,
        включая вложенные узлы, если они есть.
        """
        for index, line in enumerate(lines):
            marker = cls._marker_info(line)

            if marker is None:
                continue

            indent, name = marker

            if name != component:
                continue

            result = []

            for child_line in lines[index + 1:]:
                child_marker = cls._marker_info(child_line)

                if child_marker is not None:
                    child_indent, _ = child_marker

                    # Компонент на том же или более высоком уровне
                    # завершает текущую секцию.
                    if child_indent <= indent:
                        break

                result.append(child_line)

            return result

        return []

    @classmethod
    def _find_component(
        cls,
        lines: list[str],
        component: str,
        field: str,
    ) -> str:
        """Извлекает product или size из компонента lshw."""
        component_lines = cls._component_lines(
            lines, component
        )

        # Сначала ищем значение в самом компоненте.
        # Вложенные секции здесь не нужны.
        for line in component_lines:
            if cls._marker_info(line) is not None:
                break

            match = re.match(
                rf"^\s*{re.escape(field)}:\s*(.+?)\s*$",
                line,
            )

            if match:
                return match.group(1).strip()

        return ""

    @classmethod
    def _find_disks(cls, lines: list[str]) -> list[str]:
        """Извлекает названия физических накопителей и их размеры."""
        disks = []

        for index, line in enumerate(lines):
            marker = cls._marker_info(line)

            if marker is None:
                continue

            indent, component = marker

            # Интересуют дисковые устройства, а не разделы,
            # контроллеры или файловые системы.
            if component not in {"nvme", "disk"}:
                continue

            block = []

            for child_line in lines[index + 1:]:
                child_marker = cls._marker_info(child_line)

                if child_marker is not None:
                    child_indent, _ = child_marker

                    if child_indent <= indent:
                        break

                block.append(child_line)

            product = ""
            size = ""

            for block_line in block:
                if self_marker := cls._marker_info(block_line):
                    # Название и размер самого устройства
                    # ищем также во вложенных namespace.
                    pass

                product_match = re.match(
                    r"^\s*product:\s*(.+?)\s*$",
                    block_line,
                )
                size_match = re.match(
                    r"^\s*size:\s*(.+?)\s*$",
                    block_line,
                )

                if product_match and not product:
                    product = product_match.group(1).strip()

                if size_match and not size:
                    size = size_match.group(1).strip()

            if product:
                description = product

                if size:
                    description += f" ({size})"

                if description not in disks:
                    disks.append(description)

        return disks
