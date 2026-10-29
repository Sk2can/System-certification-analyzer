import json
from collections import OrderedDict
from pathlib import Path
import re
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class WordTableBuilder:
    """
    Универсальный генератор таблиц Word на основе JSON-конфигурации.

    Основные возможности:
    - настройка страницы и шрифтов;
    - создание таблицы произвольной структуры;
    - объединение ячеек;
    - группировка данных;
    - агрегация значений в одну ячейку;
    - удаление дубликатов;
    - повтор заголовка на каждой странице;
    - запрет разрыва строк таблицы между страницами.
    """

    ALIGNMENTS = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }

    PAGE_SIZES = {
        "A4": (21.0, 29.7),
        "A3": (29.7, 42.0),
        "LETTER": (21.59, 27.94),
    }

    def __init__(self, config_path, output_path):
        """
        config_path — путь к JSON-конфигурации.
        output_path — путь для сохранения DOCX.
        """
        self.config_path = Path(config_path)
        self.output_path = Path(output_path)

        print(f"Загрузка конфигурации: {self.config_path}")

        self.config = self.load_config()
        self.document = None
        self.table = None

    # ---------------------------------------------------------
    # ЗАГРУЗКА КОНФИГУРАЦИИ
    # ---------------------------------------------------------

    def load_config(self):
        """Загружает JSON-конфигурацию."""
        if not self.config_path.exists():
            raise FileNotFoundError(
                f"Файл конфигурации не найден: {self.config_path}"
            )

        with self.config_path.open("r", encoding="utf-8") as file:
            config = json.load(file)

        if "table" not in config:
            raise ValueError(
                "В конфигурации отсутствует раздел 'table'."
            )

        if not config["table"].get("columns"):
            raise ValueError(
                "В разделе 'table' необходимо определить 'columns'."
            )

        return config

    # ---------------------------------------------------------
    # СОЗДАНИЕ ДОКУМЕНТА И НАСТРОЙКА СТРАНИЦЫ
    # ---------------------------------------------------------

    def create_document(self):
        """Создаёт документ Word и применяет настройки."""
        self.document = Document()

        self.configure_page()
        self.configure_default_style()

        title = self.config.get("document", {}).get("title")

        if title:
            paragraph = self.document.add_paragraph()
            paragraph.add_run(title)

            title_config = self.config.get("document", {}).get(
                "title_style", {}
            )

            self.configure_paragraph(
                paragraph,
                alignment=title_config.get("align", "center"),
                bold=title_config.get("bold", True),
                font_size=title_config.get("font_size_pt", 14),
                font_name=title_config.get("font_name"),
            )

        return self.document


    def configure_page(self):
        """Настраивает формат страницы, ориентацию и поля."""
        document_config = self.config.get("document", {})
        section = self.document.sections[0]
        self.prepare = self.config.get("data", {}).get("prepare", False)
        
        page_size = document_config.get("page_size", "A4").upper()

        if page_size not in self.PAGE_SIZES:
            raise ValueError(
                f"Неизвестный формат страницы: {page_size}"
            )

        width, height = self.PAGE_SIZES[page_size]

        orientation = document_config.get(
            "orientation", "portrait"
        ).lower()

        if orientation == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Cm(max(width, height))
            section.page_height = Cm(min(width, height))
        elif orientation == "portrait":
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Cm(min(width, height))
            section.page_height = Cm(max(width, height))
        else:
            raise ValueError(
                "orientation должен быть 'portrait' или 'landscape'."
            )

        margins = document_config.get("margins_cm", {})

        section.top_margin = Cm(margins.get("top", 1.5))
        section.bottom_margin = Cm(margins.get("bottom", 1.5))
        section.left_margin = Cm(margins.get("left", 1.5))
        section.right_margin = Cm(margins.get("right", 1.5))

        section.header_distance = Cm(
            document_config.get("header_distance_cm", 0.7)
        )
        section.footer_distance = Cm(
            document_config.get("footer_distance_cm", 0.7)
        )


    def configure_default_style(self):
        """Задаёт шрифт по умолчанию для документа."""
        document_config = self.config.get("document", {})
        font_config = document_config.get("font", {})

        font_name = font_config.get("name", "Times New Roman")
        font_size = font_config.get("size_pt", 9)

        style = self.document.styles["Normal"]
        style.font.name = font_name
        style.font.size = Pt(font_size)

        style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

        paragraph_format = style.paragraph_format
        paragraph_config = document_config.get("paragraph", {})

        paragraph_format.space_before = Pt(
            paragraph_config.get("space_before_pt", 0)
        )
        paragraph_format.space_after = Pt(
            paragraph_config.get("space_after_pt", 0)
        )
        paragraph_format.line_spacing = paragraph_config.get(
            "line_spacing", 1.0
        )


    # ---------------------------------------------------------
    # ФОРМАТИРОВАНИЕ ТЕКСТА И ЯЧЕЕК
    # ---------------------------------------------------------
    def configure_paragraph(
        self,
        paragraph,
        alignment="left",
        bold=False,
        font_size=None,
        font_name=None,
    ):
        """Применяет настройки к абзацу и его тексту."""
        paragraph.alignment = self.ALIGNMENTS.get(
            alignment, WD_ALIGN_PARAGRAPH.LEFT
        )

        document_config = self.config.get("document", {})
        paragraph_config = document_config.get("paragraph", {})
        font_config = document_config.get("font", {})

        paragraph.paragraph_format.space_before = Pt(
            paragraph_config.get("space_before_pt", 0)
        )
        paragraph.paragraph_format.space_after = Pt(
            paragraph_config.get("space_after_pt", 0)
        )
        paragraph.paragraph_format.line_spacing = paragraph_config.get(
            "line_spacing", 1.0
        )

        if font_name is None:
            font_name = font_config.get("name", "Times New Roman")

        if font_size is None:
            font_size = font_config.get("size_pt", 9)

        for run in paragraph.runs:
            run.bold = bold
            run.font.name = font_name
            run.font.size = Pt(font_size)

            run._element.get_or_add_rPr().get_or_add_rFonts().set(
                qn("w:eastAsia"), font_name
            )


    def configure_cell(
        self,
        cell,
        alignment="left",
        bold=False,
        font_size=None,
        font_name=None,
    ):
        """Форматирует текст, вертикальное выравнивание и поля ячейки."""
        table_config = self.config.get("table", {})
        margins = table_config.get("cell_margins_cm", {})

        cell.vertical_alignment = self._vertical_alignment(
            table_config.get("vertical_alignment", "center")
        )

        # В Word отступы ячеек задаются в twips.
        # 1 см примерно равен 567 twips.
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()

        tc_mar = tc_pr.find(qn("w:tcMar"))
        if tc_mar is None:
            tc_mar = OxmlElement("w:tcMar")
            tc_pr.append(tc_mar)

        for side in ("top", "bottom", "left", "right"):
            tag = qn(f"w:{side}")
            element = tc_mar.find(tag)

            if element is None:
                element = OxmlElement(f"w:{side}")
                tc_mar.append(element)

            element.set(
                qn("w:w"),
                str(round(margins.get(side, 0.1) * 567))
            )
            element.set(qn("w:type"), "dxa")

        for paragraph in cell.paragraphs:
            self.configure_paragraph(
                paragraph,
                alignment=alignment,
                bold=bold,
                font_size=font_size,
                font_name=font_name,
            )


    @staticmethod
    def _vertical_alignment(value):
        """Преобразует строку в тип вертикального выравнивания Word."""
        mapping = {
            "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
            "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
            "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
        }
        return mapping.get(
            value, WD_CELL_VERTICAL_ALIGNMENT.CENTER
        )


    # ---------------------------------------------------------
    # НАСТРОЙКА ТАБЛИЦЫ
    # ---------------------------------------------------------
    def create_table(self):
        """Создаёт таблицу и заполняет её заголовками."""
        table_config = self.config["table"]
        columns = table_config["columns"]

        self.table = self.document.add_table(
            rows=1,
            cols=len(columns)
        )

        style_name = table_config.get("style", "Table Grid")

        try:
            self.table.style = style_name
        except KeyError:
            raise ValueError(
                f"Стиль таблицы Word не найден: {style_name}"
            )

        self.table.autofit = table_config.get("autofit", False)

        # Фиксированная ширина столбцов.
        tbl_pr = self.table._tbl.tblPr
        layout = tbl_pr.find(qn("w:tblLayout"))

        if layout is None:
            layout = OxmlElement("w:tblLayout")
            tbl_pr.append(layout)

        layout.set(qn("w:type"), "fixed")

        # Заголовки и ширина каждого столбца.
        header_row = self.table.rows[0]

        for index, column in enumerate(columns):
            cell = header_row.cells[index]
            cell.text = column.get("header", "")

            width_cm = column.get("width_cm")
            if width_cm is not None:
                cell.width = Cm(width_cm)
                self.table.columns[index].width = Cm(width_cm)

            header_style = table_config.get("header_style", {})

            self.configure_cell(
                cell,
                alignment=header_style.get(
                    "align", "center"
                ),
                bold=header_style.get("bold", True),
                font_size=header_style.get("font_size_pt"),
                font_name=header_style.get("font_name"),
            )

        # При необходимости объединяем ячейки заголовка.
        merge_config = table_config.get("merge", {})

        if merge_config.get("header", False):
            header_columns = merge_config.get(
                "header_columns",
                merge_config.get("data_columns", [])
            )

            if header_columns:
                merged_cell = self.merge_cells(
                    header_row, header_columns
                )

                merged_cell.text = merge_config.get(
                    "header_text", ""
                )

                header_style = table_config.get("header_style", {})

                self.configure_cell(
                    merged_cell,
                    alignment=header_style.get("align", "center"),
                    bold=header_style.get("bold", True),
                    font_size=header_style.get("font_size_pt"),
                    font_name=header_style.get("font_name"),
                )

        if table_config.get("repeat_header", True):
            self.set_repeat_table_header(header_row)

        return self.table


    @staticmethod
    def set_repeat_table_header(row):
        """Помечает строку как повторяемый заголовок таблицы."""
        tr_pr = row._tr.get_or_add_trPr()

        if tr_pr.find(qn("w:tblHeader")) is None:
            element = OxmlElement("w:tblHeader")
            element.set(qn("w:val"), "true")
            tr_pr.append(element)


    @staticmethod
    def set_row_cant_split(row):
        """Запрещает разделять одну строку таблицы между страницами."""
        tr_pr = row._tr.get_or_add_trPr()

        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))


    @staticmethod
    def merge_cells(row, column_indexes):
        """
        Объединяет последовательность ячеек строки.

        Например, [2, 3, 4] объединяет столбцы с индексами 2–4.
        Индексы начинаются с нуля.
        """
        if not column_indexes:
            raise ValueError("Список столбцов для объединения пуст.")

        indexes = sorted(set(column_indexes))

        if indexes[0] < 0 or indexes[-1] >= len(row.cells):
            raise IndexError(
                "Индекс столбца для объединения выходит за границы таблицы."
            )

        expected = list(range(indexes[0], indexes[-1] + 1))

        if indexes != expected:
            raise ValueError(
                "Можно объединять только соседние столбцы."
            )

        if len(indexes) == 1:
            return row.cells[indexes[0]]

        return row.cells[indexes[0]].merge(
            row.cells[indexes[-1]]
        )


    # ---------------------------------------------------------
    # ПОДГОТОВКА ДАННЫХ
    # ---------------------------------------------------------
    def prepare_rows(self, records):
        """Группирует ПО по ПК и подготавливает строки таблицы."""

        if isinstance(records, dict):
            normalized_records = []

            for equipment_id, packages in records.items():
                for package in packages:
                    normalized_records.append({
                        **package,
                        "equipment_id": equipment_id,
                    })

            records = normalized_records
        else:
            records = list(records)

        data_config = self.config.get("data", {})
        group_by = data_config.get("group_by", "equipment_id")
        name_field = data_config.get("aggregate_field", "name")
        separator = data_config.get("aggregate_separator", ", ")
        empty_value = data_config.get("empty_value", "")

        additional_category = "Дополнительное ПО"
        computers = OrderedDict()

        # Группируем программы по компьютерам.
        for package in records:
            equipment_id = package.get(group_by, empty_value)

            name = str(package.get(name_field, "") or "").strip()

            if not name:
                continue

            # Удаляем архитектурный суффикс из названия пакета.
            name = re.sub(
                r"\.(?:x86-64|x86_64|x86-32|i[3-6]86|noarch|"
                r"aarch64|arm64|ppc64le|ppc64|s390x)$",
                "",
                name,
                flags=re.IGNORECASE,
            )

            # Добавляем версию после названия через пробел.
            version_field = self.config.get("data", {}).get(
                "version_field", "version"
            )
            version = str(package.get(version_field, "") or "").strip()

            if version:
                name = f"{name} {version}"

            computer = computers.setdefault(
                equipment_id,
                {
                    "classified": [],
                    "additional": [],
                    "additional_seen": set(),
                },
            )

            category = package.get("category", additional_category)
            purpose = package.get("purpose", "")

            if category == additional_category:
                normalized_name = name.casefold()

                if normalized_name not in computer["additional_seen"]:
                    computer["additional_seen"].add(normalized_name)
                    computer["additional"].append(name)
            else:
                computer["classified"].append({
                    "software": name,
                    "software_type": category,
                    "purpose": purpose,
                })

        result = []

        # Номер назначается один раз на каждый ПК.
        for pc_number, (equipment_id, computer) in enumerate(
            computers.items(), start=1
        ):
            computer_rows = list(computer["classified"])

            # Всё дополнительное ПО — одной строкой.
            if computer["additional"]:
                computer_rows.append({
                    "software": (
                        f"{additional_category}: "
                        + separator.join(computer["additional"])
                    ),
                    "software_type": "",
                    "purpose": "",
                    "_merge_additional_columns": True,
                })

            for row_data in computer_rows:
                result.append({
                    "_pc_number": pc_number,
                    "equipment_id": equipment_id,
                    **row_data,
                })

        return result


    # ---------------------------------------------------------
    # ДОБАВЛЕНИЕ СТРОК
    # ---------------------------------------------------------
    def add_rows(self, records):
        """Добавляет строки и объединяет ячейки по компьютерам.

        self.prepare=True  — подготовка строк ПО через prepare_rows().
        self.prepare=False — запись готовых данных, например характеристик железа.
        """

        if self.table is None:
            raise RuntimeError(
                "Сначала необходимо вызвать create_table()."
            )

        table_config = self.config["table"]
        columns = table_config["columns"]
        merge_config = table_config.get("merge", {})
        data_merge_columns = merge_config.get("data_columns", [])
        empty_value = self.config.get("data", {}).get("empty_value", "")

        # ПО требует предварительной подготовки, железо — нет.
        if self.prepare:
            prepared_records = self.prepare_rows(records)
        else:
            prepared_records = list(records)

        # Индексы нужных столбцов.
        column_indexes = {
            column.get("key"): index
            for index, column in enumerate(columns)
        }

        # Добавляем строки.
        for row_number, record in enumerate(prepared_records, start=1):
            row = self.table.add_row()

            for index, column in enumerate(columns):
                key = column.get("key")

                if key == "number":
                    value = record.get(
                        "_pc_number",
                        row_number,
                    )
                else:
                    value = record.get(key, empty_value)

                if value is None:
                    value = empty_value

                cell = row.cells[index]
                cell.text = str(value)

                self.configure_cell(
                    cell,
                    alignment=column.get("align", "left"),
                    bold=column.get("bold", False),
                    font_size=column.get("font_size_pt"),
                    font_name=column.get("font_name"),
                )

                width_cm = column.get("width_cm")
                if width_cm is not None:
                    cell.width = Cm(width_cm)

            # Горизонтальное объединение из JSON.
            # Для железа не применяем настройки объединения таблицы ПО.
            if self.prepare and data_merge_columns:
                merged_cell = self.merge_cells(
                    row, data_merge_columns
                )

                first_index = min(data_merge_columns)
                first_column = columns[first_index]
                first_key = first_column.get("key")

                if first_key == "number":
                    merged_value = record.get("_pc_number", "")
                else:
                    merged_value = record.get(first_key, empty_value)

                merged_cell.text = str(
                    merged_value
                    if merged_value is not None
                    else empty_value
                )

                self.configure_cell(
                    merged_cell,
                    alignment=first_column.get("align", "left"),
                    bold=first_column.get("bold", False),
                    font_size=first_column.get("font_size_pt"),
                    font_name=first_column.get("font_name"),
                )

            # Объединяем столбцы только у строки дополнительного ПО.
            if self.prepare and record.get("_merge_additional_columns"):
                required_keys = (
                    "software",
                    "software_type",
                    "purpose",
                )

                if all(key in column_indexes for key in required_keys):
                    indexes = sorted(
                        column_indexes[key]
                        for key in required_keys
                    )

                    # Объединять можно только соседние столбцы.
                    if indexes == list(
                        range(indexes[0], indexes[-1] + 1)
                    ):
                        merged_cell = self.merge_cells(row, indexes)
                        merged_cell.text = record.get("software", "")

                        software_column = columns[
                            column_indexes["software"]
                        ]

                        self.configure_cell(
                            merged_cell,
                            alignment=software_column.get("align", "left"),
                            bold=software_column.get("bold", False),
                            font_size=software_column.get("font_size_pt"),
                            font_name=software_column.get("font_name"),
                        )

            if table_config.get("allow_row_split", False) is False:
                self.set_row_cant_split(row)

        # Вертикальное объединение номера и названия ПК
        # выполняем только для подготовленных строк ПО.
        if self.prepare:
            start = 0

            while start < len(prepared_records):
                pc_number = prepared_records[start].get("_pc_number")
                end = start + 1

                # Находим строки одного ПК.
                while (
                    end < len(prepared_records)
                    and prepared_records[end].get("_pc_number") == pc_number
                ):
                    end += 1

                # +1, поскольку первая строка таблицы — заголовок.
                if end - start > 1:
                    for key in ("number", "equipment_id"):
                        if key not in column_indexes:
                            continue

                        col_index = column_indexes[key]

                        top_cell = self.table.cell(start + 1, col_index)
                        bottom_cell = self.table.cell(end, col_index)

                        value = str(
                            prepared_records[start].get(
                                "_pc_number"
                                if key == "number"
                                else "equipment_id",
                                "",
                            ) or ""
                        )

                        top_cell.text = value
                        bottom_cell.text = ""

                        top_cell.merge(bottom_cell)
                        top_cell.text = value

                        column = columns[col_index]
                        self.configure_cell(
                            top_cell,
                            alignment=column.get("align", "left"),
                            bold=column.get("bold", False),
                            font_size=column.get("font_size_pt"),
                            font_name=column.get("font_name"),
                        )

                start = end

        return len(prepared_records)


    # ---------------------------------------------------------
    # Проверка имени файла
    # ---------------------------------------------------------
    @staticmethod
    def input_filename(output_dir: str | Path, filename: str, default_name: str) -> Path:
        """Запрашивает и проверяет имя итогового Word-файла.

        Args:
            output_dir (str | Path): Директория для сохранения файла.
            filename (str): Предлагаемое имя файла пользователем.
            default_name (str): Имя файла по умолчанию, если filename пустое.

        Returns:
            Path: Полный путь к итоговому Word-файлу.
        """

        output_dir = Path(output_dir)

        # Зарезервированные имена Windows: CON, NUL, COM1 и т. д.
        reserved_names = {
            "CON", "PRN", "AUX", "NUL",
            *(f"COM{i}" for i in range(1, 10)),
            *(f"LPT{i}" for i in range(1, 10)),
        }

        # Символы, запрещённые в именах файлов Windows.
        invalid_chars = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

        while True:
            if not filename:
                filename = default_name
                print(f"Используется имя по умолчанию: {filename}")

            # Не допускаем имена с пробелом или точкой в конце.
            if filename.endswith((" ", ".")):
                print("Ошибка: имя файла не должно заканчиваться пробелом или точкой.")
                continue

            if invalid_chars.search(filename):
                print('Ошибка: имя содержит запрещённые символы: <>:"/\\|?*')
                continue

            # Если расширение не указано — добавляем .docx.
            if filename.lower().endswith(".docx"):
                filename = filename[:-5]
            elif "." in filename:
                print("Ошибка: используйте расширение .docx.")
                continue

            if not filename:
                filename = default_name
                print(f"Используется имя по умолчанию: {filename}")

            # Windows считает зарезервированными также имена вида CON.txt.
            if filename.split(".")[0].upper() in reserved_names:
                print("Ошибка: это зарезервированное имя Windows.")
                continue

            filename += ".docx"
            output_path = output_dir / filename

            if output_path.exists():
                overwrite = input(
                    f"Файл «{filename}» уже существует. "
                    "Перезаписать? (Д/н): "
                ).strip().lower()

                if overwrite not in {"д", "да", "y", "yes", ""}:
                    print("Введите другое имя файла.")
                    continue

            return output_path


    # ---------------------------------------------------------
    # СОХРАНЕНИЕ
    # ---------------------------------------------------------
    def save(self):
        """Сохраняет документ в указанный файл."""
        self.output_path.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        self.document.save(self.output_path)

        return self.output_path


    def build(self, records):
        """
        Полный цикл создания документа.

        Возвращает количество строк данных.
        """
        self.create_document()
        self.create_table()

        row_count = self.add_rows(records)

        saved_path = self.save()

        print(f"Документ сохранён: {saved_path}")
        print(f"Количество строк данных: {row_count}")

        return row_count