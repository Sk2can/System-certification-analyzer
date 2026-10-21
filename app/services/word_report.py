import json
from collections import OrderedDict
from pathlib import Path
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
        """
        Подготавливает записи для таблицы.

        Если задан group_by, группирует записи по этому полю
        и объединяет значения aggregate_field в одну строку.
        """


    def prepare_rows(self, records):
        """Подготавливает записи для таблицы Word."""

        # Преобразуем dict[str, list[dict]] в list[dict].
        if isinstance(records, dict):
            normalized_records = []

            for equipment_id, packages in records.items():
                for package in packages:
                    normalized_records.append({
                        **package,
                        "equipment_id": equipment_id,
                    })

            records = normalized_records

        data_config = self.config.get("data", {})
        group_by = data_config.get("group_by")

        if not group_by:
            return list(records)


        data_config = self.config.get("data", {})

        group_by = data_config.get("group_by")

        if not group_by:
            return list(records)

        aggregate_field = data_config.get(
            "aggregate_field", "name"
        )
        output_field = data_config.get(
            "aggregate_output_field", "software"
        )
        separator = data_config.get(
            "aggregate_separator", ", "
        )
        prefix = data_config.get("aggregate_prefix", "")
        empty_value = data_config.get("empty_value", "")

        include_version = data_config.get(
            "include_version", False
        )
        version_field = data_config.get(
            "version_field", "version"
        )

        deduplicate = data_config.get("deduplicate", True)
        case_sensitive = data_config.get(
            "deduplicate_case_sensitive", False
        )

        groups = OrderedDict()

        for record in records:
            group_value = record.get(group_by, empty_value)

            if group_value not in groups:
                groups[group_value] = {
                    "record": dict(record),
                    "values": [],
                    "seen": set(),
                }

            value = record.get(aggregate_field, empty_value)

            if value is None:
                value = empty_value

            value = str(value).strip()

            if not value:
                continue

            if include_version:
                version = record.get(version_field, empty_value)

                if version is not None and str(version).strip():
                    value = f"{value} {str(version).strip()}"

            comparison_value = (
                value if case_sensitive else value.casefold()
            )

            group = groups[group_value]

            if deduplicate and comparison_value in group["seen"]:
                continue

            group["seen"].add(comparison_value)
            group["values"].append(value)

        result = []

        for group in groups.values():
            record = group["record"]

            aggregated = separator.join(group["values"])

            if aggregated:
                aggregated = prefix + aggregated

            record[output_field] = aggregated

            result.append(record)

        return result

    # ---------------------------------------------------------
    # ДОБАВЛЕНИЕ СТРОК
    # ---------------------------------------------------------

    def add_rows(self, records):
        """Добавляет данные в таблицу и возвращает число строк."""
        if self.table is None:
            raise RuntimeError(
                "Сначала необходимо вызвать create_table()."
            )

        table_config = self.config["table"]
        columns = table_config["columns"]
        merge_config = table_config.get("merge", {})
        data_merge_columns = merge_config.get("data_columns", [])

        data_config = self.config.get("data", {})
        empty_value = data_config.get("empty_value", "")

        prepared_records = self.prepare_rows(records)

        for row_number, record in enumerate(prepared_records, start=1):
            row = self.table.add_row()

            # Сначала заполняем ячейки до объединения.
            for index, column in enumerate(columns):
                key = column.get("key")
                value = row_number if key == "number" else record.get(
                    key, empty_value
                )

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

            # Объединяем ячейки строки данных, если это настроено.
            if data_merge_columns:
                merged_cell = self.merge_cells(
                    row, data_merge_columns
                )

                # В объединённой ячейке остаётся значение
                # первого столбца из группы объединения.
                first_index = min(data_merge_columns)
                first_column = columns[first_index]
                first_key = first_column.get("key")

                merged_value = (
                    row_number
                    if first_key == "number"
                    else record.get(first_key, empty_value)
                )

                if merged_value is None:
                    merged_value = empty_value

                merged_cell.text = str(merged_value)

                self.configure_cell(
                    merged_cell,
                    alignment=first_column.get("align", "left"),
                    bold=first_column.get("bold", False),
                    font_size=first_column.get("font_size_pt"),
                    font_name=first_column.get("font_name"),
                )

            if table_config.get("allow_row_split", False) is False:
                self.set_row_cant_split(row)

        return len(prepared_records)

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