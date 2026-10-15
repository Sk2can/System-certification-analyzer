from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT


class SoftwareTableBuilder:
    """Формирование общей Word-таблицы установленного ПО."""

    HEADERS = [
        "№ п/п",
        "Наименование ОТСС (идентификатор)",
        "Наименование ПО",
        "ПО (его тип)",
        "Назначение",
    ]

    def __init__(self, input_dir: str | Path, output_file: str | Path):
        self.input_dir = Path(input_dir)
        self.output_file = Path(output_file)

        self.scans = {}
        self.document = None
        self.table = None

    # --------------------------------------------------
    # 1. Поиск TXT-файлов
    # --------------------------------------------------
    def get_scan_files(self) -> list[Path]:
        """Возвращает список TXT-файлов в каталоге."""

        if not self.input_dir.is_dir():
            raise NotADirectoryError(
                f"Каталог не найден: {self.input_dir}"
            )

        files = sorted(
            self.input_dir.glob("*.txt"),
            key=lambda path: path.stat().st_mtime,
        )

        if not files:
            raise FileNotFoundError(
                f"В каталоге нет TXT-файлов: {self.input_dir}"
            )

        return files

    # --------------------------------------------------
    # 2. Чтение TXT-файла
    # --------------------------------------------------
    def read_scan_file(self, file_path: Path) -> tuple[str, list[dict]]:
        """Извлекает идентификатор оборудования и установленные пакеты."""

        with file_path.open(
            "r",
            encoding="utf-8-sig",
            errors="replace",
        ) as file:
            lines = file.readlines()

        if not lines or not lines[0].strip():
            raise ValueError(
                f"Не указан идентификатор оборудования: {file_path}"
            )

        equipment_id = lines[0].strip()

        packages = self.parse_packages(lines)

        return equipment_id, packages

    # --------------------------------------------------
    # 3. Разбор блока установленных пакетов
    # --------------------------------------------------
    def parse_packages(self, lines: list[str]) -> list[dict]:
        """Извлекает название, версию и репозиторий пакетов."""

        start_index = next(
            (
                index
                for index, line in enumerate(lines)
                if "Установленные пакеты" in line
            ),
            None,
        )

        if start_index is None:
            raise ValueError("Блок 'Установленные пакеты' не найден")

        packages = []

        for line in lines[start_index + 1:]:
            line = line.strip()

            if not line:
                continue

            if line.startswith("*") and set(line) == {"*"}:
                if packages:
                    break
                continue

            parts = line.split()

            if len(parts) < 2:
                continue

            packages.append({
                "name": parts[0],
                "version": parts[1],
                "repository": parts[2] if len(parts) > 2 else "",
            })

        return packages

    # --------------------------------------------------
    # 4. Сбор результатов сканирования
    # --------------------------------------------------
    def collect_scans(self) -> dict[str, list[dict]]:
        """
        Собирает результаты по оборудованию.

        Если для одного оборудования найдено несколько файлов,
        используется файл с наиболее поздним временем изменения.
        """

        self.scans = {}

        for file_path in self.get_scan_files():
            try:
                equipment_id, packages = self.read_scan_file(file_path)
            except (OSError, ValueError) as error:
                print(f"Пропущен файл {file_path.name}: {error}")
                continue

            self.scans[equipment_id] = packages

        if not self.scans:
            raise ValueError(
                "Не удалось обработать ни одного TXT-файла"
            )

        return self.scans

    # --------------------------------------------------
    # 5. Создание документа и заголовка
    # --------------------------------------------------
    def create_document(self) -> None:
        """Создаёт документ без лишних разрывов страниц."""

        self.document = Document()
        self.configure_page()

        title = self.document.add_heading(
            "Сведения об установленном ПО",
            level=1,
        )
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Заголовок не должен переноситься отдельно от таблицы.
        title.paragraph_format.keep_with_next = True
        title.paragraph_format.space_before = Pt(0)
        title.paragraph_format.space_after = Pt(6)

        for run in title.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)

    # --------------------------------------------------
    # 6. Создание таблицы и заголовков столбцов
    # --------------------------------------------------
    def create_table(self) -> None:
        """Создаёт таблицу с повторяемой строкой заголовков."""

        if self.document is None:
            raise RuntimeError("Документ не создан")

        self.table = self.document.add_table(
            rows=1,
            cols=5,
        )
        self.table.style = "Table Grid"
        self.table.autofit = False

        widths = [Cm(1), Cm(5), Cm(8), Cm(6), Cm(6.7)]

        for column, width in zip(self.table.columns, widths):
            column.width = width

        header_cells = self.table.rows[0].cells

        for cell, header, width in zip(
            header_cells,
            self.HEADERS,
            widths,
        ):
            cell.width = width
            cell.text = header
            self.configure_cell(cell)

            for paragraph in cell.paragraphs:
                self.configure_paragraph(
                    paragraph,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER,
                    bold=True,
                )

        # Заголовок таблицы повторяется при переходе на новую страницу.
        tr_pr = self.table.rows[0]._tr.get_or_add_trPr()
        repeat_header = OxmlElement("w:tblHeader")
        repeat_header.set(qn("w:val"), "true")
        tr_pr.append(repeat_header)

        # Не разрешаем Word принудительно растягивать столбцы.
        tbl_pr = self.table._tbl.tblPr
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tbl_pr.append(layout)

    # --------------------------------------------------
    # 8. Заполнение таблицы
    # --------------------------------------------------
    def fill_table(self) -> int:
        """Заполняет таблицу и выводит статистику обработки."""

        if self.table is None:
            raise RuntimeError("Таблица не создана")

        row_count = 0
        total_packages = 0

        for equipment_id, packages in self.scans.items():
            if not equipment_id.strip():
                print("Пропущена запись без идентификатора оборудования")
                continue

            self.add_package_row(
                number=row_count + 1,
                equipment_id=equipment_id,
                packages=packages,
            )

            row_count += 1
            total_packages += len({
                package["name"].casefold()
                for package in packages
                if package.get("name", "").strip()
            })

        print(f"Обработано АРМ: {row_count}")
        print(f"Уникальных пакетов: {total_packages}")

        return row_count

    # --------------------------------------------------
    # 9. Сохранение документа
    # --------------------------------------------------
    def save_document(self) -> None:
        """Создаёт каталог результата и сохраняет документ."""

        if self.document is None:
            raise RuntimeError("Документ не создан")

        self.output_file.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.document.save(self.output_file)

    # --------------------------------------------------
    # 10. Основной метод формирования таблицы
    # --------------------------------------------------
    def build(self) -> None:
        """Формирует и сохраняет итоговую таблицу."""

        self.collect_scans()
        self.create_document()
        self.create_table()

        row_count = self.fill_table()

        self.save_document()

        print(f"Документ сохранён: {self.output_file}")
        print(f"Строк в таблице: {row_count}")

    # --------------------------------------------------
    # 11. Настройка страницы и полей документа
    # --------------------------------------------------
    def configure_page(self) -> None:
        """Настраивает альбомную ориентацию и поля страницы."""

        for section in self.document.sections:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Cm(29.7)
            section.page_height = Cm(21.0)

            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)

            section.header_distance = Cm(0.7)
            section.footer_distance = Cm(0.7)

    # --------------------------------------------------
    # 12. Настройка абзацев
    # --------------------------------------------------
    def configure_paragraph(
        self,
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        bold=False,
    ) -> None:
        """Настраивает шрифт, интервалы и выравнивание абзаца."""

        paragraph.alignment = alignment

        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0

        for run in paragraph.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)
            run.bold = bold

            # Задаём шрифт также для кириллицы.
            run_properties = run._element.get_or_add_rPr()
            fonts = run_properties.rFonts

            if fonts is None:
                fonts = OxmlElement("w:rFonts")
                run_properties.insert(0, fonts)

            fonts.set(qn("w:ascii"), "Times New Roman")
            fonts.set(qn("w:hAnsi"), "Times New Roman")
            fonts.set(qn("w:eastAsia"), "Times New Roman")
            fonts.set(qn("w:cs"), "Times New Roman")

    # --------------------------------------------------
    # 13. Настройка ячеек
    # --------------------------------------------------
    def configure_cell(
        self,
        cell,
        top=57,
        start=85,
        bottom=57,
        end=85,
    ) -> None:
        """Задаёт внутренние отступы ячейки в twips."""

        tc_pr = cell._tc.get_or_add_tcPr()

        tc_mar = tc_pr.find(qn("w:tcMar"))

        if tc_mar is None:
            tc_mar = OxmlElement("w:tcMar")
            tc_pr.append(tc_mar)

        margins = {
            "top": top,
            "start": start,
            "bottom": bottom,
            "end": end,
        }

        for margin_name, value in margins.items():
            element = tc_mar.find(qn(f"w:{margin_name}"))

            if element is None:
                element = OxmlElement(f"w:{margin_name}")
                tc_mar.append(element)

            element.set(qn("w:w"), str(value))
            element.set(qn("w:type"), "dxa")

        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # --------------------------------------------------
    # 14. Метод добавления строки с объединёнными ячейками
    # --------------------------------------------------
    def add_package_row(
        self,
        number: int,
        equipment_id: str,
        packages: list[dict],
    ) -> None:
        """Добавляет одну строку на АРМ."""

        cells = self.table.add_row().cells

        cells[0].text = str(number)
        cells[1].text = equipment_id

        # Уникальные пакеты сохраняются в исходном порядке.
        unique_packages = {}
        for package in packages:
            name = package.get("name", "").strip()
            if name:
                unique_packages.setdefault(name.casefold(), name)

        software_text = "Дополнительное ПО: " + ", ".join(
            unique_packages.values()
        )

        # Объединяем три ячейки в одну.
        merged_cell = cells[2].merge(cells[4])
        merged_cell.text = software_text

        for cell in (cells[0], cells[1], merged_cell):
            self.configure_cell(cell)

            for paragraph in cell.paragraphs:
                alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                    if cell is cells[0]
                    else WD_ALIGN_PARAGRAPH.LEFT
                )
                self.configure_paragraph(
                    paragraph,
                    alignment=alignment,
                )

                for run in paragraph.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(9)