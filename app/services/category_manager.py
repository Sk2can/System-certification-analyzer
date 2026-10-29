import json
import os
import sys
from pathlib import Path

# Путь к исходному JSON, включённому в приложение.
if getattr(sys, "frozen", False):
    RESOURCE_DIR = Path(sys._MEIPASS) / "app" / "settings"
else:
    RESOURCE_DIR = Path(__file__).resolve().parent.parent / "settings"

DEFAULT_CONFIG_PATH = RESOURCE_DIR / "software_categories.json"

# Постоянное хранилище пользовательских настроек.
APP_DATA_DIR = Path(
    os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")
) / "SCA"

CONFIG_PATH = (
    Path(os.environ["APPDATA"])
    / "SCA"
    / "software_categories.json"
)

FIELDS = ("name", "match", "category", "purpose")


def load_categories() -> list[dict]:
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)

    if not CONFIG_PATH.exists():
        if not DEFAULT_CONFIG_PATH.is_file():
            raise FileNotFoundError(
                f"Не найден файл шаблона категорий: {DEFAULT_CONFIG_PATH}"
            )

        CONFIG_PATH.write_text(
            DEFAULT_CONFIG_PATH.read_text(encoding="utf-8"),
            encoding="utf-8",
        )

    with CONFIG_PATH.open("r", encoding="utf-8") as file:
        data = json.load(file)

    if not isinstance(data, list):
        raise ValueError("Корень JSON должен быть списком записей.")

    for index, item in enumerate(data, start=1):
        if not isinstance(item, dict):
            raise ValueError(f"Запись №{index} должна быть объектом.")

        if not all(field in item for field in FIELDS):
            raise ValueError(
                f"Запись №{index} должна содержать поля: "
                f"{', '.join(FIELDS)}"
            )

        for field in ("name", "category", "purpose"):
            if not isinstance(item[field], str) or not item[field].strip():
                raise ValueError(
                    f"Поле '{field}' в записи №{index} должно быть непустой строкой."
                )

        matches = item["match"]
        if (
            not isinstance(matches, list)
            or not matches
            or not all(
                isinstance(value, str) and value.strip()
                for value in matches
            )
        ):
            raise ValueError(
                f"Поле 'match' в записи №{index} должно содержать "
                "список непустых строк."
            )

    return data


def save_categories(records: list[dict]) -> None:
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = CONFIG_PATH.with_suffix(".json.tmp")

    try:
        with temp_path.open("w", encoding="utf-8") as file:
            json.dump(records, file, ensure_ascii=False, indent=2)
            file.write("\n")

        temp_path.replace(CONFIG_PATH)
    finally:
        if temp_path.exists():
            temp_path.unlink()

def input_nonempty(prompt: str) -> str:
    """Запрашивает непустую строку."""

    while True:
        value = input(prompt).strip()

        if value:
            return value

        print("Значение не может быть пустым.")


def select_record(records: list[dict]) -> int | None:
    """Выбирает запись по номеру."""

    if not records:
        print("Список записей пуст.")
        return None

    for index, item in enumerate(records, start=1):
        print(
            f"{index}. {item['name']} "
            f"[{item['category']}]"
        )

    choice = input("Введите номер записи: ").strip()

    if not choice.isdigit():
        print("Введите числовой номер.")
        return None

    index = int(choice) - 1

    if not 0 <= index < len(records):
        print("Записи с таким номером нет.")
        return None

    return index


def input_matches() -> list[str]:
    """Принимает ключевые слова через запятую."""

    while True:
        raw = input(
            "Ключевые слова match через запятую: "
        ).strip()

        matches = list(dict.fromkeys(
            value.strip()
            for value in raw.split(",")
            if value.strip()
        ))

        if matches:
            return matches

        print("Нужно указать хотя бы одно ключевое слово.")


def add_record(records: list[dict]) -> None:
    """Добавляет новую запись."""

    name = input_nonempty("Название ПО: ")

    if any(item["name"].casefold() == name.casefold() for item in records):
        print("Запись с таким названием уже существует.")
        return

    matches = input_matches()
    category = input_nonempty("Категория: ")
    purpose = input_nonempty("Назначение: ")

    records.append({
        "name": name,
        "match": matches,
        "category": category,
        "purpose": purpose,
    })

    save_categories(records)
    print("Запись добавлена.")


def edit_record(records: list[dict]) -> None:
    """Изменяет поля существующей записи."""

    index = select_record(records)
    if index is None:
        return

    item = records[index]

    print("\nОставь поле пустым, чтобы не менять его.")

    name = input(f"Название [{item['name']}]: ").strip()
    raw_matches = input(
        f"Ключевые слова [{', '.join(item['match'])}]: "
    ).strip()
    category = input(f"Категория [{item['category']}]: ").strip()
    purpose = input(f"Назначение [{item['purpose']}]: ").strip()

    new_name = name or item["name"]

    if any(
        other["name"].casefold() == new_name.casefold()
        for i, other in enumerate(records)
        if i != index
    ):
        print("Другая запись уже имеет такое название.")
        return

    new_matches = (
        list(dict.fromkeys(
            value.strip()
            for value in raw_matches.split(",")
            if value.strip()
        ))
        if raw_matches
        else item["match"]
    )

    if raw_matches and not new_matches:
        print("Список ключевых слов не может быть пустым.")
        return

    item.update({
        "name": new_name,
        "match": new_matches,
        "category": category or item["category"],
        "purpose": purpose or item["purpose"],
    })

    save_categories(records)
    print("Запись изменена.")


def delete_record(records: list[dict]) -> None:
    """Удаляет запись с подтверждением."""

    index = select_record(records)
    if index is None:
        return

    item = records[index]
    confirm = input(
        f"Удалить «{item['name']}»? (д/н): "
    ).strip().casefold()

    if confirm not in {"д", "да", "y", "yes"}:
        print("Удаление отменено.")
        return

    removed = records.pop(index)
    save_categories(records)

    print(f"Запись «{removed['name']}» удалена.")


def show_records(records: list[dict]) -> None:
    """Выводит все записи."""

    if not records:
        print("Список записей пуст.")
        return

    for index, item in enumerate(records, start=1):
        print(f"\n{index}. {item['name']}")
        print(f"   Ключевые слова: {', '.join(item['match'])}")
        print(f"   Категория: {item['category']}")
        print(f"   Назначение: {item['purpose']}")


def run_category_manager() -> None:
    """Меню управления записями о ПО."""

    actions = {
        "1": ("Просмотреть записи", show_records),
        "2": ("Добавить запись", add_record),
        "3": ("Изменить запись", edit_record),
        "4": ("Удалить запись", delete_record),
    }

    try:
        while True:
            print("\n=== Управление категориями ПО ===")

            for key, (title, _) in actions.items():
                print(f"{key}. {title}")

            print("0. Назад")

            choice = input("Выберите действие: ").strip()

            if choice == "0":
                return

            if choice not in actions:
                print("Выберите пункт от 0 до 4.")
                continue

            try:
                records = load_categories()
                _, action = actions[choice]
                action(records)

            except (OSError, ValueError, json.JSONDecodeError) as error:
                print(f"Ошибка: {error}")

            input("\nНажмите Enter для продолжения.")

    except (KeyboardInterrupt, EOFError):
        print("\nВыход из управления категориями.")
