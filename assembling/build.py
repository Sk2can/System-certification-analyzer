import os
import shutil
import subprocess
import sys
from pathlib import Path


# Пути проекта
ASSEMBLING_DIR = Path(__file__).resolve().parent
PROJECT_DIR = ASSEMBLING_DIR.parent

PYTHON = Path(sys.executable)
MAIN = PROJECT_DIR / "main.py"
SETTINGS = PROJECT_DIR / "app" / "settings"

WORK_DIR = ASSEMBLING_DIR / "build"
DIST_DIR = ASSEMBLING_DIR / "dist"
SPEC_FILE = ASSEMBLING_DIR / "SCA.spec"
ICON = ASSEMBLING_DIR / "icon.ico"


def check_paths() -> bool:
    """Проверяет наличие необходимых файлов."""

    required_paths = {
        "Главный файл": MAIN,
        "Каталог настроек": SETTINGS,
        "Файл настроек": SETTINGS / "software_table.json",
    }

    for name, path in required_paths.items():
        if not path.exists():
            print(f"ОШИБКА: {name} не найден: {path}")
            return False

    print(f"Корень проекта: {PROJECT_DIR}")
    print(f"Python: {PYTHON}")
    print(f"Главный файл: {MAIN}")
    print(f"Настройки: {SETTINGS}")

    return True


def clean_build() -> None:
    """Удаляет результаты предыдущей сборки."""

    for path in (WORK_DIR, DIST_DIR):
        if path.exists():
            shutil.rmtree(path)

    if SPEC_FILE.exists():
        SPEC_FILE.unlink()

    WORK_DIR.mkdir(parents=True, exist_ok=True)


def build_exe() -> int:
    """Запускает PyInstaller с явным списком аргументов."""

    # Удаляем предыдущую сборку
    exe_path = DIST_DIR / "SCA.exe"

    if exe_path.exists():
        try:
            exe_path.unlink()
            print(f"Удалён старый файл: {exe_path}")
        except PermissionError:
            print(
                "Ошибка: не удалось удалить старый SCA.exe. "
                "Закрой приложение и попробуй снова."
            )
            return 1

    
    # Разделитель ресурсов соответствует текущей ОС.
    data_argument = f"{PROJECT_DIR / 'app'}{os.pathsep}app"

    command = [
        str(PYTHON),
        "-m",
        "PyInstaller",
        str(MAIN),  # scriptname передаётся отдельным аргументом
        "--clean",
        "--noconfirm",
        "--onefile",
        "--icon",
        str(ICON),
        "--console",
        "--name",
        "SCA",
        "--workpath",
        str(WORK_DIR),
        "--distpath",
        str(DIST_DIR),
        "--specpath",
        str(ASSEMBLING_DIR),
        "--add-data",
        data_argument,
    ]

    print("\nКоманда сборки:")
    print(subprocess.list2cmdline(command))
    print("\nВыполняется сборка...\n")

    result = subprocess.run(
        command,
        cwd=PROJECT_DIR,
        check=False,
    )

    if result.returncode != 0:
        print("\nОШИБКА: PyInstaller завершился с ошибкой.")
        return result.returncode

    exe_path = DIST_DIR / "SCA.exe"

    if not exe_path.exists():
        print(f"\nОШИБКА: EXE не найден: {exe_path}")
        return 1

    print("\nСБОРКА УСПЕШНО ЗАВЕРШЕНА")
    print(f"Готовый файл: {exe_path}")

    return 0


def main() -> int:
    print("=" * 45)
    print("СБОРКА ПРИЛОЖЕНИЯ SCA")
    print("=" * 45)

    if not check_paths():
        return 1

    try:
        clean_build()
        return build_exe()
    except Exception as error:
        print(f"\nОШИБКА: {type(error).__name__}: {error}")
        return 1


if __name__ == "__main__":
    sys.exit(main())