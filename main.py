from app.menu import run_menu


if __name__ == "__main__":
    try:
        run_menu()
    except PermissionError as error:
        print(f"\nОШИБКА, файл занят другим процессом: {error}")
        raise SystemExit(1)
    finally:
        input("\nНажмите Enter для выхода...")