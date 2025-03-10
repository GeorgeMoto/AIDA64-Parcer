import logging
import ttkbootstrap as ttk
from ui import AidaParserUI
from html_parser import process_directory

def main():
    """Основная функция программы."""
    # Настройка логирования
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        filename='aida_parser.log',
        filemode='a'
    )

    # Создаем главное окно
    root = ttk.Window(themename="superhero")  # Темная тема
    root.title("Парсер отчетов AIDA64")

    # Запуск UI с функцией обратного вызова
    app = AidaParserUI(root, process_directory)

    # Запуск основного цикла приложения
    root.mainloop()


if __name__ == "__main__":
    main()