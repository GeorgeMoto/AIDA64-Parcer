import tkinter as tk
from tkinter import filedialog, messagebox, StringVar
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import os
import threading
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='aida_parser.log',
    filemode='a'
)


class AidaParserUI:
    def __init__(self, root, process_callback):
        self.root = root
        self.root.title("Парсер отчетов AIDA64")
        self.root.geometry("600x650")
        self.root.resizable(True, True)

        # Store the callback function for processing
        self.process_callback = process_callback

        # Variables for tracking progress
        self.progress_var = ttk.DoubleVar()
        self.status_var = StringVar(value="Готов к работе")

        # Input/output path variables
        self.input_dir_var = StringVar()
        self.output_file_var = StringVar()

        # Create UI
        self.create_ui()

    def create_ui(self):
        """Creates the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=BOTH, expand=YES)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Парсер отчетов AIDA64",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))

        # Directory selection section
        dir_frame = ttk.LabelFrame(main_frame, text="Выбор директории", padding=10)
        dir_frame.pack(fill=X, pady=10)

        # Input directory
        input_dir_label = ttk.Label(dir_frame, text="Директория с отчетами:")
        input_dir_label.grid(row=0, column=0, sticky=W, pady=5)

        input_dir_entry = ttk.Entry(dir_frame, textvariable=self.input_dir_var, width=50)
        input_dir_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        input_dir_button = ttk.Button(
            dir_frame,
            text="Обзор...",
            command=self.select_input_dir,
            width=10
        )
        input_dir_button.grid(row=0, column=2, padx=5, pady=5)

        # Output file
        output_file_label = ttk.Label(dir_frame, text="Файл отчета:")
        output_file_label.grid(row=1, column=0, sticky=W, pady=5)

        output_file_entry = ttk.Entry(dir_frame, textvariable=self.output_file_var, width=50)
        output_file_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        output_file_button = ttk.Button(
            dir_frame,
            text="Обзор...",
            command=self.select_output_file,
            width=10
        )
        output_file_button.grid(row=1, column=2, padx=5, pady=5)

        # Configure grid
        dir_frame.columnconfigure(1, weight=1)

        # Instructions section
        info_frame = ttk.LabelFrame(main_frame, text="Информация", padding=10)
        info_frame.pack(fill=X, pady=10)

        info_text = (
            "Программа анализирует HTML-отчеты AIDA64 и создает Excel-таблицу с результатами.\n\n"
            "1. Выберите директорию с HTML-отчетами (файлы .htm или .html)\n"
            "2. Укажите путь для сохранения Excel-файла с результатами\n"
            "3. Нажмите 'Начать обработку' для запуска процесса\n\n"
            "Отчет будет содержать информацию о имени ПК, операционной системе,\n"
            "установленном программном обеспечении и списке пользователей."
        )

        info_label = ttk.Label(
            info_frame,
            text=info_text,
            justify=LEFT,
            wraplength=580
        )
        info_label.pack(fill=X, padx=5, pady=5)

        # Process button
        process_button = ttk.Button(
            main_frame,
            text="Начать обработку",
            command=self.start_processing,
            bootstyle="success",
            width=20
        )
        process_button.pack(pady=20)

        # Status bar
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=X, pady=(10, 0))

        # Progress bar
        self.progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            mode="determinate",
            bootstyle="success"
        )
        self.progress_bar.pack(fill=X, side=TOP, pady=(0, 5))

        # Status text
        status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Arial", 10)
        )
        status_label.pack(side=LEFT, padx=5)

    def select_input_dir(self):
        """Opens file dialog to select input directory."""
        directory = filedialog.askdirectory(title="Выберите директорию с отчетами AIDA64")
        if directory:
            self.input_dir_var.set(directory)
            # Set default output file location in the same directory
            if not self.output_file_var.get():
                self.output_file_var.set(os.path.join(directory, "results.xlsx"))

    def select_output_file(self):
        """Opens file dialog to select output file."""
        file_path = filedialog.asksaveasfilename(
            title="Сохранить отчет как",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.output_file_var.set(file_path)

    def start_processing(self):
        """Starts the processing of HTML files."""
        input_dir = self.input_dir_var.get()
        output_file = self.output_file_var.get()

        # Validate inputs
        if not input_dir:
            messagebox.showerror("Ошибка", "Выберите директорию с отчетами")
            return

        if not output_file:
            messagebox.showerror("Ошибка", "Укажите файл для сохранения результатов")
            return

        if not os.path.isdir(input_dir):
            messagebox.showerror("Ошибка", f"Директория '{input_dir}' не существует")
            return

        # Check if we can create the output file
        try:
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Ошибка", f"Невозможно создать файл отчета: {e}")
            return

        # Reset progress
        self.progress_var.set(0)
        self.status_var.set("Подготовка к обработке...")

        # Start processing in a separate thread
        threading.Thread(
            target=self.process_thread,
            args=(input_dir, output_file),
            daemon=True
        ).start()

    def process_thread(self, input_dir, output_file):
        """Processing thread to avoid UI freezing."""
        try:
            # Call the callback with a progress update function
            self.process_callback(
                input_dir,
                output_file,
                self.update_progress
            )
            # Processing completed successfully
            self.root.after(0, lambda: self.complete_processing(True))
        except Exception as e:
            logging.error(f"Error processing files: {e}", exc_info=True)
            # Processing failed
            self.root.after(0, lambda: self.complete_processing(False, str(e)))

    def update_progress(self, current, total, message=None):
        """Updates progress bar and status message."""
        progress = (current / total) * 100 if total > 0 else 0

        def update_ui():
            self.progress_var.set(progress)
            if message:
                self.status_var.set(message)
            else:
                self.status_var.set(f"Обработано {current} из {total} файлов ({int(progress)}%)")

        self.root.after(0, update_ui)

    def complete_processing(self, success, error_message=None):
        """Updates UI after processing is complete."""
        if success:
            self.status_var.set("Обработка завершена успешно")
            messagebox.showinfo("Успех", f"Отчет успешно сохранен:\n{self.output_file_var.get()}")
        else:
            self.status_var.set("Ошибка при обработке")
            messagebox.showerror("Ошибка", f"Не удалось обработать файлы:\n{error_message}")

        # Reset progress bar
        self.progress_var.set(0)