# Инструмент по пакетной выгрузке ПДФ
# Импортируем необходимые библиотеки
import requests
import pandas as pd
import concurrent.futures
import PyPDF3
import re
import io
import os
import tkinter as tk
import sys
import getpass
import time
import datetime
from datetime import datetime
import logging
import pickle
import json
from cryptography.fernet import Fernet
from tkinter import filedialog

# Перед компиляцией снять комментарий
#exe_path = os.path.dirname(sys.executable)
exe_path = "E:\\vscode\\dist\\pdf_downloader\\pdf_downloader"
print(f"Рабочая папка: {exe_path}")
os.chdir(exe_path)
print(os.curdir)

start_time = time.time()
# Настраиваем параметры логирования
logging.basicConfig(
    # Указываем имя файла, в который будем записывать логи
    filename="log.log",
    # Указываем формат записи логов
    format="%(asctime)s - %(levelname)s - %(message)s",
    # Указываем уровень логирования
    level=logging.INFO
)
try:
    tool_mode = sys.argv[1]
except IndexError as e:
    while True:
        modes = ["gm","mm","alt"]
        user_input = input("Пожалуйста, укажите режим работы для инструмента: gm, mm или alt:\n")
        if user_input in modes:
            tool_mode = user_input
            break
        else:
            print("Неверный ввод. Попробуйте еще раз")

ShowPos = False
ShowDate = True
SeparateFolders = True

for arg in sys.argv[1:]:
    if arg == 'pos':
        ShowPos = True
    if arg == 'nodate':
        ShowDate = False
    if arg == 'nofolders':
        SeparateFolders = False


if tool_mode == 'gm':
    print(f"Инструмент работает в режиме выгрузки по папкам для ГМ")
    logging.info(f"Инструмент работает в режиме выгрузки по папкам для ГМ")
elif tool_mode == 'mm':
    print(f"Инструмент работает в режиме выгрузки в одну папку для ММ")
    logging.info(f"Инструмент работает в режиме выгрузки в одну папку для ММ")
elif tool_mode == 'alt':
    print(f"Инструмент работает в режиме альтернативной выгрузки в одну папку для ММ")
    logging.info(f"Инструмент работает в режиме альтернативной выгрузки в одну папку для ММ")
else:
    print(f"Режим работы инструмента не выбран, по-умолчанию выбрана загрузка в одну папку")
    logging.info(f"Режим работы инструмента не выбран, по-умолчанию выбрана загрузка в одну папку")
    tool_mode='mm'
global login, password


login = ""
password = ""
oslogin = getpass.getuser()

# Определяем функцию, которая создает и обрабатывает окно с формой входа
def show_login_form():
    # Создаем окно приложения
    window = tk.Tk()
    # Задаем заголовок окна
    window.title("Форма входа")
    # Задаем размер окна
    window.geometry("225x135")
 
    # Создаем метку для логина
    login_label = tk.Label(window, text="Логин:")
    # Размещаем метку в окне
    login_label.grid(row=0, column=0, padx=10, pady=10)
 
    # Создаем поле ввода для логина
    defaulttext = tk.StringVar()
    defaulttext.set(oslogin)
    login_entry = tk.Entry(window, textvariable=defaulttext)
    # Размещаем поле ввода в окне
    login_entry.grid(row=0, column=1, padx=10, pady=10)
 
    # Создаем метку для пароля
    password_label = tk.Label(window, text="Пароль:")
    # Размещаем метку в окне
    password_label.grid(row=1, column=0, padx=10, pady=10)
 
    # Создаем поле ввода для пароля
    password_entry = tk.Entry(window, show="*")
    # Размещаем поле ввода в окне
    password_entry.grid(row=1, column=1, padx=10, pady=10)
 
    # Определяем функцию для обработки нажатия кнопки входа
    def get_login():
        # Получаем введенный логин и пароль
        global login, password
        login = login_entry.get()
        password = password_entry.get()
        window.destroy()
 
    # Создаем кнопку для входа
    login_button = tk.Button(window, text="Войти", command=get_login)
    # Размещаем кнопку в окне
    login_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
 
    # Запускаем цикл обработки событий окна
    window.mainloop()
 
# Проверяем, есть ли файл авторизации
entry_file = "credentials.pkl"
if os.path.exists(entry_file):
    # Открываем файл в режиме чтения в бинарном формате
    with open("credentials.pkl", "rb") as f:
        # Десериализуем словарь из файла
        credentials = pickle.load(f)
        # Получаем зашифрованные логин и пароль из словаря
        encrypted_login = credentials["login"]
        encrypted_password = credentials["password"]
        key = credentials["key"]
        cipher = Fernet(key)
        # Расшифровываем логин и пароль с помощью ключа
        login = cipher.decrypt(encrypted_login).decode()
        password = cipher.decrypt(encrypted_password).decode()
# Если логин не совпал или нет файла авторизации, то открываем окно авторизации
    if password == "":
        print(f"Для пользователя {login} не был указан пароль, попробуйте еще раз")
        show_login_form()       
if oslogin != login or not os.path.exists(entry_file):
# Вызываем функцию, которая показывает форму входа
    show_login_form()
    # Открываем файл в режиме записи в бинарном формате
    with open(entry_file, "wb") as f:
        credentials = {}
        # Генерируем ключ шифрования
        key = Fernet.generate_key()
        # Создаем объект для шифрования и расшифрования
        cipher = Fernet(key)
        # Шифруем логин и пароль
        encrypted_login = cipher.encrypt(login.encode())
        encrypted_password = cipher.encrypt(password.encode())
        # Сохраняем зашифрованные логин и пароль в словарь
        credentials["login"] = encrypted_login
        credentials["password"] = encrypted_password
        credentials["key"] = key
        pickle.dump(credentials, f)

# Найти дату в ПДФ
def date_name(filename):
    # Открываем pdf-файл в бинарном режиме
    file_obj = io.BytesIO(filename)
        # Создаем объект PdfFileReader
    reader = PyPDF3.PdfFileReader(file_obj)
    # Определяем регулярное выражение для даты в формате dd.mm.yy
    date_pattern = r"\d{2}\.\d{2}\.\d{2}"
    # Создаем переменную для хранения найденной даты
    date_found = None
    # Перебираем все страницы в файле
    for i in range(reader.numPages):
        # Получаем страницу по индексу
        page = reader.getPage(i)
        # Получаем текст с страницы
        text = page.extractText()
        # Ищем дату в тексте с помощью регулярного выражения
        match = re.search(date_pattern, text)
        # Проверяем, есть ли совпадение
        if match:
            # Получаем найденную дату
            date_found = match.group()
            # Прерываем цикл
            break
    return(date_found)

# Найти количество ТП в ПДФ
def count_tp(filename):
    # Открываем pdf-файл в бинарном режиме
    file_obj = io.BytesIO(filename)
        # Создаем объект PdfFileReader
    reader = PyPDF3.PdfFileReader(file_obj)
    unique_code = set()
    # Определяем регулярное выражение для кода ТП
    code_pattern = re.compile(r'\d{10}')
    # Перебираем все страницы в файле
    for i in range(reader.numPages):
        # Получаем страницу по индексу
        page = reader.getPage(i)
        # Получаем текст с страницы
        text = page.extractText()
        # Ищем ТП в тексте с помощью регулярного выражения
        matches = code_pattern.findall(text)
        # Заполняем набор
        unique_code.update(matches)
    code_count = len(unique_code)
    return(code_count)

# Записываем в лог сообщение о начале загрузки
logging.info(f"Считываем список объектов и зон из файла urls.xlsx")
print(f"Считываем список объектов и зон из файла urls.xlsx")
# Читаем файл excel с url
df = pd.read_excel("urls.xlsx")

if tool_mode in ['mm','alt']:
    start_folder = filedialog.askdirectory(title='Выберите стартовую папку')
    start_folder=start_folder.replace("/","\\")
    if start_folder == "":
        start_folder = exe_path
    print(f"Стартовой папкой выбрана {start_folder}")

CACHE_FILE = "branch_office_cache.json"
BASE_URL = "https://merchant.corp.tander.ru/api/branch-office/list"
# Словарь соответствий для типов PDF
PDF_TYPE_SUFFIX = {
    "Стандартная ПГ с фото ТП": "_images.pdf",
    "Блочная ПГ": "_block.pdf",
    "Блочная ПГ обезличенная": "_nblock.pdf",
    "ПГ для схемограмм": "_schemogram.pdf",
    "ПГ для ценникодержателей": "_barcode.pdf",
    "Без фото": ".pdf"
}

def load_cached_data():
    """
    Загружает данные из кэша или запрашивает их, если кэш устарел или отсутствует.
    """
    # Проверяем существование файла
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cached_data = json.load(f)
            last_updated = cached_data.get("last_updated")
            # Проверяем актуальность данных (в данном случае, по времени кэша)
            if last_updated and (datetime.now() - datetime.fromisoformat(last_updated)).days < 7:
                print("Загружаем данные из кэша...")
                return cached_data["data"]
    
    # Если кэш отсутствует или устарел, делаем запрос
    print("Обновляем данные с сервера...")
    response = requests.get(BASE_URL, auth, verify=False)
    if response.status_code == 200:
        data = response.json()
        # Сохраняем данные в кэш
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_updated": datetime.now().isoformat(), "data": data}, f, ensure_ascii=False, indent=4)
        return data
    else:
        raise Exception(f"Ошибка загрузки данных: {response.status_code}")

def extract_zones(object_name, object_cache, auth):
    """
    Загружает данные о торговых залах и зонах для заданного объекта из кэша.
    Если данные отсутствуют в кэше, делает запрос к API.
    
    Параметры:
        object_name (str): Имя объекта.
        object_cache (dict): Кэш объектов.
        auth (tuple): Авторизационные данные (логин, пароль).
        
    Возвращает:
        list: Список торговых залов с их зонами.
    """
    matched_object = next((obj for obj in object_cache if obj["name"] == object_name), None)
    if not matched_object:
        print(f"Объект {object_name} не найден в списке.")
        return []

    id_ws = matched_object["id_ws"]
    layout_url = f"http://merchant.corp.tander.ru/api/branch-office/{id_ws}/layout-zones-tree/"
    
    try:
        response = requests.get(layout_url, auth)
        response.raise_for_status()  # Поднимаем исключение, если статус ответа не 200
        zones_data = response.json()
        
        # Форматируем данные, добавляя флаг "approved" для удобства
        formatted_zones = []
        for tz in zones_data:
            is_approved = "(Утверждённый)" in tz["text"]
            tz_name = tz["text"].replace(" (Утверждённый)", "").strip()
            formatted_zones.append({
                "id": tz["id"],
                "name": tz_name,
                "approved": is_approved,
                "zones": tz.get("zones", [])  # Зоны внутри торгового зала
            })
        return formatted_zones
    except requests.exceptions.RequestException as e:
        print(f"Ошибка загрузки зон для {object_name}: {e}")
        return []
    
def generate_links(dataframe, object_cache, auth):
    """
    Генерирует ссылки для каждой строки пользовательского списка.
    Возвращает список URL.
    """
    links = []
    for _, row in dataframe.iterrows():
        store_name = row["Имя магазина"]
        zone_name = row["Имя зоны"]
        tz_name = row["Имя ТЗ"] if not pd.isna(row["Имя ТЗ"]) else None
        pdf_type = row["Тип PDF"]
        path = row["Полный путь"]

        # Загружаем зоны для магазина
        zones_data = extract_zones(store_name, object_cache, auth)
        if not zones_data:
            links.append(None)  # Пропуск строки, если зоны не найдены
            continue

        # Если ТЗ не указан, выбираем первый утверждённый
        if not tz_name:
            matching_tz = next((tz for tz in zones_data if tz["approved"]), None)
            if not matching_tz:
                links.append(None)
                continue
        else:
            matching_tz = next((tz for tz in zones_data if tz["name"] == tz_name), None)
            if not matching_tz:
                links.append(None)
                continue

        # Ищем зону внутри торгового зала
        matching_zone = next((zone for zone in matching_tz["zones"] if zone["name"] == zone_name), None)
        if not matching_zone:
            links.append(None)
            continue

        # Формируем ссылку
        id_zone = matching_zone["id"]
        pdf_suffix = PDF_TYPE_SUFFIX.get(pdf_type, ".pdf")
        link = f"http://merchant.corp.tander.ru/api/layout-zone/{path}/{id_zone}{pdf_suffix}"
        links.append(link)

    return links

# Создаем список url
object_cache = load_cached_data()
urls = generate_links(df, object_cache)

# Добавляем ссылки как новый столбец
df["Ссылка"] = urls

# Упрощаем доступ к данным: готовый DataFrame уже содержит всё необходимое
pairs = df[["Имя объекта", "Имя зоны", "Имя ТЗ", "Тип PDF", "Полный путь", "Формат", "Ссылка"]].to_records(index=False)


# Определяем функцию для загрузки файла по url
def download_file(pair):
    name, zone, path, tz_name, pdf_type, path, ws_format, url = pair
    global login, password, auth
    # Задаем параметры авторизации
    auth = (login, password)
    # Отправляем запрос к url с авторизацией
    response = requests.get(url, auth=auth)
    # Записываем в лог сообщение о начале загрузки
    logging.info(f"Получаем ответ от сервера: {response.status_code}")
    if response.status_code == 401:
        logging.error(f"Авторизация пользователя {login} не удалась")
        print(f"Авторизация пользователя {login} не удалась")
    if response.status_code == 500:
        logging.error(f"Ошибка генерации планограммы {url}, скорее всего не настроена блочная подсветка")    
        print(f"Ошибка генерации планограммы {url}, скорее всего не настроена блочная подсветка")   
    print(f"Получаем ответ от сервера: {response.status_code}")
    # Проверяем статус ответа
    if response.status_code == 200:
        # Получаем имя файла из url
        filename = response.content
        # Записываем в лог сообщение о начале загрузки
        logging.info(f"Начинаем загрузку файла по url: {url}")
        print(f"Начинаем загрузку файла по url: {url}")
        date_found = date_name(filename)
        # Проверяем, надо ли показывать дату
        if date_found == None or ShowDate == False:
            date_found = ""
        else:
            date_found= " " + date_found[:6] + '20' + date_found[-2:]
        code_count = count_tp(filename)
        # Проверяем, надо ли показывать количество позиций
        if ShowPos == True:
            showedPos = " (" + str(code_count) + " ТП)"
        else:
            showedPos = ""
        
        # Задаем шаблон регулярного выражения, который соответствует всем символам, которые нельзя использовать в качестве имени файла
        # Это символы: \ / : * ? " < > |
        pattern = r"[\\\/:*?\"<>|]"
        replacement = "-"
        zone = re.sub(pattern, replacement, zone)
        if tool_mode == 'gm':
            filename = (f"{path}\{zone}{date_found}.pdf")
        elif tool_mode == 'mm':
            if SeparateFolders == False:
                path = start_folder
            else:
                path = start_folder + "\\" + zone
            path = start_folder + "\\" + zone
            filename = (f"{path}\{name}{date_found}{showedPos}.pdf")
            print(filename)
        elif tool_mode == 'alt':
            if SeparateFolders == False:
                path = start_folder
            else:
                path = start_folder + "\\" + name
            filename = (f"{path}\{zone}{date_found}{showedPos}.pdf")
            print(filename)
        if not os.path.exists(path):
            os.makedirs(path)
        # Открываем файл для записи в бинарном режиме
        with open(filename, "wb") as f:
            # Записываем содержимое ответа в файл
            f.write(response.content)
            logging.info(f"Сохраняем файл {filename}, который содержит {code_count} позиций")
            print(f"Сохраняем файл {filename}, который содержит {code_count} позиций")
        # Возвращаем имя файла и размер в байтах
        return filename, len(response.content)
    else:
        # Возвращаем None в случае ошибки
        logging.error(f"Ошибка {response.status_code}")
        print(f"Ошибка {response.status_code}")
        return None
 
# Создаем пул потоков для параллельной загрузки файлов
with concurrent.futures.ThreadPoolExecutor() as executor:
    # Запускаем функцию для каждого url в списке
    results = executor.map(download_file, pairs)
    # Выводим результаты
    i = 0
    err = 0
    for result in results:
        if result is not None:
            # Распаковываем имя файла и размер
            filename, size = result
            # Выводим сообщение об успешной загрузке
            logging.info(f"Файл {filename} успешно загружен, размер {round(size/1024/1024, 2)} Мб")
            print(f"Файл {filename} успешно загружен, размер {round(size/1024/1024, 2)} Мб")
            i+=1
        else:
            # Выводим сообщение об ошибке
            logging.error(f"Ошибка при загрузке файла")
            print(f"Ошибка при загрузке файла")
            err +=1

end_time = time.time()
elapsed_time = end_time-start_time
delta = datetime.timedelta(seconds=round(elapsed_time,0))
if err > 0:
    err_text = ". Не было выгружено "+str(err) + " файлов."
else:
    err_text=""
logging.info(f"Было выгружено {i} файлов за {str(delta)}{err_text}")
print(f"Было выгружено {i} файлов за {str(delta)}{err_text}")
input("Нажмите Enter для завершения...")