from flask import Flask, request, render_template, session, redirect, g
from log import LoggerWriter
from pathlib import Path
from logging.handlers import TimedRotatingFileHandler
from selenium.webdriver.chrome.options import Options
import webview
import logging
import sys
import selenium_code
import os, sys
import threading


base_dir = '.'
if hasattr(sys, '_MEIPASS'):
    base_dir = os.path.join(sys._MEIPASS)
app = Flask(__name__, template_folder=os.path.join(base_dir, 'templates'))  
app.secret_key = "m6k5z7e8@"
window = webview.create_window('Kundelik', app, width=800, height=600, resizable=False)

# Конфигурация драйвера 
# executable_path = './chromedriver.exe'
# chrome_options = Options()
# # chrome_options.add_argument('--headless')
# # chrome_options.add_argument('--no-sandbox')
# # chrome_options.add_argument('--disable-dev-shm-usage')
# service_log_path = 'chromedriver.log'
# service_log = open(service_log_path, 'w')
# service = webdriver.chrome.service.Service('chromedriver', log_path=service_log_path)
# driver = webdriver.Chrome(executable_path=executable_path, chrome_options=chrome_options, service=service)

# Выбор языка
@app.route('/language', methods=['POST'])
def set_language():
    language = request.form['language']
    session['language'] = language
    if language == 'kz':
        return render_template('form_kz.html')
    else:
        return render_template('form_ru.html') 

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password1']
        if username == 'admin.sko' and password == 'Qwerty123@@@':
            return render_template('form_ru.html')
        else:
            return render_template('login_error_ru.html')
    else:
        return render_template('login.html')

# Страница интерфейса
@app.route('/main', methods =["GET", "POST"])
def start():
    language = session.get('language', 'ru')
    if language == 'kz':
        return render_template('form_kz.html')
    else:
        return render_template('form_ru.html')

# Запуск процесса выгрузки
@app.route('/selenium-start', methods =["GET", "POST"])
def selenium_start():
    name = ''
    groupNumber = ''
    group = False
    role = 0
    name_not_found = ''
    class_not_found = ''
    if request.method == "POST":
        name = request.form.get("fio")
        groupNumber = request.form.get("class")
        group_True = request.form.get("check")
        if group_True:
            group = True
        role = request.form['role']
        status = selenium_code.main(int(role), name, groupNumber, group)
        if status == 'Ученик не найден':
            name_not_found = name
        elif status == 'Класс не найден':
            class_not_found = groupNumber
        language = session.get('language', 'ru')
        if language == 'kz':
            return render_template('success_kz.html', name = name_not_found, group = class_not_found)
        else:
            return render_template('success_ru.html', name = name_not_found, group = class_not_found)
        
        
# Create Logger if doesn't exist
# Path("log").mkdir(parents=True, exist_ok=True)
# formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
# handler = TimedRotatingFileHandler('log/error.log', when="midnight", 
# interval=1, encoding='utf8')
# handler.suffix = "%Y-%m-%d"
# handler.setFormatter(formatter)
# logger = logging.getLogger()
# logger.setLevel(logging.ERROR)
# logger.addHandler(handler)
# sys.stdout = LoggerWriter(logging.debug)
# sys.stderr = LoggerWriter(logging.warning)


if __name__=='__main__':
    # threading.Timer(1.25, webbrowser.open('http://127.0.0.1:8989/')).start()
    # app.run(port=8989, debug=True)
    webview.start()