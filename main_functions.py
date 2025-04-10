import json
import os
import shutil
import tempfile
import traceback
from traceback import print_exc
import msvcrt
from glob import glob
import zipfile
import win32com.client
import subprocess
import re
import time
import ctypes
import winshell
import sys
from datetime import datetime
from PyQt5.QtWidgets import (QDialog, QLineEdit, QCheckBox, QWidget, QComboBox, QFormLayout, QVBoxLayout,
                             QDialogButtonBox, QMessageBox, QTableWidgetItem, QLabel, QPushButton)
from PyQt5 import uic
from PyQt5.QtGui import QIcon
import winreg
import win32print
import random
import pythoncom
from uuid import uuid4
from PyPDF2 import PdfReader, PdfWriter
from threading import Thread, Lock
from urllib.request import getproxies
import requests
from queue import Queue
import urllib3
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import base64
from cryptography.fernet import Fernet
import hashlib
urllib3.disable_warnings()


def save_config(config_file, config):
    try:
        with open(config_file, 'w') as json_file:
            json.dump(config, json_file, indent=4)
    except:
        print_exc()

config_file = os.path.join(os.getcwd(), 'config.json')
message_queue = Queue()
temp_path = os.path.join(os.getcwd(), 'temp')
shutil.rmtree(temp_path, True)
os.mkdir(temp_path)


def load_or_create_default_config(config_file):
    default_configuration = {
        'job_id': 0,
        'lineEdit_get_path': r'',
        'lineEdit_put_path': r'',
        'checkBox_archive_files': True,
        'lineEdit_archname': 'Archive',
        'checkBox_with_sig_only': True,
        'checkBox_use_encryption': True,
        'lineEdit_csp_path': r"C:\Program Files\Crypto Pro\CSP",
        'comboBox_certs': 'Сертификат не выбран',
        'lineEdit_recipients': '',
        'lineEdit_subject': '',
        'plainTextEdit_body': '',
        'spinBox_max_attach_weight': 9,
        'checkBox_periodic': True,
        'checkBox_schedule': False,
        'timeEdit_send_period': '01:00',
        'lineEdit_schedule': '13:10,17:20',
        'checkBox_autorun': False,
        'checkBox_autostart': False,
        'timeEdit_connecting_delay': '0:00:30',
        'checkBox_use_outlook': True,
        'lineEdit_report_smtp_address': ''
    }
    if not os.path.exists(config_file):
        save_config(config_file, default_configuration)
    return default_configuration


def get_system_key():
    """Функция для получения уникального ключа на основе системных данных"""
    system_info = os.environ.get('COMPUTERNAME', '') + os.environ.get('USERPROFILE', '')
    key = hashlib.sha256(system_info.encode()).digest()
    return base64.urlsafe_b64encode(key)


def encrypt_data(data, key):
    """Шифруем данные"""
    fernet = Fernet(key)
    return fernet.encrypt(data)  # Убираем .encode(), так как data уже байты


def decrypt_data(data, key):
    """Расшифровываем данные"""
    fernet = Fernet(key)
    return fernet.decrypt(data)


def save_credentials(server, port, login, password, use_ssl):
    """Сохраняем зашифрованные учетные данные в файл"""
    encrypted_server = encrypt_data(server.encode(), get_system_key())
    encrypted_port = encrypt_data(port.encode(), get_system_key())
    encrypted_login = encrypt_data(login.encode(), get_system_key())
    encrypted_password = encrypt_data(password.encode(), get_system_key())
    encrypted_use_ssl = encrypt_data(use_ssl.encode(), get_system_key())
    with open('smtp_credentials.txt', 'wb') as f:
        f.write(encrypted_server + b'\n' +
                encrypted_port + b'\n' +
                encrypted_login + b'\n' +
                encrypted_password + b'\n' +
                encrypted_use_ssl)


def load_credentials():
    """Загружаем учетные данные из файла и расшифровываем их"""
    with open('smtp_credentials.txt', 'rb') as f:
        encrypted_server = f.readline().strip()
        encrypted_port = f.readline().strip()
        encrypted_login = f.readline().strip()
        encrypted_password = f.readline().strip()
        encrypted_use_ssl = f.readline().strip()
    server = decrypt_data(encrypted_server, get_system_key()).decode()
    port = decrypt_data(encrypted_port, get_system_key()).decode()
    login = decrypt_data(encrypted_login, get_system_key()).decode()
    password = decrypt_data(encrypted_password, get_system_key()).decode()
    use_ssl = decrypt_data(encrypted_use_ssl, get_system_key()).decode()
    return server, port, login, password, use_ssl


def read_create_config(config_file):
    default_configuration = load_or_create_default_config(config_file)
    try:
        with open(config_file, 'r') as configfile:
            config = json.load(configfile)
            for key, value in default_configuration.items():
                config.setdefault(key, value)
            save_config(config_file, config)
    except:
        save_config(config_file, default_configuration)
        config = default_configuration
        print_exc()
    return config


config = read_create_config(config_file)


def get_cert_names(cert_mgr_path):
    if os.path.exists(cert_mgr_path):
        try:
            result = subprocess.run([cert_mgr_path, '-list'], capture_output=True, text=True, check=True, encoding='cp866', creationflags=subprocess.CREATE_NO_WINDOW)
            output = result.stdout
        except subprocess.CalledProcessError as e:
            print(f"Ошибка выполнения команды: {e}")
            output = ''
        subject_lines1 = re.findall(r'Субъект\s+:\s+([^\n]+)', output)
        subject_lines2 = re.findall(r'Subject\s+:\s+([^\n]+)', output)
        cn_list1 = [re.search(r'CN=([^\n]+)', line).group(1) if 'CN=' in line else '' for line in subject_lines1]
        cn_list2 = [re.search(r'CN=([^\n]+)', line).group(1) if 'CN=' in line else '' for line in subject_lines2]
        cn_list = []
        cn_list.extend(cn_list1)
        cn_list.extend(cn_list2)
        return cn_list
    else:
        return []


def is_file_locked(filepath):
    file_handle = None
    try:
        file_handle = open(filepath, 'a')  # Попытка открыть файл для добавления данных
        msvcrt.locking(file_handle.fileno(), msvcrt.LK_NBLCK, 1)  # Попытка заблокировать файл
        # Если мы здесь, значит файл не заблокирован и его можно обрабатывать
        return False
    except:
        return True
    finally:
        if file_handle:
            file_handle.close()  # Не забываем закрыть файл


def split_files_into_groups(file_paths):
    current_group_size = 0
    current_group = []
    all_groups = []
    try:
        for file_path in file_paths:
            file_size_mb = os.path.getsize(file_path) / (1024 ** 2)  # Размер файла в мегабайтах
            if current_group_size + file_size_mb > config['spinBox_max_attach_weight']:
                all_groups.append(current_group)
                current_group = []
                current_group_size = 0
            current_group.append(file_path)
            current_group_size += file_size_mb
        if current_group:
            all_groups.append(current_group)
    except:
        print_exc()
    return all_groups


def archive_groups(groups):
    archived_groups_for_send = []
    archived_groups_names_files = []
    for i, group in enumerate(groups):
        config["job_id"] = config["job_id"] + i + 1
        save_config(config_file, config)
        read_create_config(config_file)
        archname = config['lineEdit_archname']
        zip_filename = os.path.join(config['lineEdit_get_path'], f'{archname}_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}_job{config["job_id"]}.zip')
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            filenames = []
            for file_path in group:
                arcname = os.path.basename(file_path)
                zip_file.write(file_path, arcname=arcname)
                filenames.append(os.path.basename(file_path))
                if config['checkBox_with_sig_only']:
                    sig_file_path1 = file_path + '.sig'
                    sig_file_path2 = file_path + '..sig'
                    sig_file_path3 = file_path + '.1.sig'
                    if os.path.isfile(sig_file_path1):
                        sig_file_path = sig_file_path1
                    elif os.path.isfile(sig_file_path2):
                        sig_file_path = sig_file_path2
                    elif os.path.isfile(sig_file_path3):
                        sig_file_path = sig_file_path3
                    arcname = os.path.basename(sig_file_path)
                    zip_file.write(sig_file_path, arcname=arcname)
                    filenames.append(os.path.basename(sig_file_path))
                    os.unlink(sig_file_path)
                os.unlink(file_path)
        archived_groups_for_send.append([zip_filename])
        archived_groups_names_files.append(filenames)
    return archived_groups_for_send, archived_groups_names_files


def encode_file(fp):
    subprocess.run([os.path.join(config['lineEdit_csp_path'], 'csptest.exe'), '-sfenc', '-encrypt', '-in', fp,
                    '-out', f'{fp}.enc', '-cert', f'{config["comboBox_certs"]}'], creationflags=subprocess.CREATE_NO_WINDOW)
    return f'{fp}.enc'


def gather_mail():
    current_filelist = glob(config['lineEdit_get_path'] + '/*')
    current_filelist = [fp for fp in current_filelist if
                        os.path.isfile(fp) and not fp.endswith(('desktop.ini', 'swapfile.sys', 'Thumbs.db')) and not is_file_locked(fp)]
    new_list = []
    if config['checkBox_with_sig_only']:
        for fp in current_filelist:
            if fp + '.sig' in current_filelist or fp + '..sig' in current_filelist or fp + '.1.sig' in current_filelist:
                new_list.append(fp)
    else:
        new_list = current_filelist
    groups_for_send = split_files_into_groups(new_list)
    groups_filenames = [[os.path.basename(fp) for fp in group] for group in groups_for_send]
    if config['checkBox_archive_files']:
        groups_for_send, groups_filenames = archive_groups(groups_for_send)
    return groups_for_send, groups_filenames


def send_mail(attachments, manual=True):
    try:
        recipients = config['lineEdit_recipients'].split(";")
        if not recipients:
            print('Получатели не обнаружены')
            return
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        message = outlook.CreateItem(0)
        message.Subject = config['lineEdit_subject']
        attachments_list = []
        encrypted_files = []
        for main_num, att in enumerate(attachments):
            if att.endswith('zip'):
                with zipfile.ZipFile(att, 'r') as zipObj:
                    attachments_list.extend([f'{main_num+1}.{num + 1}. {zipf.filename}' for num, zipf in enumerate(zipObj.filelist)])
            else:
                attachments_list.append(f"{main_num+1}. {os.path.basename(att)}")
            orig_att = att
            if config['checkBox_use_encryption']:
                orig_att = att
                att = encode_file(att)
                encrypted_files.append(att)
            temp_filepath = os.path.join(temp_path, os.path.basename(att))
            shutil.copy(att, temp_filepath)
            message.Attachments.Add(temp_filepath)
            os.unlink(temp_filepath)
            shutil.move(orig_att, config['lineEdit_put_path'])
        body_text = config['plainTextEdit_body'] if not attachments_list else config['plainTextEdit_body']+"\n\n\n"+"Список направляемых документов:\n"+"\n".join(attachments_list)
        message.Body = body_text
        for r in recipients:
            if validate_email(r):
                recipient = message.Recipients.Add(r)
                recipient.Type = 1
        sender = namespace.CreateRecipient(namespace.CurrentUser.Address)
        sender.Resolve()
        message.SendUsingAccount = sender
        if manual:
            message.Display()
        else:
            message.Send()
        del outlook
        if config['checkBox_use_encryption']:
            [os.unlink(att) for att in encrypted_files]
        return True, 'Удалось'
    except Exception as e:
        error_txt_path = os.path.splitext(orig_att)[0] + ".txt"
        with open(error_txt_path, "w", encoding="utf-8") as f:
            body_text = config['plainTextEdit_body'] if not attachments_list else config['plainTextEdit_body']+"\n\n\n"+"Список направляемых документов:\n"+"\n".join(attachments_list)
            f.write(body_text)
        return False, str(e)


def validate_email(email):
    # Регулярное выражение для проверки формата email-адреса
    pattern = re.compile(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')
    if re.match(pattern, email):
        return True
    else:
        return False


def check_time(time_str):
    try:
        hh, mm = time_str.split(':')
        if 0 <= int(hh) <= 23 and 0 <= int(mm) <= 59:
            return True
        else:
            return False
    except:
        return False


def add_to_startup():
    startup_folder = winshell.startup()
    shortcut_path = os.path.join(startup_folder, f"OtpravkaRosreestr.lnk")
    if os.path.isfile(shortcut_path):
        if config['checkBox_autorun']:
            return
        else:
            os.unlink(shortcut_path)
            ctypes.windll.user32.MessageBoxW(0, f"Программа успешно удалена из автозапуска!", "Успех", 1)
            return
    try:
        create_shortcut(shortcut_path)
        ctypes.windll.user32.MessageBoxW(0, f"Программа успешно добавлена в автозапуск!", "Успех", 1)
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f"Ошибка: {e}", "Ошибка", 1)


def create_shortcut(shortcut_path):
    from win32com.client import Dispatch
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = sys.argv[0]
    shortcut.WorkingDirectory = os.path.dirname(sys.argv[0])
    shortcut.save()


def send_mail_smtp(attachments):
    """Отправляет письмо с вложениями через SMTP.

    Возвращает:
        tuple: (success: bool, message: str)
    """
    try:
        try:
            smtp_server, smtp_port, sender_email, password, use_ssl = load_credentials()
        except:
            return False, f'Ошибка орашифровки данных подключения, настройте подключение SMTP заново.'
        recipients = config['lineEdit_recipients'].split(";")
        if not recipients:
            return False, 'Получатели не обнаружены'
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ', '.join(recipients)
        report_address = config['lineEdit_report_smtp_address'] if config['lineEdit_report_smtp_address'] else sender_email
        msg['Bcc'] = report_address
        recipients.append(report_address)
        msg['Subject'] = config['lineEdit_subject']
        attachments_list = []
        encrypted_files = []
        for main_num, att in enumerate(attachments):
            try:
                if att.endswith('zip'):
                    with zipfile.ZipFile(att, 'r') as zipObj:
                        attachments_list.extend(
                            [f'{main_num + 1}.{num + 1}. {zipf.filename}' for num, zipf in enumerate(zipObj.filelist)])
                else:
                    attachments_list.append(f"{main_num + 1}. {os.path.basename(att)}")
                orig_att = att
                if config['checkBox_use_encryption']:
                    att = encode_file(att)
                    encrypted_files.append(att)
                temp_filepath = os.path.join(temp_path, os.path.basename(att))
                shutil.copy(att, temp_filepath)
                with open(temp_filepath, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(att)}')
                msg.attach(part)
                os.unlink(temp_filepath)
                shutil.move(orig_att, config['lineEdit_put_path'])
            except Exception as e:
                if config['checkBox_use_encryption']:
                    [os.unlink(att) for att in encrypted_files]
                return False, f'Ошибка обработки вложения {att}: {str(e)}'
        body_text = config['plainTextEdit_body']
        if attachments_list:
            body_text += "\n\nСписок направляемых документов:\n" + "\n".join(attachments_list)
        msg.attach(MIMEText(body_text, 'plain'))
        timeout = 120
        if use_ssl == 'True':
            context = ssl.create_default_context()
            try:
                with smtplib.SMTP_SSL(smtp_server, int(smtp_port), context=context, timeout=timeout) as server:
                    server.login(sender_email, password)
                    server.sendmail(sender_email, recipients, msg.as_string())
            except ssl.SSLError as e:
                if config['checkBox_use_encryption']:
                    [os.unlink(att) for att in encrypted_files]
                return False, f'SSL ошибка: {str(e)}'
        else:
            try:
                with smtplib.SMTP(smtp_server, int(smtp_port), timeout=timeout) as server:
                    server.starttls()
                    server.login(sender_email, password)
                    server.sendmail(sender_email, recipients, msg.as_string())
            except smtplib.SMTPException as e:
                if config['checkBox_use_encryption']:
                    [os.unlink(att) for att in encrypted_files]
                return False, f'SMTP ошибка: {str(e)}'
        if config['checkBox_use_encryption']:
            [os.unlink(att) for att in encrypted_files]
        return True, "Отправлено успешно"
    except Exception as e:
        traceback.print_exc()
        return False, f'{str(e)}'


class SMTPConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.credentials_path = 'smtp_credentials.txt'
        self.setWindowTitle("Настройки подключения к SMTP")

        layout = QVBoxLayout()

        self.server_label = QLabel("Сервер:")
        self.server_input = QLineEdit()
        self.server_input.setPlaceholderText('smtp.yandex.ru')
        layout.addWidget(self.server_label)
        layout.addWidget(self.server_input)

        self.port_label = QLabel("Порт:")
        self.port_input = QLineEdit()
        self.port_input.setPlaceholderText('465')
        self.port_input.setText('465')
        layout.addWidget(self.port_label)
        layout.addWidget(self.port_input)

        self.login_label = QLabel("Логин:")
        self.login_input = QLineEdit()
        layout.addWidget(self.login_label)
        layout.addWidget(self.login_input)

        self.password_label = QLabel("Пароль:")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)

        self.use_ssl = QCheckBox('Использовать SSL')
        layout.addWidget(self.use_ssl)

        self.test_save_button = QPushButton("Проверить и сохранить")
        self.test_save_button.clicked.connect(self.test_and_save_connection)
        layout.addWidget(self.test_save_button)

        self.setLayout(layout)

    def test_and_save_connection(self):
        """Проверяет подключение к SMTP и сохраняет учетные данные"""
        credentials = {
            "server": self.server_input.text(),
            "port": self.port_input.text(),
            "login": self.login_input.text(),
            "password": self.password_input.text(),
            "use_ssl": str(self.use_ssl.isChecked())
        }
        if not all(credentials.values()):
            QMessageBox.critical(self, "Ошибка", "Все поля должны быть заполнены!")
            return
        try:
            sender_email = credentials["login"]
            recipients = [sender_email]
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = sender_email
            msg['Subject'] = "Тестовое письмо"
            msg.attach(MIMEText("Это тестовое письмо для проверки подключения к SMTP OutlookAutosender.", 'plain'))
            timeout = 5
            if credentials["use_ssl"] == "True":
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(
                        credentials["server"],
                        int(credentials["port"]),
                        context=context,
                        timeout=timeout
                ) as server:
                    server.login(sender_email, credentials["password"])
                    server.sendmail(sender_email, recipients, msg.as_string())
            else:
                with smtplib.SMTP(
                        credentials["server"],
                        int(credentials["port"]),
                        timeout=timeout
                ) as server:
                    server.starttls()
                    server.login(sender_email, credentials["password"])
                    server.sendmail(sender_email, recipients, msg.as_string())
            save_credentials(
                credentials["server"],
                credentials["port"],
                credentials["login"],
                credentials["password"],
                credentials["use_ssl"]
            )
            QMessageBox.information(self, "Успех", "Подключение успешно! Настройки сохранены.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(
                self,
                "Ошибка подключения",
                f"Не удалось подключиться к SMTP серверу:\n{str(e)}"
            )