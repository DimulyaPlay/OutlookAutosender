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
from PyQt5.QtWidgets import QDialog, QLineEdit, QCheckBox, QWidget, QComboBox, QFormLayout, QVBoxLayout, QDialogButtonBox, QMessageBox, QTableWidgetItem
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

urllib3.disable_warnings()


def save_config(config_file, config):
    try:
        with open(config_file, 'w') as json_file:
            json.dump(config, json_file, indent=4)
    except:
        print_exc()

config_path = os.path.dirname(sys.argv[0])
if not os.path.exists(config_path):
    os.mkdir(config_path)
config_file = os.path.join(config_path, 'config.json')
message_queue = Queue()

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
        'lineedit_rr_address': '',
        'mail_rules': {'noreply-site@rosreestr.ru': [{'rule_name': 'Росреестр 1',
                                                      'subject_contains': 'о завершении обработки',
                                                      're_filename': r'<b>(.*?)</b>',
                                                      're_link': r'<a href="(.*?)">по ссылке</a>',
                                                      'save_folder': 'C://'}]},
        'limit_rate': 0
    }
    if not os.path.exists(config_file):
        save_config(config_file, default_configuration)
    return default_configuration


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
    recipients = config['lineEdit_recipients'].split(";")
    if not recipients:
        print('Получатели не обнаружены')
        return
    outlook = win32com.client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    message = outlook.CreateItem(0)
    message.Subject = config['lineEdit_subject']
    attachments_list = []
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
        temp_filepath = os.path.join(temp_path, os.path.basename(att))
        shutil.copy(att, temp_filepath)
        message.Attachments.Add(temp_filepath)
        os.unlink(temp_filepath)
        shutil.move(orig_att, config['lineEdit_put_path'])
        if config['checkBox_use_encryption']:
            os.unlink(att)
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
