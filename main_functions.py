import json
import os
import shutil
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
from PyQt5.QtWidgets import QDialog, QLineEdit, QCheckBox
from PyQt5 import uic
from PyQt5.QtGui import QIcon
import winreg
import win32print
import random
import pythoncom
from PyPDF2 import PdfReader, PdfWriter


def save_config(config_file, config):
    try:
        with open(config_file, 'w') as json_file:
            json.dump(config, json_file)
    except:
        print_exc()

config_path = os.path.dirname(sys.argv[0])
temp_path = os.path.join(config_path, 'temp')
if os.path.isdir(temp_path):
    shutil.rmtree(temp_path)
    os.mkdir(temp_path)
else:
    os.mkdir(temp_path)
if not os.path.exists(config_path):
    os.mkdir(config_path)
ReportsPrinted = os.path.join(config_path, 'ReportsPrinted')
if not os.path.exists(ReportsPrinted):
    os.mkdir(ReportsPrinted)
config_file = os.path.join(config_path, 'config.json')
PROCESSED_ITEMS_FILE = os.path.join(config_path, 'saved_msg.txt')


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
        'checkbox_use_edo': False,
        'lineedit_input_edo': '',
        'lineedit_output_edo': '',
        'lineedit_rr_address': '',
        'checkBox_autosend_edo': False,
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


def agregate_edo_messages(current_filelist):
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        sent_files = []
        for archive in current_filelist:
            foldername_extracted = archive[:-4]
            if os.path.exists(foldername_extracted):
                foldername_extracted = foldername_extracted + str(random.randint(1, 9999))
            with zipfile.ZipFile(archive, 'r') as zipObj:
                zipObj.extractall(foldername_extracted)
                zip_filelist = [f'{num+1}. {zipf.filename}' for num, zipf in enumerate(zipObj.filelist)]
            with open(foldername_extracted+'\\meta.json', 'r') as metafile:
                meta = json.load(metafile)
            subject = f'-{meta["id"]}] {meta["subject"]}' if meta['subject'] else f'-{meta["id"]}] {os.path.basename(archive)}'
            body_text = meta['body'] + '\n' + "Список направляемых файлов:" + '\n' + "\n".join(zip_filelist)
            [os.remove(fp) for fp in glob(ReportsPrinted + "\\" + '*.pdf')]
            if meta['rr']:
                att_enc = encode_file(archive)
                subject_rr = f"[rr-{meta['thread']}{subject}"
                recipient_rr = config.get('lineedit_rr_address')
                try:
                    message = outlook.CreateItem(0)
                    message.Subject = subject_rr
                    message.Body = body_text
                    message.Attachments.Add(att_enc)
                    os.unlink(att_enc)
                    recipient = message.Recipients.Add(recipient_rr)
                    recipient.Type = 1
                    sender = namespace.CreateRecipient(namespace.CurrentUser.Address)
                    sender.Resolve()
                    message.SendUsingAccount = sender
                    message.Send()
                    sent_files.append(f'В РР отправлены файлы: {", ".join(meta["fileNames"])}')
                except Exception as e:
                    sent_files.append(f'Ошибка отправки в РР: {e}')
            if meta['emails']:
                subject_eml = f"[tid-{meta['thread']}{subject}"
                attachments = [os.path.join(foldername_extracted, fileName) for fileName in meta['fileNames']]
                recipients = meta['emails'].split(';')
                try:
                    message = outlook.CreateItem(0)
                    message.Subject = subject_eml
                    message.Body = body_text
                    for att in attachments:
                        temp_filepath = os.path.join(temp_path, os.path.basename(att))
                        shutil.copy(att, temp_filepath)
                        message.Attachments.Add(temp_filepath)
                        os.unlink(temp_filepath)
                        if config['checkBox_use_encryption']:
                            os.unlink(att)
                    for r in recipients:
                        recipient = message.Recipients.Add(r)
                        recipient.Type = 1
                    message.Send()
                    sent_files.append(f'В адреса {meta["emails"]} отправлены файлы: {", ".join(meta["fileNames"])}')
                except Exception as e:
                    sent_files.append(f'Ошибка отправки: {e}')
                pdf_report = os.path.join(config['lineedit_output_edo'], f'report-{meta["thread"]}-{meta["id"]}-eml.msg')
                sent_files.extend(gather_report_from_sent_items(namespace, subject_eml, pdf_report))
            shutil.rmtree(foldername_extracted)
            shutil.move(archive, os.path.join(config['lineedit_input_edo'], 'sent', os.path.basename(archive)))
        if not sent_files:
            return 0
        return '\n'.join(sent_files)
    except:
        print_exc()
        return -1
    finally:
        pythoncom.CoUninitialize()


def gather_report_from_sent_items(namespace, tracked_msg_subject, msg_report):
    sent_files = []
    try:
        sent_message = None
        timeout = time.time() + 40  # Ждем не более 40 секунд
        while not sent_message and time.time() < timeout:
            try:
                sent_folder = namespace.GetDefaultFolder(5)  # Папка "Отправленные"
                items = list(sent_folder.Items)
                sorted_items = sorted(items, key=lambda x: x.CreationTime, reverse=True)
                for item in sorted_items[:5]:
                    if item.Subject == tracked_msg_subject:
                        sent_message = item
                        break
            except Exception as e:
                sent_files.append(f'Ошибка при поиске сообщения: {e}')
                print_exc()
            time.sleep(1)
        if sent_message:
            try:
                sent_message.SaveAs(msg_report)
                sent_files.append(f'Отчет об отправке сохранен по пути {msg_report}')
            except Exception as e:
                sent_files.append(f'Не удалось сохранить отчет об отправке: {e}')
                print_exc()
        else:
            sent_files.append(f'Истекло время ожидания появления письма в папке отправленных.')
    except Exception as e:
        sent_files.append(f'Общая ошибка: {e}')
        print_exc()
    return sent_files


def check_inbox_for_responses(namespace, config, stop_event):
    processed_items = load_processed_items()
    try:
        inbox_folder = namespace.GetDefaultFolder(6)  # Папка "Входящие"
        items = inbox_folder.Items
        items.Sort("[ReceivedTime]", True)  # Сортировка по времени получения
        pattern = r'.*\[tid-(\d+)-(\d+)\].*'  # Паттерн для поиска
        for i in range(1, min(21, len(items) + 1)):  # Ограничение до 20 элементов
            if stop_event.is_set():
                break
            item = items.Item(i)
            if item.EntryID not in processed_items:
                match = re.search(pattern, item.Subject)
                if match:
                    print(f'Найдено сообщение с паттерном: {item.Subject}')
                    thread_id = match.group(1)
                    unique_id = match.group(2)
                    pattern_extracted = f'{thread_id}-{unique_id}'
                    msg_response = os.path.join(config['lineedit_output_edo'], f'response-{pattern_extracted}-eml.msg')
                    try:
                        item.SaveAs(msg_response)
                        print(f'Сообщение сохранено по пути: {msg_response}')
                    except Exception as e:
                        print(f'Не удалось сохранить сообщение: {e}')
                    processed_items.add(item.EntryID)
                    save_processed_items(processed_items)  # Обновляем файл после добавления нового сообщения
    except Exception as e:
        print(f'Ошибка при мониторинге папки входящих: {e}')


def monitor_inbox_periodically(namespace, config, interval, stop_event):
    while not stop_event.is_set():
        check_inbox_for_responses(namespace, config, stop_event)
        stop_event.wait(interval)  # Используем wait с таймаутом вместо time.sleep

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
    message.Body = config['plainTextEdit_body']+"\n"+"Список направляемых документов:"+"\n"+"\n".join(attachments_list)
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


def get_timestamp():
    now = datetime.now()
    months = [
        'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
        'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
    ]
    timestamp = f"{now.day} {months[now.month - 1]} {now.year} г., {now.hour:02d}:{now.minute:02d}"
    return timestamp


def load_processed_items():
    if not os.path.exists(PROCESSED_ITEMS_FILE):
        return set()
    with open(PROCESSED_ITEMS_FILE, 'r') as f:
        return set(line.strip() for line in f)


def save_processed_items(processed_items):
    with open(PROCESSED_ITEMS_FILE, 'w') as f:
        for item in processed_items:
            f.write(f"{item}\n")


class EdoWindow(QDialog):
    def __init__(self, config):
        super().__init__()
        ui_file = 'UI/edo.ui'
        uic.loadUi(ui_file, self)
        icon = QIcon("UI/icons8-carrier-pigeon-64.png")
        self.config = {}
        self.setWindowIcon(icon)
        self.checkBox_use_edo = self.findChild(QCheckBox, 'checkbox_use_edo')
        self.checkBox_use_edo.setChecked(config.get('checkbox_use_edo', False))
        self.checkBox_use_edo.clicked.connect(lambda: self.save_params('checkbox_use_edo'))
        self.lineEdit_input_edo = self.findChild(QLineEdit, 'lineedit_input_edo')
        self.lineEdit_input_edo.setText(config.get('lineedit_input_edo', ''))
        self.lineEdit_input_edo.textChanged.connect(lambda: self.save_params('lineedit_input_edo'))
        self.lineEdit_output_edo = self.findChild(QLineEdit, 'lineedit_output_edo')
        self.lineEdit_output_edo.setText(config.get('lineedit_output_edo', ''))
        self.lineEdit_output_edo.textChanged.connect(lambda: self.save_params('lineedit_output_edo'))
        self.lineEdit_rr_address = self.findChild(QLineEdit, 'lineedit_rr_address')
        self.lineEdit_rr_address.setText(config.get('lineedit_rr_address', ''))
        self.lineEdit_rr_address.textChanged.connect(lambda: self.save_params('lineedit_rr_address'))

    def save_params(self, lineEdit_name):
        lineEdit = self.findChild(QLineEdit, lineEdit_name)
        if lineEdit:
            self.config[lineEdit_name] = lineEdit.text()
        checkBox = self.findChild(QCheckBox, 'checkbox_use_edo')
        self.config['checkbox_use_edo'] = checkBox.isChecked()
