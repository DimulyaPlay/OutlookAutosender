import json
import os
import shutil
import traceback
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
from zipfile import ZipFile
import random
import pythoncom
pythoncom.CoInitialize()


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


def read_create_config(config_file):
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
        'radioButton_periodic': True,
        'radioButton_monitoring': False,
        'radioButton_schedule': False,
        'timeEdit_send_period': '01:00',
        'lineEdit_schedule': '13:10,17:20',
        'checkBox_autorun': False,
        'checkBox_autostart': False,
        'timeEdit_connecting_delay': '0:00:30',
        'checkbox_use_edo': False,
        'lineedit_input_edo': '',
        'lineedit_output_edo': '',
        'lineedit_rr_address': '',
        'checkBox_autosend_edo': False
    }
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as configfile:
                config = json.load(configfile)
                for key in default_configuration.keys():
                    if key not in config.keys():
                        config[key] = default_configuration[key]
        except Exception as e:
            print(e)
            os.remove(config_file)
            config = default_configuration
            with open(config_file, 'w') as configfile:
                json.dump(config, configfile)
    else:
        config = default_configuration
        with open(config_file, 'w') as configfile:
            json.dump(config, configfile)
    return config


config = read_create_config(config_file)


def save_config(config):
    try:
        with open(config_file, 'w') as json_file:
            json.dump(config, json_file)
        config = read_create_config(config_file)
    except:
        traceback.print_exc()


def get_cert_names(cert_mgr_path):
    if os.path.exists(cert_mgr_path):
        try:
            result = subprocess.run([cert_mgr_path, '-list'], capture_output=True, text=True, check=True, encoding='cp866', creationflags=subprocess.CREATE_NO_WINDOW)
            output = result.stdout
        except subprocess.CalledProcessError as e:
            print(f"Ошибка выполнения команды: {e}")
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
    try:
        file_handle = open(filepath, 'a')
        msvcrt.locking(file_handle.fileno(), msvcrt.LK_NBLCK, 1)
        return True
    except:
        return False


def split_files_into_groups(file_paths):
    try:
        current_group_size = 0
        current_group = []
        all_groups = []
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
        traceback.print_exc()
    return all_groups


def archive_groups(groups):
    archived_groups_for_send = []
    archived_groups_names_files = []
    for i, group in enumerate(groups):
        config["job_id"] = config["job_id"] + i + 1
        save_config(config)
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
                        os.path.isfile(fp) and not fp.endswith(('desktop.ini', 'swapfile.sys', 'Thumbs.db')) and is_file_locked(fp)]
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


def agregate_edo_messages():
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        current_filelist = glob(config['lineedit_input_edo'] + '/*')
        current_filelist = [fp for fp in current_filelist if
                            os.path.isfile(fp) and fp.endswith('.zip') and is_file_locked(fp)]
        sent_files = []
        for archive in current_filelist:
            foldername_extracted = archive[:-4]
            if os.path.exists(foldername_extracted):
                foldername_extracted = foldername_extracted + str(random.randint(1, 9999))
            with ZipFile(archive, 'r') as zipObj:
                zipObj.extractall(foldername_extracted)
            with open(foldername_extracted+'\\meta.json', 'r') as metafile:
                meta = json.load(metafile)
            if meta['rr']:
                att_enc = encode_file(archive)
                message = outlook.CreateItem(0)
                message.Subject = meta['subject'] if meta['subject'] else os.path.basename(archive)
                message.Body = meta['body']
                message.Attachments.Add(att_enc)
                os.unlink(att_enc)
                recipient = message.Recipients.Add(config.get('lineedit_rr_address', 'no_addres'))
                recipient.Type = 1
                sender = namespace.CreateRecipient(namespace.CurrentUser.Address)
                sender.Resolve()
                message.SendUsingAccount = sender
                message.Send()
                sent_files.append(f'В РР отправлены файлы: {", ".join(meta["fileNames"])}')
            if meta['emails']:
                message = outlook.CreateItem(0)
                subject = f'[{meta["id"]}] '+meta['subject']
                message.Subject = subject
                message.Body = meta['body']
                attachments = [os.path.join(foldername_extracted, fileName) for fileName in meta['fileNames']]
                for att in attachments:
                    orig_att = att
                    temp_filepath = os.path.join(temp_path, os.path.basename(att))
                    shutil.copy(att, temp_filepath)
                    attachment = message.Attachments.Add(temp_filepath)
                    os.unlink(temp_filepath)
                    if config['checkBox_use_encryption']:
                        os.unlink(att)
                recipients = meta['emails'].split(';')
                for r in recipients:
                    if validate_email(r):
                        recipient = message.Recipients.Add(r)
                        recipient.Type = 1
                sender = namespace.CreateRecipient(namespace.CurrentUser.Address)
                sender.Resolve()
                message.SendUsingAccount = sender
                message.Send()
                sent_files.append(f'В адреса {meta["emails"]} отправлены файлы: {", ".join(meta["fileNames"])}')
                try:
                    sent_folder = namespace.GetDefaultFolder(5)
                    sorted_items = sorted(sent_folder.Items, key=lambda x: x.CreationTime, reverse=True)
                    sent_message = None
                    timeout = time.time() + 30  # Ждем не более 30 секунд
                    while not sent_message and time.time() < timeout:
                        sent_folder = namespace.GetDefaultFolder(5)
                        sorted_items = sorted(sent_folder.Items, key=lambda x: x.CreationTime, reverse=True)
                        for item in sorted_items[:5]:
                            if item.Subject == subject:
                                sent_message = item
                                break
                        time.sleep(1)
                    [os.remove(fp) for fp in glob(ReportsPrinted + "\\" + '*.pdf')]
                    printer_name = 'PDF24EDO'
                    default_printer = win32print.GetDefaultPrinter()
                    if default_printer != printer_name:
                        win32print.SetDefaultPrinter(printer_name)
                    sent_message.PrintOut()
                    win32print.SetDefaultPrinter(default_printer)
                    report_found = False
                    pdf_report = os.path.join(config['lineedit_output_edo'], f'\\{meta["id"]}.pdf')
                    while not report_found:
                        flist = glob(ReportsPrinted + "\\" + '*.pdf')
                        for f in flist:
                            if is_file_locked(f):
                                shutil.move(f, pdf_report)
                                report_found = True
                    sent_files.append(f'Отчет об отправке сохранен по пути {pdf_report}')
                except Exception as e:
                    sent_files.append(f'Но не удалось сохранить отчет об отправке: {e}')
                    traceback.print_exc()
            shutil.rmtree(foldername_extracted)
            shutil.move(archive, os.path.join(config['lineedit_input_edo'], 'sent'))
        if not sent_files:
            return 0
        return '\n'.join(sent_files)
    except:
        traceback.print_exc()
        return -1
    finally:
        pythoncom.CoUninitialize()


def send_mail(attachments, manual=True):
    recipients = config['lineEdit_recipients'].split(";")
    if not recipients:
        print('Получатели не обнаружены')
        return
    outlook = win32com.client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')

    message = outlook.CreateItem(0)
    message.Subject = config['lineEdit_subject']
    message.Body = config['plainTextEdit_body']

    for att in attachments:
        orig_att = att
        if config['checkBox_use_encryption']:
            orig_att = att
            att = encode_file(att)
        temp_filepath = os.path.join(temp_path, os.path.basename(att))
        shutil.copy(att, temp_filepath)
        attachment = message.Attachments.Add(temp_filepath)
        os.unlink(temp_filepath)
        shutil.move(orig_att, config['lineEdit_put_path'])
        if config['checkBox_use_encryption']:
            os.unlink(att)
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


class EdoWindow(QDialog):
    def __init__(self, config):
        super().__init__()
        ui_file = 'UI/edo.ui'
        uic.loadUi(ui_file, self)
        icon = QIcon("UI/icons8-carrier-pigeon-64.png")
        self.setWindowIcon(icon)
        self.config = config
        self.checkBox_use_edo = self.findChild(QCheckBox, 'checkbox_use_edo')
        self.checkBox_use_edo.setChecked(self.config.get('checkbox_use_edo', False))
        self.checkBox_use_edo.clicked.connect(lambda: self.save_params('checkbox_use_edo'))
        self.lineEdit_input_edo = self.findChild(QLineEdit, 'lineedit_input_edo')
        self.lineEdit_input_edo.setText(self.config.get('lineedit_input_edo', ''))
        self.lineEdit_input_edo.textChanged.connect(lambda: self.save_params('lineedit_input_edo'))
        self.lineEdit_output_edo = self.findChild(QLineEdit, 'lineedit_output_edo')
        self.lineEdit_output_edo.setText(self.config.get('lineedit_output_edo', ''))
        self.lineEdit_output_edo.textChanged.connect(lambda: self.save_params('lineedit_output_edo'))
        self.lineEdit_rr_address = self.findChild(QLineEdit, 'lineedit_rr_address')
        self.lineEdit_rr_address.setText(self.config.get('lineedit_rr_address', ''))
        self.lineEdit_rr_address.textChanged.connect(lambda: self.save_params('lineedit_rr_address'))

    def save_params(self, lineEdit_name):
        lineEdit = self.findChild(QLineEdit, lineEdit_name)
        if lineEdit:
            self.config[lineEdit_name] = lineEdit.text()
        checkBox = self.findChild(QCheckBox, 'checkbox_use_edo')
        self.config['checkbox_use_edo'] = checkBox.isChecked()
