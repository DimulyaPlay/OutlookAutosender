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
from PyQt5.QtWidgets import QDialog, QLineEdit, QCheckBox, QWidget, QComboBox, QFormLayout, QVBoxLayout, QDialogButtonBox, QMessageBox, QTableWidgetItem
from PyQt5 import uic
from PyQt5.QtGui import QIcon
import winreg
import win32print
import random
import pythoncom
from PyPDF2 import PdfReader, PdfWriter
from threading import Thread, Lock
from urllib.request import getproxies


def save_config(config_file, config):
    try:
        with open(config_file, 'w') as json_file:
            json.dump(config, json_file, indent=4)
    except:
        print_exc()

config_path = os.path.dirname(sys.argv[0])
temp_path = os.path.join(config_path, 'temp')
if os.path.isdir(temp_path):
    try:
        shutil.rmtree(temp_path)
        os.mkdir(temp_path)
    except:
        try:
            import psutil
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] == 'wget.exe':
                    proc.terminate()
                    break
            shutil.rmtree(temp_path)
            os.mkdir(temp_path)
        except:
            pass
else:
    os.mkdir(temp_path)
if not os.path.exists(config_path):
    os.mkdir(config_path)
ReportsPrinted = os.path.join(config_path, 'ReportsPrinted')
if not os.path.exists(ReportsPrinted):
    os.mkdir(ReportsPrinted)
config_file = os.path.join(config_path, 'config.json')
PROCESSED_ITEMS_FILE = os.path.join(config_path, 'saved_msg.txt')
DOWNLOADED_ITEMS_FILE = os.path.join(config_path, 'downloaded_msg.txt')
in_queue_items = set()
items_lock = Lock()


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
        'checkBox_start_dm': False,
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
                zip_filelist = [f'{num+1}. {zipf.filename}' for num, zipf in enumerate(zipObj.filelist) if zipf.filename != 'meta.json']
            with open(foldername_extracted+'\\meta.json', 'r') as metafile:
                meta = json.load(metafile)
            subject = f'-{meta["id"]}] {meta["subject"]}' if meta['subject'] else f'-{meta["id"]}] {os.path.basename(archive)}'
            body_text = meta['body'] + '\n' + "Список направляемых файлов:" + '\n\n' + "\n".join(zip_filelist)
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
                    message.Subject = u"{}".format(subject_eml)
                    message.Body = u"{}".format(body_text)
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
                pdf_report = os.path.join(config['lineedit_output_edo'], f'report-{meta["thread"]}-{meta["id"]}-eml.zip')
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
                items = sent_folder.Items
                items.Sort("[SentOn]", True)  # Сортировка по времени отправки
                for i in range(1, min(5, len(items) + 1)):  # Ограничение до 6 элементов
                    item = items.Item(i)
                    if item.Subject == tracked_msg_subject:
                        sent_message = item
                        break
            except Exception as e:
                sent_files.append(f'Ошибка при поиске сообщения: {e}')
                print_exc()
            time.sleep(1)
        if sent_message:
            try:
                save_msg_to_zip(sent_message, msg_report)
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


def check_inbox_for_responses(root, namespace, config, download_queue):
    try:
        processed_items = load_processed_items()
        downloaded_items = load_downloaded_items()
        inbox_folder = namespace.GetDefaultFolder(6)  # Папка "Входящие"
        items = inbox_folder.Items
        items.Sort("[ReceivedTime]", True)  # Сортировка по времени получения
        for i in range(1, min(101, len(items) + 1)):  # Ограничение до 20 элементов
            item = items.Item(i)
            if item.EntryID not in processed_items and item.EntryID not in in_queue_items and item.EntryID not in downloaded_items:
                if config['checkBox_autosend_edo']:
                    pattern = r'.*\[tid-(\d+)-(\d+)\].*'  # Паттерн для поиска
                    match = re.search(pattern, item.Subject)
                    if match:
                        print(f'Найдено сообщение с паттерном: {item.Subject}')
                        thread_id = match.group(1)
                        unique_id = match.group(2)
                        pattern_extracted = f'{thread_id}-{unique_id}'
                        zip_path = os.path.join(config['lineedit_output_edo'], f'response-{pattern_extracted}.zip')
                        save_msg_to_zip(item, zip_path)
                        print(f'Сообщение и вложения сохранены в архив: {zip_path}')
                        processed_items.add(item.EntryID)
                        save_processed_items(processed_items)  # Обновляем файл после добавления нового сообщения
                if config['checkBox_start_dm']:
                    sender_email = resolve_sender_email_address(item)
                    if sender_email in config['mail_rules'].keys():
                        root.add_log_message(f'Найдено письмо от {sender_email}. проверка правил.')
                        message_subject = item.Subject
                        message_body = item.HTMLBody
                        rules = config['mail_rules'][sender_email]
                        for rule in rules:
                            root.add_log_message('______________________')
                            root.add_log_message(f'Правило {rule["rule_name"]}.')
                            if rule['subject_contains'] and rule['subject_contains'].lower() not in message_subject.lower():
                                root.add_log_message(f'Тема не содержит {rule["subject_contains"]}. Письмо ОТКЛОНЕНО.')
                                continue
                            filename, download_link = extract_re_rules(rule, message_subject + ' ' + message_body)
                            if not download_link:
                                root.add_log_message(f'Не найдена ссылка для скачивания. Письмо ОТКЛОНЕНО.')
                                continue
                            if not filename:
                                root.add_log_message(f'Не найдено имя. Файл будет сохранен с временным штампом.')
                            else:
                                root.add_log_message(f'Файл найден, будет сохранен с именем {filename}')
                            file_name = safe_filename(filename)
                            save_folder = rule['save_folder']
                            save_filename = os.path.join(save_folder, f'{file_name}.zip')
                            if os.path.exists(save_filename):
                                root.add_log_message(f'По пути {save_filename} уже имеется файл. Загрузка ОТМЕНЕНА.')
                                processed_items.add(item.EntryID)
                                save_processed_items(processed_items)  # Обновляем файл после добавления нового сообщения
                                continue
                            download_queue.put([item.EntryID, save_filename, download_link])
                            in_queue_items.add(item.EntryID)
                    else:
                        processed_items.add(item.EntryID)
                        save_processed_items(processed_items)  # Обновляем файл после добавления нового сообщения
    except Exception as e:
        print_exc()
        print(f'Ошибка при мониторинге папки входящих: {e}')
        return 1


def resolve_sender_email_address(item):
    try:
        try:
            return item.Sender.GetExchangeUser().PrimarySmtpAddress
        except:
            return item.SenderEmailAddress
    except:
        return 'OutlookMailServer'


def resolve_recipients_email_address(item):
    addresses = []
    for x in item.Recipients:
        try:
            addresses.append(x.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
        except AttributeError:
            addresses.append(x.AddressEntry.Address)
    return addresses


def save_msg_to_zip(item, zip_name):
    temp_dir = zip_name[:-4]
    os.makedirs(temp_dir, exist_ok=True)
    meta = {
        'date': item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
        'sender': resolve_sender_email_address(item),
        'recipients': resolve_recipients_email_address(item),
        'subject': item.Subject,
        'attachments': {},
        'body': item.Body
    }
    for att in item.Attachments:
        file_uuid = str(uuid4()) + '.' + att.FileName.split('.')[-1]
        file_path = os.path.join(temp_dir, file_uuid)
        att.SaveAsFile(file_path)
        meta['attachments'][file_uuid] = att.FileName
    meta_path = os.path.join(temp_dir, 'meta.json')
    with open(meta_path, 'w', encoding='utf-8') as meta_file:
        json.dump(meta, meta_file, ensure_ascii=False, indent=4)
    zip_path = zip_name
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(meta_path, os.path.basename(meta_path))
        for file_uuid, original_name in meta['attachments'].items():
            file_path = os.path.join(temp_dir, file_uuid)
            zipf.write(file_path, file_uuid)
    shutil.rmtree(temp_dir)
    print(f'Сообщение и вложения сохранены в архив: {zip_path}')


def monitor_inbox_periodically(root, namespace, config, download_queue):
    while True:
        error = check_inbox_for_responses(root, namespace, config, download_queue)
        if error:
            root.add_log_message('Возникла ошибка при проверке входящих сообщений. Повтор через 10 мин.')
            time.sleep(600)
        time.sleep(60)


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


def get_timestamp_date():
    now = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    return now


def load_processed_items():
    if not os.path.exists(PROCESSED_ITEMS_FILE):
        return set()
    with open(PROCESSED_ITEMS_FILE, 'r') as f:
        return set(line.strip() for line in f)


def save_processed_items(processed_items):
    with open(PROCESSED_ITEMS_FILE, 'w') as f:
        for item in processed_items:
            f.write(f"{item}\n")


def load_downloaded_items():
    if not os.path.exists(DOWNLOADED_ITEMS_FILE):
        return set()
    with open(DOWNLOADED_ITEMS_FILE, 'r') as f:
        return set(line.strip() for line in f)


def save_downloaded_items(downloaded_items):
    with open(DOWNLOADED_ITEMS_FILE, 'w') as f:
        for item in downloaded_items:
            f.write(f"{item}\n")

def extract_re_rules(rules, email_text):
    pattern_for_find = rules['re_filename']
    pattern_link = rules['re_link']
    filename_match = re.search(pattern_for_find, email_text)
    filename = filename_match.group(1) if filename_match else None
    link_match = re.search(pattern_link, email_text)
    download_link = link_match.group(1) if link_match else None
    return filename, download_link


def download_wget(kuvi_link, kuvi_path):
    proxy = getproxies().get('http')
    temp_path = os.path.join(os.curdir, 'temp', os.path.basename(kuvi_path))
    try:
        command = ["wget/wget.exe"]
        if proxy:
            command.extend(["-e", "use_proxy=yes",
                            "-e", f"http_proxy={proxy}",
                            "-e", f"https_proxy={proxy}"])
        if config['limit_rate']:
            command.append(f"--limit-rate={config['limit_rate']}k")
        command.append('--no-check-certificate')
        command.append('--tries=3')
        command.extend(["-O", temp_path, kuvi_link])
        subprocess.run(command, check=True)
    except:
        print_exc()
        return 1
    if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
        shutil.move(temp_path, kuvi_path)
        return 0


def safe_filename(filename):
    # Заменяем недопустимые символы на "-"
    if filename:
        filename = re.sub(r'[<>:"/\\|?*]', '-', filename)
        return filename
    else:
        return get_timestamp_date()


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


class DMThread(Thread):
    def __init__(self, config, task_queue):
        Thread.__init__(self, daemon=True)
        self.config = config
        self.task_queue = task_queue

    def run(self):
        while True:
            try:
                entryid, save_filename, download_link = self.task_queue.get()
                error = download_wget(download_link, save_filename)
                if error:
                    self.task_queue.put([entryid, save_filename, download_link])
                    print(f'Не удалось загрузить {save_filename}. Отправлен в конец очереди.')
                else:
                    with items_lock:
                        downloaded_items = load_downloaded_items()
                        downloaded_items.add(entryid)
                        save_downloaded_items(downloaded_items)
                        processed_items = load_processed_items()
                        processed_items.add(entryid)
                        save_processed_items(processed_items)  # Обновляем файл после добавления нового сообщения
                        in_queue_items.remove(entryid)
                    self.task_queue.task_done()
            except Exception as e:
                print_exc()
                print(f'Ошибка в процессе обработки: {e}')
