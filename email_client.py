import email.header
import smtplib
import imaplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders, policy
from concurrent.futures import ThreadPoolExecutor
from PyQt5.QtWidgets import QDialog, QLineEdit, QRadioButton, QWidget, QComboBox, QPushButton, QMessageBox
from PyQt5 import uic
from PyQt5.QtGui import QIcon
import queue


def send_email(subject: str, body: str, attachments: list, recipients: list, cfg: dict):
    """
    :param subject: tema str
    :param body: telo str
    :param attachments: paths to attachmenrs list of str
    :param recipients: adresati list of str
    :param cfg: app config dict
    :return: resultcode 0 - success -1 - error
    """
    try:
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = cfg.get('email_login', 'Email')
        msg['To'] = ', '.join(recipients)
        msg.attach(MIMEText(body, 'plain'))
        for file_path in attachments:
            part = MIMEBase('application', "octet-stream")
            with open(file_path, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(file_path.split('/')[-1]))
            msg.attach(part)
        address, port = cfg['email_out'].split(':')
        use_ssl = cfg.get('email_out_enc', 'Нет') != 'Нет'
        port = int(port)  # Преобразование порта в число
        if use_ssl:
            mail = smtplib.SMTP_SSL(address, port)
        else:
            mail = smtplib.SMTP(address, port)
        mail.login(cfg['email_login'], cfg['email_password'])
        mail.sendmail(cfg['email_login'], recipients, msg.as_string())
        mail.quit()
        print(f'Письмо отправлено: {recipients}')
        return 0, ''
    except Exception as e:
        print_exc()
        return -1, f"Ошибка при отправке тестового письма: {e}"


def check_email(subject: str, cfg: dict):
    """
    :param subject: тема для поиска
    :param cfg: параметры
    :return: resultcode 0 - success -1 - error 1 - not found, comment
    """
    try:
        imap_server, imap_port = cfg["email_in"].split(':')
        use_ssl = cfg.get('email_in_enc', 'Нет') != 'Нет'
        imap_port = int(imap_port)
        if use_ssl:
            mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        else:
            mail = imaplib.IMAP4(imap_server, imap_port)
        mail.login(cfg["email_login"], cfg["email_password"])
        mail.select('INBOX')
        result, data = mail.search(None, 'ALL')
        if result == 'OK':
            mail_ids = data[0].split()
            # Берем последние 10 ID писем для проверки
            last_five_mail_ids = mail_ids[-10:]
            for mail_id in last_five_mail_ids[::-1]:
                typ, msg_data = mail.fetch(mail_id, '(RFC822)')
                msg = email.message_from_bytes(msg_data[0][1])
                if email.header.decode_header(msg["Subject"])[0][0].decode() == subject:
                    print("Тест пройден. Письмо найдено.")
                    return 0, ''
            print("Тест не пройден. Среди последних писем не обнаружно отправленное письмо.")
            return 1, "Тест не пройден. Среди последних писем не обнаружно отправленное письмо."
        mail.close()
        mail.logout()
    except Exception as e:
        print_exc()
        print(f"Ошибка при получении тестового письма: {e}")
        return -1, f"Ошибка при получении тестового письма: {e}"


def email_worker(queue):
    while True:
        try:
            email_info = queue.get(timeout=3)  # Ожидание задачи в течение 3 секунд
            send_email(email_info['subject'], email_info['body'], email_info['attachments'], email_info['recipients'])
        except queue.Empty:
            print("Очередь пуста, воркер завершает работу")
            break
        finally:
            queue.task_done()


class ConnectionWindow(QDialog):
    def __init__(self, config):
        super().__init__()
        ui_file = 'UI/connection.ui'
        uic.loadUi(ui_file, self)
        icon = QIcon("UI/icons8-carrier-pigeon-64.png")
        self.config = {}
        self.setWindowIcon(icon)
        email_use_outlook = self.findChild(QRadioButton, 'email_use_outlook')
        email_use_outlook.setChecked(config.get('email_use_outlook', True))
        email_use_outlook.toggled.connect(self.save_params)
        email_use_user = self.findChild(QRadioButton, 'email_use_user')
        email_use_user.setChecked(config.get('email_use_user', False))
        email_use_user.toggled.connect(self.save_params)
        check_connection = self.findChild(QPushButton, 'email_check_connection')
        check_connection.clicked.connect(self.check_connection)
        for widget in self.findChildren(QLineEdit):
            widget.setText(config.get(widget.objectName(), ''))
            widget.textChanged.connect(self.save_params)
        for widget in self.findChildren(QComboBox):
            widget.setCurrentText(config.get(widget.objectName(), ''))
            widget.currentTextChanged.connect(self.save_params)

    def save_params(self):
        for widget in self.findChildren(QLineEdit):
            self.config[widget.objectName()] = widget.text()
        for widget in self.findChildren(QComboBox):
            self.config[widget.objectName()] = widget.currentText()
        rb = self.findChild(QRadioButton, 'email_use_outlook')
        self.config['email_use_outlook'] = rb.isChecked()
        self.config['email_use_user'] = not rb.isChecked()

    def check_connection(self):
        self.save_params()
        if not self.config["email_login"] or not self.config["email_password"]:
            return "Логин или пароль не могут быть пустыми."
        if not self.config["email_out"] or not self.config["email_in"]:
            return "Необходимо указать серверы исходящей и входящей почты."
        subject = "Тест соединения"
        body = "Это тестовое письмо для проверки соединения."
        recipients = [self.config.get('email_login')]
        attachments = []
        res1, message = send_email(subject, body, attachments, recipients, self.config)
        if not res1:
            time.sleep(2)
            res2, message = check_email(subject, self.config)
            if not res2:
                print('Тест подключения пройден')
                QMessageBox.information(self, 'Успех', 'Тест подключения пройден.')
            else:
                print('Не удалось проверить получение письма')
                QMessageBox.warning(self, 'Ошибка', message)
        else:
            print('Не удалось отправить письмо')
            QMessageBox.warning(self, 'Ошибка', message)


## vefqvzadbkvizjqw