import glob
import shutil
import sys
import os
import tempfile
import traceback
from PyQt5.QtWidgets import QMainWindow, QLabel, QLineEdit, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, \
    QCheckBox, QGroupBox, QComboBox, QPlainTextEdit, QSpinBox, QTimeEdit, QToolButton, QFileDialog,\
    QPushButton, QApplication, QSystemTrayIcon, QAction, QMenu
from PyQt5.QtGui import QIcon
from PyQt5 import uic, QtCore
from PyQt5.QtCore import QTimer, QDateTime, QTime, Qt, QThread, pyqtSignal
from datetime import datetime, timedelta
import time
from queue import Queue
from main_functions import (save_config, message_queue, config_file, get_cert_names, gather_mail, send_mail,
                            validate_email, check_time, add_to_startup, config, is_file_locked,
                            send_mail_smtp, SMTPConfigDialog)
import pythoncom
import win32com.client


class MainWindow(QMainWindow):
    def __init__(self, config):
        super().__init__()
        self.monitor_inbox_thread = None
        self.download_queue = Queue()
        self.connection_window = None
        self.edo_window = None
        self.observer = None
        self.soed_available = False
        ui_file = 'UI/main_2.ui'
        uic.loadUi(ui_file, self)
        icon = QIcon("UI/icons8-carrier-pigeon-64.png")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint)
        self.setWindowIcon(icon)
        self.running = False
        # Создаем и запускаем поток мониторинга очереди
        self.queue_thread = QueueMonitorThread()
        self.queue_thread.message_signal.connect(self.add_log_message)
        self.queue_thread.start()
        self.schedule_timers = []
        self.timer = QTimer(self)  # Создание таймера
        def tray_activated(reason):
            if reason == QSystemTrayIcon.DoubleClick:
                self.show_window()
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(icon)
        self.tray.setVisible(True)
        show_action = QAction("Показать", self)
        quit_action = QAction("Выйти", self)
        hide_action = QAction("Скрыть", self)
        show_action.triggered.connect(self.show_window)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(lambda: sys.exit(0))
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)
        self.tray.setContextMenu(tray_menu)
        self.tray.activated.connect(tray_activated)
        self.config = config
        self.daemon_running = False
        for obj, param in self.config.items():
            if obj.startswith('lineEdit'):
                qtobj = self.findChild(QLineEdit, obj)
                if qtobj:
                    qtobj.setText(param)
                    qtobj.textChanged.connect(self.save_params)
            elif obj.startswith('checkBox'):
                qtobj = self.findChild(QWidget, obj)
                if qtobj:
                    qtobj.setChecked(param)
                    qtobj.stateChanged.connect(self.save_params)
        toolButton_lineEdit_get_path = self.findChild(QToolButton, 'toolButton_lineEdit_get_path')
        toolButton_lineEdit_get_path.clicked.connect(lambda: self.set_user_dir('lineEdit_get_path'))
        toolButton_lineEdit_put_path = self.findChild(QToolButton, 'toolButton_lineEdit_put_path')
        toolButton_lineEdit_put_path.clicked.connect(lambda: self.set_user_dir('lineEdit_put_path'))
        self.comboBox_certs = self.findChild(QComboBox, 'comboBox_certs')
        cert_names = get_cert_names(os.path.join(self.config['lineEdit_csp_path'], 'certmgr.exe'))
        self.comboBox_certs.addItems(cert_names)
        if config['comboBox_certs'] in cert_names:
            self.comboBox_certs.setCurrentText(config['comboBox_certs'])
        self.comboBox_certs.currentTextChanged.connect(self.save_params)
        self.pushButton_configure_smtp = self.findChild(QPushButton, 'pushButton_configure_smtp')
        self.pushButton_configure_smtp.clicked.connect(self.configure_smtp)
        self.plainTextEdit_body = self.findChild(QPlainTextEdit, 'plainTextEdit_body')
        self.plainTextEdit_body.setPlainText(config['plainTextEdit_body'])
        self.plainTextEdit_body.textChanged.connect(self.save_params)
        self.spinBox_max_attach_weight = self.findChild(QSpinBox, 'spinBox_max_attach_weight')
        self.spinBox_max_attach_weight.setValue(config['spinBox_max_attach_weight'])
        self.spinBox_max_attach_weight.valueChanged.connect(self.save_params)
        self.timeEdit_send_period = self.findChild(QTimeEdit, 'timeEdit_send_period')
        self.timeEdit_send_period.setTime(QtCore.QTime.fromString(config['timeEdit_send_period'], 'HH:mm'))
        self.timeEdit_send_period.timeChanged.connect(self.save_params)
        self.timeEdit_connecting_delay = self.findChild(QTimeEdit, 'timeEdit_connecting_delay')
        self.timeEdit_connecting_delay.setTime(QtCore.QTime.fromString(config['timeEdit_connecting_delay'], 'mm:ss'))
        self.timeEdit_connecting_delay.timeChanged.connect(self.save_params)
        self.start_scheduler = self.findChild(QPushButton, 'pushButton_start_stop_workers')
        self.start_scheduler.clicked.connect(self.toggleScheduler)
        self.start_manual = self.findChild(QPushButton, 'pushButton_create_mail_now')
        self.start_manual.clicked.connect(lambda: self.send_mail(manual=True))
        self.plainTextEdit_log = self.findChild(QPlainTextEdit, 'plainTextEdit_log')
        pushButton_log = self.findChild(QPushButton, 'pushButton_log')
        pushButton_log.clicked.connect(lambda: os.startfile(os.path.join(os.getcwd(), 'log.log')))
        autorun = self.findChild(QCheckBox, 'checkBox_autorun')
        autorun.clicked.connect(add_to_startup)
        if config['checkBox_autorun'] and config['checkBox_autostart']:
            mm, ss = config['timeEdit_connecting_delay'].split(':')
            secs = int(mm) * 60 + int(ss)
            self.add_log_message(f'Включен автостарт, работа будет запущена через {secs} секунд')
            self.autostart_timer = QTimer(self)
            self.autostart_timer.setSingleShot(True)
            self.autostart_timer.timeout.connect(self.toggleScheduler)
            self.autostart_timer.start(secs * 1000)
        else:
            self.show()

    def toggleScheduler(self):
        if self.running:
            self.stopScheduler()
            self.start_scheduler.setText('Запустить работу по расписанию')
            self.running = False
        else:
            if not self.daemon_running:  # Исключаем повторные запуски
                res = self.startScheduler()
                if res == 0:  # Успешный запуск
                    self.start_scheduler.setText('Остановить работу по расписанию')
                    self.running = True
                else:
                    self.add_log_message('Запуск задач отменен из-за ошибок.')

    def startScheduler(self):
        errors = self.check_fields(False)
        if errors:
            self.add_log_message('\n'.join(errors))
            return -1  # Ошибка валидации
        try:
            self.daemon_running = True  # Флаг активности
            self.schedule_timers.clear()  # Очищаем старые таймеры
            self.timer_index = 1  # Устанавливаем счетчик таймеров
            # Периодический таймер
            if self.config['checkBox_periodic']:
                period = self.timeEdit_send_period.time().hour() * 3600 + self.timeEdit_send_period.time().minute() * 60
                if period <= 0:
                    self.add_log_message('Некорректный период отправки.')
                    return -1
                timer_id = self.timer_index  # Сохраняем номер таймера
                self.timer.setInterval(period * 1000)  # Интервал в миллисекундах
                self.timer.timeout.connect(lambda tid=timer_id, p=period: self.logTimerTriggered(tid, p, periodic=True))
                self.timer.start()
                self.add_log_message(f'Запущен периодический таймер №{timer_id}: каждые {period} секунд.')
                self.timer_index += 1
            # Таймеры по расписанию
            if self.config['checkBox_schedule']:
                now = QTime.currentTime()
                schedule_times = self.config['lineEdit_schedule'].split(',')
                for time_str in schedule_times:
                    time = QTime.fromString(time_str.strip(), "HH:mm")
                    if not time.isValid():
                        self.add_log_message(f'Ошибка формата времени: {time_str}')
                        continue
                    delay_ms = now.msecsTo(time)
                    if delay_ms < 0:
                        delay_ms += 86400000  # Запуск на следующий день
                    timer_id = self.timer_index  # Сохраняем номер таймера
                    timer = QTimer(self)
                    timer.setSingleShot(True)
                    timer.timeout.connect(lambda tid=timer_id, t=time: self.logTimerTriggered(tid, t))
                    QTimer.singleShot(delay_ms, lambda t=timer: t.start())  # Избегаем удаления объекта
                    self.schedule_timers.append(timer)
                    self.add_log_message(
                        f'Запущен таймер №{timer_id} на {time.toString("HH:mm")}, через {delay_ms / 1000:.0f} секунд.')
                    self.timer_index += 1
            self.add_log_message('Работа по расписанию запущена.')
            return 0  # Успешный запуск

        except Exception as e:
            self.add_log_message(f'Ошибка при запуске таймеров: {e}')
            return -1

    def stopScheduler(self):
        self.timer.stop()
        for timer in self.schedule_timers:
            timer.stop()
        self.schedule_timers.clear()
        self.daemon_running = False
        self.add_log_message('Работа по расписанию остановлена.')

    def configure_smtp(self):
        dialog = SMTPConfigDialog(self)
        dialog.exec_()

    def logTimerTriggered(self, timer_index, next_time, periodic=False):
        if periodic:
            next_run = QTime.currentTime().addSecs(next_time)
            self.add_log_message(
                f'Сработал периодический таймер №{timer_index}, следующий запуск в {next_run.toString("HH:mm")}')
        else:
            self.add_log_message(f'Сработал таймер №{timer_index}, следующий запуск в {next_time.toString("HH:mm")}')
        self.send_mail()

    def hideEvent(self, event):
        event.ignore()
        self.hide_to_tray()

    def closeEvent(self, event):
        message_queue.put(None)
        sys.exit(0)

    def hide_to_tray(self):
        self.hide()
        self.tray.showMessage("Приложение свернуто", "Приложение свернуто в трей", QSystemTrayIcon.Information)

    def show_window(self):
        self.showNormal()  # Показываем окно в обычном состоянии (не минимизированном)
        self.activateWindow()

    def set_user_dir(self, lineeditname):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        directory = QFileDialog.getExistingDirectory(self, "Выбрать директорию", options=options)
        if directory != '':
            lineedit = self.findChild(QLineEdit, lineeditname)
            lineedit.setText(fr'{directory}')

    def save_params(self):
        for obj, param in self.config.items():
            if obj.startswith('lineEdit'):
                qtobj = self.findChild(QLineEdit, obj)
                if qtobj:
                    self.config[obj] = qtobj.text()
            elif obj.startswith('checkBox'):
                qtobj = self.findChild(QWidget, obj)
                if qtobj:
                    self.config[obj] = qtobj.isChecked()
        self.config['comboBox_certs'] = self.comboBox_certs.currentText()
        self.config['plainTextEdit_body'] = self.plainTextEdit_body.toPlainText()
        self.config['spinBox_max_attach_weight'] = self.spinBox_max_attach_weight.value()
        self.config['timeEdit_send_period'] = self.timeEdit_send_period.time().toString('HH:mm')
        self.config['timeEdit_connecting_delay'] = self.timeEdit_connecting_delay.time().toString('mm:ss')
        save_config(config_file, self.config)

    def send_mail(self, manual=False):
        errors = self.check_fields(manual)
        if errors:
            self.add_log_message('\n'.join(errors))
            return
        messages, names = gather_mail()
        if not messages:
            self.add_log_message(f'Файлы для отправки не найдены')
        for message_attachments, filenames in zip(messages, names):
            try:
                if self.config['checkBox_use_outlook']:
                    result, error = send_mail(message_attachments, manual=manual)
                else:
                    self.add_log_message('Отправка на SMTP сервер, окно программы зависнет на время отправки.')
                    result, error = send_mail_smtp(message_attachments)
                if result:
                    sent_files = ', '.join([os.path.basename(fp) for fp in message_attachments])
                    self.add_log_message(f'Отправлены файлы: {sent_files}')
                    if self.config['checkBox_archive_files']:
                        sent_file_names = ',\n'.join([os.path.basename(fp) for fp in filenames])
                        self.add_log_message(f'В архиве {sent_files} содержатся:\n{sent_file_names}')
                    self.tray.showMessage("Успех", "Документы были успешно отправлены", QSystemTrayIcon.Information)
                else:
                    self.add_log_message(f'Ошибка: {error}')
                    self.tray.showMessage("Ошибка", f"Ошибка при отправке:{error}", QSystemTrayIcon.Information)
            except Exception as e:
                exc_type, exc_value, exc_traceback = sys.exc_info()
                traceback_str = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
                self.add_log_message(f'ОШИБКА отправки файлов: {message_attachments}')
                self.add_log_message(traceback_str)
                traceback.print_exc()

    def add_log_message(self, message):
        current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{current_datetime} - {message}"
        self.plainTextEdit_log.appendPlainText(log_entry)
        log_path = os.path.join(os.path.dirname(sys.argv[0]), 'log.log')
        with open(log_path, "a") as log_file:
            log_file.write(log_entry + "\n")

    def check_fields(self, manual=False):
        errors = []
        if not os.path.isdir(self.config['lineEdit_get_path']):
            errors.append('Некорректный путь для исходящих')
        if not os.path.isdir(self.config['lineEdit_put_path']):
            errors.append('Некорректный путь для отправленных')
        if os.path.normpath(self.config['lineEdit_get_path']) == os.path.normpath(self.config['lineEdit_put_path']):
            errors.append('Пути исходящих и отправленных не могут совпадать')
        if not (os.path.isfile(os.path.join(self.config['lineEdit_csp_path'], 'certmgr.exe')) and os.path.isfile(os.path.join(self.config['lineEdit_csp_path'], 'csptest.exe'))):
            errors.append('Некорректный путь к csptest')
        if self.config['comboBox_certs'] == "Сертификат не выбран" and self.config['checkBox_use_encryption']:
            errors.append('Не выбран сертификат для подписи')
        if self.config['lineEdit_recipients']:
            valid = False
            for i in self.config['lineEdit_recipients'].split(";"):
                if validate_email(i):
                    valid = True
            if not valid:
                errors.append('Не обнаружен валидный email адрес')
        else:
            errors.append('Должен быть указан хотя бы один адресат')
        if self.config['checkBox_schedule'] and not manual:
            valid = True
            periods = self.config['lineEdit_schedule'].split(',')
            for period in periods:
                if not check_time(period):
                    valid = False
            if not valid:
                errors.append('Некорректно указано время в расписании')
        return errors


class QueueMonitorThread(QThread):
    message_signal = pyqtSignal(str)

    def run(self):
        while True:
            message = message_queue.get()
            if message is None:
                break
            self.message_signal.emit(message)

