import sys
import os
import traceback
from PyQt5.QtWidgets import QMainWindow, QLabel, QLineEdit, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, \
    QCheckBox, QRadioButton, QGroupBox, QComboBox, QPlainTextEdit, QSpinBox, QTimeEdit, QToolButton, QFileDialog,\
    QPushButton, QApplication, QSystemTrayIcon, QAction, QMenu
from PyQt5.QtGui import QIcon
from PyQt5 import uic, QtCore
from PyQt5.QtCore import QTimer, QDateTime, QTime, Qt
from datetime import datetime, timedelta
import time
from threading import Thread, Lock
from main_functions import save_config, get_cert_names, gather_mail, send_mail, validate_email, check_time, add_to_startup, config_path, config, EdoWindow, agregate_edo_messages


class MainWindow(QMainWindow):
    def __init__(self, config):
        super().__init__()
        ui_file = 'UI/main_2.ui'
        uic.loadUi(ui_file, self)
        icon = QIcon("UI/icons8-carrier-pigeon-64.png")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint)
        self.setWindowIcon(icon)
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
                qtobj.setText(param)
                qtobj.textChanged.connect(self.save_params)
            elif obj.startswith('checkBox'):
                qtobj = self.findChild(QWidget, obj)
                qtobj.setChecked(param)
                qtobj.stateChanged.connect(self.save_params)
        toolButton_lineEdit_get_path = self.findChild(QToolButton, 'toolButton_lineEdit_get_path')
        toolButton_lineEdit_get_path.clicked.connect(lambda: self.set_user_dir('lineEdit_get_path'))
        toolButton_lineEdit_put_path = self.findChild(QToolButton, 'toolButton_lineEdit_put_path')
        toolButton_lineEdit_put_path.clicked.connect(lambda: self.set_user_dir('lineEdit_put_path'))
        toolButton_lineEdit_put_path = self.findChild(QToolButton, 'toolButton_lineEdit_csp_path')
        toolButton_lineEdit_put_path.clicked.connect(lambda: self.set_user_dir('lineEdit_csp_path'))
        self.radioButton_periodic = self.findChild(QRadioButton, 'radioButton_periodic')
        self.radioButton_periodic.setChecked(config['radioButton_periodic'])
        self.radioButton_schedule = self.findChild(QRadioButton, 'radioButton_schedule')
        self.radioButton_schedule.setChecked(config['radioButton_schedule'])
        self.radioButton_periodic.toggled.connect(self.save_params)
        self.radioButton_schedule.toggled.connect(self.save_params)
        self.comboBox_certs = self.findChild(QComboBox, 'comboBox_certs')
        cert_names = get_cert_names(os.path.join(self.config['lineEdit_csp_path'], 'certmgr.exe'))
        self.comboBox_certs.addItems(cert_names)
        if config['comboBox_certs'] in cert_names:
            self.comboBox_certs.setCurrentText(config['comboBox_certs'])
        self.comboBox_certs.currentTextChanged.connect(self.save_params)
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
        self.start_scheduler.clicked.connect(self.switch_tasker)
        self.start_manual = self.findChild(QPushButton, 'pushButton_create_mail_now')
        self.start_manual.clicked.connect(lambda: self.send_mail_manual(True))
        self.plainTextEdit_log = self.findChild(QPlainTextEdit, 'plainTextEdit_log')
        pushButton_log = self.findChild(QPushButton, 'pushButton_log')
        pushButton_log.clicked.connect(lambda: os.startfile(os.path.join(config_path, 'log.log')))
        pushButton_edo = self.findChild(QPushButton, 'pushButton_edo')
        pushButton_edo.clicked.connect(self.open_edo_settings)
        autorun = self.findChild(QCheckBox, 'checkBox_autorun')
        autorun.clicked.connect(add_to_startup)
        if config['checkBox_autorun'] and config['checkBox_autostart']:
            mm, ss = config['timeEdit_connecting_delay'].split(':')
            secs = int(mm) * 60 + int(ss)
            self.add_log_message(f'Включен автостарт, работа будет запущена через {secs} секунд')
            self.autostart_timer = QTimer(self)
            self.autostart_timer.timeout.connect(self.switch_tasker)
            self.autostart_timer.start(secs * 1000)
        else:
            self.show()

    def hideEvent(self, event):
        event.ignore()
        self.hide_to_tray()

    def closeEvent(self, event):
        sys.exit(0)

    def open_edo_settings(self):
        self.connection_window = EdoWindow(self.config)
        res = self.connection_window.exec_()
        if res:
            self.config = self.connection_window.config
            save_config(self.config)

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
                self.config[obj] = qtobj.text()
            elif obj.startswith('checkBox'):
                qtobj = self.findChild(QWidget, obj)
                self.config[obj] = qtobj.isChecked()
        self.config['comboBox_certs'] = self.comboBox_certs.currentText()
        self.config['plainTextEdit_body'] = self.plainTextEdit_body.toPlainText()
        self.config['spinBox_max_attach_weight'] = self.spinBox_max_attach_weight.value()
        self.config['timeEdit_send_period'] = self.timeEdit_send_period.time().toString('HH:mm')
        self.config['timeEdit_connecting_delay'] = self.timeEdit_connecting_delay.time().toString('mm:ss')
        self.config['radioButton_periodic'] = self.radioButton_periodic.isChecked()
        self.config['radioButton_schedule'] = self.radioButton_schedule.isChecked()
        save_config(self.config)

    def send_mail_manual(self, manual):
        errors = self.check_fields(manual)
        if errors:
            self.add_log_message('\n'.join(errors))
            return
        messages, names = gather_mail()
        for message_attachments, filenames in zip(messages, names):
            try:
                send_mail(message_attachments, manual=manual)
                sent_files = ', '.join([os.path.basename(fp) for fp in message_attachments])
                self.add_log_message(f'Отправлены файлы: {sent_files}')
                if self.config['checkBox_archive_files']:
                    sent_file_names = ',\n'.join([os.path.basename(fp) for fp in filenames])
                    self.add_log_message(f'В архиве {sent_files} содержатся:\n{sent_file_names}')
            except Exception as e:
                exc_type, exc_value, exc_traceback = sys.exc_info()
                traceback_str = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
                sent_files = ', '.join([os.path.basename(fp) for fp in message_attachments])
                self.add_log_message(f'ОШИБКА отправки файлов: {message_attachments}')
                self.add_log_message(traceback_str)
                traceback.print_exc()
        if self.config.get('checkbox_use_edo', False):
            res = agregate_edo_messages()
            if res:
                self.add_log_message(f'Эл. письма из СО ЭД отправлены')
            else:
                self.add_log_message(f'Ошибка при отправке писем СО ЭД')

    def switch_tasker(self):
        if self.autostart_timer.isActive():
            self.autostart_timer.stop()
        errors = self.check_fields()
        if errors:
            self.add_log_message('\n'.join(errors))
            return
        rb1 = self.findChild(QRadioButton, 'radioButton_periodic')
        rb1.setEnabled(self.daemon_running)
        rb2 = self.findChild(QRadioButton, 'radioButton_schedule')
        rb2.setEnabled(self.daemon_running)
        if self.config['radioButton_periodic']:
            self.switch_periodic()
        elif self.config['radioButton_schedule']:
            self.switch_scheduled()

    def handle_timer_timeout(self):
        self.send_mail_manual(False)
        selected_times = [QTime.fromString(i, 'HH:mm') for i in self.config['lineEdit_schedule'].split(',')]
        current_datetime = QDateTime.currentDateTime()
        current_time = current_datetime.time()
        next_run_times = [self.calculate_time_difference(current_time, time) for time in selected_times]
        min_time_difference = min(next_run_times)
        self.timer_interval = min_time_difference
        self.timer.setInterval(self.timer_interval)

    @staticmethod
    def calculate_time_difference(current_time, selected_time):
        next_run_datetime = QDateTime.currentDateTime()
        next_run_datetime.setTime(selected_time)
        current_date = QDateTime.currentDateTime().date()
        next_run_datetime.setDate(current_date)
        if next_run_datetime < QDateTime.currentDateTime():
            next_run_datetime = next_run_datetime.addDays(1)
        time_difference = QDateTime.currentDateTime().msecsTo(next_run_datetime)
        return time_difference

    def switch_scheduled(self):
        if not self.daemon_running:
            self.timer = QTimer(self)
            self.timer.timeout.connect(self.handle_timer_timeout)
            selected_times = [QTime.fromString(i, 'HH:mm') for i in self.config['lineEdit_schedule'].split(',')]
            current_datetime = QDateTime.currentDateTime()
            current_time = current_datetime.time()
            next_run_times = [self.calculate_time_difference(current_time, time) for time in selected_times]
            min_time_difference = min(next_run_times)
            self.timer_interval = min_time_difference
            self.timer.setInterval(self.timer_interval)
            self.timer.start()
            self.daemon_running = True
            self.start_scheduler.setText('Остановить работу')
            self.add_log_message('Задача запустится в заданные временные точки')
        else:
            self.daemon_running = False
            self.timer.stop()
            self.start_scheduler.setText('Запустить работу в выбранном режиме')
            self.add_log_message('Задача остановлена пользователем')

    def switch_periodic(self):
        if not self.daemon_running:
            self.timer = QTimer(self)
            self.timer.timeout.connect(lambda: self.send_mail_manual(False))
            hours, minutes = self.timeEdit_send_period.time().toString('HH:mm').split(':')
            total_seconds = (int(hours) * 3600 + int(minutes) * 60) * 1000
            self.timer_interval = total_seconds
            self.timer.setInterval(self.timer_interval)
            self.daemon_running = True
            self.send_mail_manual(False)
            self.add_log_message('Задача запущена')
            self.timer.start()
            self.start_scheduler.setText('Остановить работу')
        else:
            self.daemon_running = False
            self.start_scheduler.setText('Запустить работу в выбранном режиме')
            self.timer.stop()
            self.add_log_message('Задача остановлена пользователем')

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
        if self.config['radioButton_schedule'] and not manual:
            valid = True
            periods = self.config['lineEdit_schedule'].split(',')
            for period in periods:
                if not check_time(period):
                    print(period)
                    valid = False
            if not valid:
                errors.append('Некорректно указано время в расписании')
        if self.config['checkbox_use_edo']:
            if not os.path.isdir(self.config['lineedit_input_edo']):
                errors.append('Некорректный путь для исходящих ЭДО')
            if not os.path.isdir(self.config['lineedit_output_edo']):
                errors.append('Некорректный путь для отправленных ЭДО')
        return errors
