from main_functions import *


class DownloadMasterWindow(QDialog):
    def __init__(self, parent):
        super().__init__(parent=parent)
        ui_file = 'UI/dm.ui'
        uic.loadUi(ui_file, self)
        icon = QIcon("UI/icons8-carrier-pigeon-64.png")
        self.setWindowIcon(icon)
        self.new_config = config.copy()
        self.spinBox_rate_limit = self.findChild(QWidget, 'spinBox_rate_limit')
        self.spinBox_rate_limit.setValue(config.get('rate_limit', 0))
        self.spinBox_rate_limit.valueChanged.connect(self.save_local_params)
        pushButton_add_rule = self.findChild(QWidget, 'pushButton_add_rule')
        pushButton_add_rule.clicked.connect(self.add_new_rule)
        pushButton_delete_rule = self.findChild(QWidget, 'pushButton_delete_rule')
        pushButton_delete_rule.clicked.connect(self.remove_selected_rule)
        self.tableWidget = self.findChild(QWidget, 'tableWidget')
        self.refill_rules_table()
        self.tableWidget.cellChanged.connect(self.update_rule_from_table)

    def save_local_params(self):
        self.new_config['rate_limit'] = self.spinBox_rate_limit.value()

    def add_new_rule(self):
        dialog = AddRuleDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            new_rule = dialog.get_new_rule()
            if not new_rule['email'] or not new_rule['rule_name'] or not new_rule['re_link'] or not new_rule['save_folder']:
                QMessageBox.warning(self, 'Ошибка', 'Одно из обязательных полей(почта, название, regex ссылки, папка сохранения) не заполнено')
                return
            if not os.path.exists(new_rule['save_folder']):
                try:
                    os.mkdir(new_rule['save_folder'])
                except:
                    print_exc()
                    QMessageBox.warning(self, 'Ошибка', 'Заданная директория не существует и ее не удается создать.')
            email = new_rule.pop('email')
            if email in self.new_config['mail_rules']:
                self.new_config['mail_rules'][email].append(new_rule)
            else:
                self.new_config['mail_rules'][email] = [new_rule]
            self.refill_rules_table()

    def update_rule_from_table(self, row, column):
        email = self.tableWidget.item(row, 1).text()
        rule_name = self.tableWidget.item(row, 0).text()
        if email in self.new_config['mail_rules']:
            rules = self.new_config['mail_rules'][email]
            for rule in rules:
                if rule['rule_name'] == rule_name:
                    rule['rule_name'] = self.tableWidget.item(row, 0).text()
                    rule['re_subject'] = self.tableWidget.item(row, 2).text()
                    rule['re_body'] = self.tableWidget.item(row, 3).text()
                    rule['re_link'] = self.tableWidget.item(row, 4).text()
                    rule['filename'] = self.tableWidget.item(row, 5).text()
                    rule['save_folder'] = self.tableWidget.item(row, 6).text()
                    break

    def remove_selected_rule(self):
        row = self.tableWidget.currentRow()
        if row >= 0:
            email = self.tableWidget.item(row, 1).text()
            rule_name = self.tableWidget.item(row, 0).text()
            if email in self.new_config['mail_rules']:
                rules = self.new_config['mail_rules'][email]
                for rule in rules:
                    if rule['rule_name'] == rule_name:
                        rules.remove(rule)
                        break
                if not rules:
                    del self.new_config['mail_rules'][email]
            self.tableWidget.removeRow(row)

    def refill_rules_table(self):
        self.tableWidget.setRowCount(0)  # Очистить таблицу перед заполнением
        for email, rules in self.new_config['mail_rules'].items():
            for rule in rules:
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)
                self.tableWidget.setItem(row_position, 0, QTableWidgetItem(rule['rule_name']))
                self.tableWidget.setItem(row_position, 1, QTableWidgetItem(email))
                self.tableWidget.setItem(row_position, 2, QTableWidgetItem(rule['re_subject']))
                self.tableWidget.setItem(row_position, 3, QTableWidgetItem(rule['re_body']))
                self.tableWidget.setItem(row_position, 4, QTableWidgetItem(rule['re_link']))
                self.tableWidget.setItem(row_position, 5, QTableWidgetItem(rule['filename']))
                self.tableWidget.setItem(row_position, 6, QTableWidgetItem(rule['save_folder']))


class AddRuleDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить правило. *-обязательные поля")

        self.layout = QVBoxLayout(self)
        self.form_layout = QFormLayout()

        self.email_input = QLineEdit(self)
        self.rule_name_input = QLineEdit(self)
        self.re_subject_input = QLineEdit(self)
        self.re_body_input = QLineEdit(self)
        self.re_link_input = QLineEdit(self)
        self.filename_input = QComboBox(self)
        self.filename_input.addItems(['Тема', "Тело", "Время"])
        self.save_folder_input = QLineEdit(self)
        self.form_layout.addRow("Email*", self.email_input)
        self.form_layout.addRow("Название правила*", self.rule_name_input)
        self.form_layout.addRow("Тема содержит", self.re_subject_input)
        self.form_layout.addRow("Regex назв. файла", self.re_body_input)
        self.form_layout.addRow("Regex ссылки*", self.re_link_input)
        self.form_layout.addRow("Название файла*", self.filename_input)
        self.form_layout.addRow("Папка для сохранения*", self.save_folder_input)
        self.layout.addLayout(self.form_layout)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)

    def get_new_rule(self):
        return {
            'email': fr'{ self.email_input.text()}',
            'rule_name': fr'{ self.rule_name_input.text()}',
            're_subject': fr'{ self.re_subject_input.text()}',
            're_body': fr'{ self.re_body_input.text()}',
            're_link': fr'{ self.re_link_input.text()}',
            'filename': fr'{ self.filename_input.currentText()}',
            'save_folder': fr'{ self.save_folder_input.text()}'
        }
