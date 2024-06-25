from main_functions import *


class DownloadMasterWindow(QDialog):
    def __init__(self, parent):
        super().__init__(parent=parent)
        ui_file = 'UI/dm.ui'
        uic.loadUi(ui_file, self)
        icon = QIcon("UI/icons8-carrier-pigeon-64.png")
        self.setWindowIcon(icon)
        self.new_config = parent.config.copy()
        self.spinBox_rate_limit = self.findChild(QWidget, 'spinBox_rate_limit')
        self.spinBox_rate_limit.setValue(parent.config.get('limit_rate', 0))
        self.spinBox_rate_limit.valueChanged.connect(self.save_local_params)
        pushButton_add_rule = self.findChild(QWidget, 'pushButton_add_rule')
        pushButton_add_rule.clicked.connect(self.add_new_rule)
        pushButton_delete_rule = self.findChild(QWidget, 'pushButton_delete_rule')
        pushButton_delete_rule.clicked.connect(self.remove_selected_rule)
        buttonBox = self.findChild(QWidget, 'buttonBox')
        buttonBox.button(QDialogButtonBox.Save).clicked.connect(self.update_rule_from_table)
        self.tableWidget = self.findChild(QWidget, 'tableWidget')
        self.refill_rules_table()

    def save_local_params(self):
        self.new_config['limit_rate'] = self.spinBox_rate_limit.value()

    def add_new_rule(self):
        row_position = self.tableWidget.rowCount()
        self.tableWidget.insertRow(row_position)

    def update_rule_from_table(self):
        if self.tableWidget.rowCount() >= 0:
            self.new_config['mail_rules'] = {}
            for row in range(self.tableWidget.rowCount()):
                email = self.tableWidget.item(row, 1).text()
                new_rule = {}
                new_rule['rule_name'] = self.tableWidget.item(row, 0).text() if self.tableWidget.item(row, 0) else ''
                new_rule['subject_contains'] = self.tableWidget.item(row, 2).text() if self.tableWidget.item(row,
                                                                                                             2) else ''
                new_rule['re_filename'] = self.tableWidget.item(row, 3).text() if self.tableWidget.item(row, 3) else ''
                new_rule['re_link'] = self.tableWidget.item(row, 4).text() if self.tableWidget.item(row, 4) else ''
                new_rule['save_folder'] = self.tableWidget.item(row, 5).text() if self.tableWidget.item(row, 5) else ''
                if email in self.new_config['mail_rules']:
                    rules = self.new_config['mail_rules'][email]
                    rules.append(new_rule)
                else:
                    self.new_config['mail_rules'][email] = [new_rule]
        print(self.new_config['mail_rules'])
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
                self.tableWidget.setItem(row_position, 2, QTableWidgetItem(rule['subject_contains']))
                self.tableWidget.setItem(row_position, 3, QTableWidgetItem(rule['re_filename']))
                self.tableWidget.setItem(row_position, 4, QTableWidgetItem(rule['re_link']))
                self.tableWidget.setItem(row_position, 5, QTableWidgetItem(rule['save_folder']))
