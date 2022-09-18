import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

from numpy import isnan
from pandas import DataFrame, concat, read_excel, to_datetime
from PySide6.QtCore import QAbstractTableModel, QDate, Qt
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox,
                               QDateEdit, QDialog, QFileDialog, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QMessageBox,
                               QPushButton, QTableView, QVBoxLayout, QWidget, QComboBox)

def load_excel_data():
    df = read_excel(f'bin/{datetime.now().year}{datetime.now().month:02d}.xlsx',
                    dtype={
                        'pid': int,
                        'Serial_Number': str,
                        'Box': int,
                        'Cell': int,
                        'Place_Date': datetime,
                        'Report_Generated': bool,
                        'Takeout_Date': datetime
                    }
                    )
    df = df.loc[df['Serial_Number'] != "init",].copy().reset_index(drop=True)
    return df

def init_excel_data():
    last_month = datetime.now().replace(day=1) - timedelta(days=1)
    if Path(f'bin/{last_month.year}{last_month.month:02d}.xlsx').exists():
        data = read_excel(Path(f'bin/{last_month.year}{last_month.month:02d}.xlsx'))
        data = data.loc[data['Takeout_Date'].isna(), ].copy().reset_index(drop=True)
    else:
        data = DataFrame({'pid': 0, 'Serial_Number': 'init',  'Box': 0,  'Cell': 0, 
                        'Place_Date': to_datetime("1970-01-01 00:00:00"), 'Report_Generated': True, 'Takeout_Date': to_datetime("1970-01-01 00:00:00")}, 
                        index=[0])
    if data.shape[0] == 0:
        data = DataFrame({'pid': 0, 'Serial_Number': 'init',  'Box': 0,  'Cell': 0, 
                        'Place_Date': to_datetime("1970-01-01 00:00:00"), 'Report_Generated': True, 'Takeout_Date': to_datetime("1970-01-01 00:00:00")}, 
                        index=[0])
    data.to_excel(f'bin/{datetime.now().year}{datetime.now().month:02d}.xlsx', index=False)

def load_settings():
    with open('bin/settings.json', 'r') as f:
        data = json.load(f)
        box_amount = data['box_amount']
        cell_amount = data['cell_amount']
        empty_string = str(data['empty_string'])
        full_string = str(data['full_string'])
    return box_amount, cell_amount, empty_string, full_string

def chunk(list, n):
    for i in range(0, len(list), n):
        yield list[i:i+n]

def initBoxStatus(df, box_amount, cell_amount):
    df = df.loc[df['Takeout_Date'].isna()]
    box_status = [i for i in (empty_string * box_amount * cell_amount)]
    for _, row in df.iterrows():
        box_status[((row['Box'] - 1) * cell_amount) + (row['Cell']) - 1] = full_string
    return box_status

def insert_to_empty_cell(box_status, pattern, replacement):
    pattern = str(pattern)
    replacement = str(replacement)
    if pattern not in box_status:
        return "full"
    if replacement in box_status:
        from_beginning = False
        if not from_beginning:
            last_occur = len(box_status) - 1 - box_status[::-1].index(replacement)
            if last_occur == len(box_status) - 1:
                from_beginning = True
            else:
                i = last_occur + 1
                box_status[i] = replacement
        if from_beginning:
            i = box_status.index(pattern)
            box_status[i] = replacement
    else:
        i = 0
        box_status[i] = replacement
    return i + 1


def get_today():
    today = datetime.now()
    return {'year': today.year, 'month': today.month, 'day': today.day}


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.dataframe = load_excel_data()
        self.box_status = initBoxStatus(self.dataframe, box_amount, cell_amount)
        self.today = get_today()
        self.initializeUI()


    def initializeUI(self):
        self.setFixedSize(800, 600)
        self.center()
        self.setWindowTitle("Simple Warehose Management")

        self.initTable()
        self.setUpMainWindow()

        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def setUpMainWindow(self):
        # Navigate Buttons
        sn_label = QLabel("[   ADD    ] SN:", self)
        self.sn = QLineEdit(self)
        self.sn.returnPressed.connect(self.insert_new_item)
        add_new_item_button = QPushButton("Add item")        
        add_new_item_button.clicked.connect(self.insert_new_item)
        takeout_button = QPushButton("Search item")
        takeout_button.clicked.connect(self.takeout_item)
        report_generate_button = QPushButton("Report Generate")
        report_generate_button.clicked.connect(self.report_generate)

        search_sn_label = QLabel("[SEARCH] SN:", self)
        self.search_sn = QLineEdit(self)
        self.year_input = QComboBox(self)
        self.year_input.addItems([str(item) for item in list(range(2020,datetime.now().year+1))])
        self.year_input.setCurrentText(str(self.today['year']))
        self.month_input = QComboBox(self)
        self.month_input.addItems([str(item) for item in list(range(1,13))])
        self.month_input.setCurrentText(str(self.today['month']))
        # self.date_input = QDateEdit(calendarPopup=True)
        # self.date_input.setDate(QDate.fromString(datetime.now().strftime("%Y-%m"), "yyyy-MM"))
        self.date_checkbox = QCheckBox("Date:")

        save_button = QPushButton("Save")
        save_button.clicked.connect(lambda: self.save_data(show_box=True))

        delete_row_button = QPushButton("Delete")
        delete_row_button.clicked.connect(self.delete_row)

        # layouts
        button_box = QVBoxLayout()

        add_item_box = QHBoxLayout()
        add_item_box.addWidget(sn_label)
        add_item_box.addWidget(self.sn)
        add_item_box.addWidget(add_new_item_button)

        search_box = QHBoxLayout()
        search_box.addWidget(search_sn_label)
        search_box.addWidget(self.search_sn)
        search_box.addWidget(self.date_checkbox)
        search_box.addWidget(self.year_input)
        search_box.addWidget(self.month_input)
        # search_box.addWidget(self.date_input)
        search_box.addWidget(takeout_button)

        button_box.addLayout(add_item_box)
        button_box.addLayout(search_box)
        button_box.addWidget(report_generate_button)

        view_table_box = QHBoxLayout()
        view_table_box.addWidget(self.table)

        main_v_box = QVBoxLayout()
        main_v_box.addLayout(view_table_box)
        main_v_box.addWidget(delete_row_button)
        main_v_box.addLayout(button_box)
        main_v_box.addWidget(save_button)

        self.setLayout(main_v_box)

    def initTable(self):
        self.model = TableModel(self.dataframe.query('Report_Generated == False').copy().reset_index(drop=True))

        # Table Init
        self.table = QTableView()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ContiguousSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setModel(self.model)
        self.table.scrollToBottom()

        for column in [0, 6, 7]:
            self.table.hideColumn(column)

    def insert_new_item(self):
        sn_text = self.sn.text()

        if not sn_text:
            QMessageBox.question(
                self, 
                'Message', 
                "Please Enter Serial_Number.", 
                QMessageBox.StandardButton.Ok,
                QMessageBox.StandardButton.Ok
            )


        if sn_text:
            insert_result = insert_to_empty_cell(self.box_status, empty_string, full_string)

            if insert_result == 'full':
                error_msg = QMessageBox()
                error_msg.setIcon(QMessageBox.Icon.Critical)
                error_msg.setText("All Box Full")
                error_msg.setWindowTitle("Full")
                error_msg.exec()
                self.sn.setText('')
                return


            cell = insert_result % cell_amount

            if cell == 0:
                cell = cell_amount
                box = insert_result // cell_amount
            else:
                box = insert_result // cell_amount + 1
            
            last_pid = self.dataframe['pid'].max()
            if isnan(last_pid):
                last_pid = 0
            new_df = DataFrame({
                'pid': (last_pid + 1),
                'Serial_Number': sn_text,
                'Box': box,
                'Cell': cell,
                'Place_Date': datetime.now().replace(microsecond=0),
                'Report_Generated': False,
                'Takeout_Date': None}, index=[0])
            self.dataframe = concat([self.dataframe, new_df], axis=0, ignore_index=True)
            self.model = TableModel(self.dataframe.query("Report_Generated == False").copy().reset_index(drop=True))
            self.table.setModel(self.model)
            self.box_status = initBoxStatus(self.dataframe[self.dataframe['Takeout_Date'].isna()].copy().reset_index(drop=True), box_amount, cell_amount)
            self.sn.setText(None)
            self.table.scrollToBottom()

    def closeEvent(self, event):
        answer = QMessageBox.question(self, "Quit?",
                                      "Save before Quit?",
                                      QMessageBox.StandardButton.Cancel |
                                      QMessageBox.StandardButton.No |
                                      QMessageBox.StandardButton.Yes,
                                      QMessageBox.StandardButton.Yes)
        if answer == QMessageBox.StandardButton.Yes:
            self.save_data(show_box=False)
            event.accept()
        if answer == QMessageBox.StandardButton.No:
            event.accept()
        if answer == QMessageBox.StandardButton.Cancel:
            event.ignore()

    def takeout_item(self):
        #(取消) 增加搜尋之前月份的功能
        search_sn = None
        search_date = None

        if not self.date_checkbox.isChecked() and not self.search_sn.text():
            error_msg = QMessageBox()
            error_msg.setIcon(QMessageBox.Icon.Critical)
            error_msg.setText("Please SET Search Condition.")
            error_msg.setWindowTitle("SET Search Condition")
            error_msg.exec()
            return

        if self.search_sn.text():
            search_sn = self.search_sn.text()

        if not self.date_checkbox.isChecked():
            search_window = TakeoutitemWindow(self.dataframe, search_sn, search_date)
        elif self.date_checkbox.isChecked():
            search_window = TakeoutitemWindow(self.dataframe, search_sn, search_date)
        # else:
        #     previous_excel = read_excel(
        #         f'bin/{self.year_input.currentText()}{self.month_input.currentText().zfill(2)}.xlsx',
        #         dtype={
        #                 'pid': int,
        #                 'Serial_Number': str,
        #                 'Box': int,
        #                 'Cell': int,
        #                 'Place_Date': datetime,
        #                 'Report_Generated': bool,
        #                 'Takeout_Date': datetime
        #             }
        #     )
        #     search_window = TakeoutitemWindow(previous_excel, search_sn, search_date)
        search_window.exec()
        self.update_data_model(update_model=True)


    def report_generate(self):
        answer = QMessageBox.question(
            self, 
            "Save?",
            "Save before generate report?",
            QMessageBox.StandardButton.Cancel |QMessageBox.StandardButton.No |QMessageBox.StandardButton.Yes,
            QMessageBox.StandardButton.Yes
        )
        
        if answer == QMessageBox.StandardButton.Cancel:
            return
        if answer == QMessageBox.StandardButton.Yes:
            self.save_data(show_box=False)
        
        report_window = GenerateReportWindow(self.dataframe)
        report_window.exec()

        self.update_data_model(update_model=True)



    def save_data(self, show_box):
        # date = self.date_input.date().toString("yyyy-MM-dd")
        # self.dataframe.loc[self.dataframe['Place_Date'].dt.strftime('%Y-%m-%d') == date, 'Report_Generated'] = True
        self.dataframe.loc[self.dataframe['Report_Generated'] == False, 'Report_Generated'] = True
        self.dataframe.to_excel(f'bin/{datetime.now().year}{datetime.now().month:02d}.xlsx', index=False)
        if show_box:
            QMessageBox.information(
                self,
                'Message',
                "Data Saved!",
                QMessageBox.StandardButton.Ok,
                QMessageBox.StandardButton.Ok
            )
        self.update_data_model(update_model=True)

    def delete_row(self):
        pid_list = []
        indices = self.table.selectionModel().selectedRows()
        for index in sorted(indices):
            pid_list.append(int(self.table.model().index(index.row(), 0).data()))
        for i in pid_list:
            box = self.dataframe.loc[self.dataframe['pid']== i, 'Box'].values[0]
            cell = self.dataframe.loc[self.dataframe['pid'] == i, 'Cell'].values[0]
            cell_pos = ((box -1) * cell_amount) + (cell - 1)
            self.box_status[cell_pos] = empty_string
        self.dataframe = self.dataframe.loc[~self.dataframe['pid'].isin(pid_list), ].copy().reset_index(drop=True)
        self.model = TableModel(self.dataframe.query("Report_Generated == False").copy().reset_index(drop=True))
        self.table.setModel(self.model)

    def update_data_model(self, update_model):
        self.dataframe = load_excel_data()
        self.box_status = initBoxStatus(self.dataframe, box_amount, cell_amount)
        if update_model:
            self.model = TableModel(self.dataframe.query("Report_Generated == False").copy().reset_index(drop=True))
            self.table.setModel(self.model)
        pass

class GenerateReportWindow(QDialog):
    def __init__(self, df):
        super().__init__()
        self.dataframe = df
        self.initializeUI()

    def initializeUI(self):
        self.setFixedSize(400, 160)
        self.setWindowTitle("Report")
        self.setUpWindow()

    def setUpWindow(self):
        self.date_input = QDateEdit(calendarPopup=True)
        self.date_input.setDate(QDate.fromString(datetime.now().strftime("%Y-%m-%d"), "yyyy-MM-dd"))

        close_button = QPushButton("Close")
        close_button.clicked.connect(lambda: self.close())
        generate_button = QPushButton("Generate")
        generate_button.clicked.connect(self.generate_report)

        main_box = QVBoxLayout()
        main_box.addWidget(self.date_input)
        main_box.addWidget(generate_button)
        main_box.addWidget(close_button)

        self.setLayout(main_box)

    def generate_report(self):
        name, _ = QFileDialog.getSaveFileName(self, 'Save File', f"{self.date_input.date().toString('yyyyMMdd')}_.xlsx", "Excel Files (*.xlsx)")
        if not name:
            return
        date = self.date_input.date().toString("yyyy-MM-dd")
        # self.dataframe.loc[self.dataframe['Place_Date'].dt.strftime('%Y-%m-%d') == date, 'Report_Generated'] = True
        query_data = self.dataframe[self.dataframe['Place_Date'].dt.strftime('%Y-%m-%d') == date]
        query_data.to_excel(name, index = False)
        self.dataframe.to_excel(f'bin/{datetime.now().year}{datetime.now().month:02d}.xlsx', index=False)
        QMessageBox.information(
            self,
            'Message',
            "Report Generated!",
            QMessageBox.StandardButton.Ok,
            QMessageBox.StandardButton.Ok
        )

class TakeoutitemWindow(QDialog):
    def __init__(self, df, sn=None, date=None):
        super().__init__()
        self.dataframe = df
        self.prepared_df = (df
            .query(f'Takeout_Date == "" | Takeout_Date.isna()')
            .query(f'Report_Generated == True')
        )
        
        if sn:
            self.prepared_df = self.prepared_df[self.prepared_df['Serial_Number'].str.match(sn)].copy().reset_index(drop=True)
        if date:
            self.date = date
            self.prepared_df = self.prepared_df[self.prepared_df['Place_Date'].dt.strftime('%Y-%m-%d') == self.date.strftime('%Y-%m-%d')].copy().reset_index(drop=True)
        self.model = TableModel(self.prepared_df)
        self.initializeUI()

    def initializeUI(self):
        self.setFixedSize(800, 320)
        self.setWindowTitle("Search item")
        self.setUpWindow()

    def setUpWindow(self):
        self.table = QTableView()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ContiguousSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setModel(self.model)

        close_button = QPushButton("Close")
        close_button.clicked.connect(lambda: self.close())
        takeout_button = QPushButton("Take Out")
        takeout_button.clicked.connect(self.take_out)
        save_button = QPushButton("Save")
        save_button.clicked.connect(self.file_save)

        button_box = QHBoxLayout()
        button_box.addWidget(takeout_button)
        button_box.addWidget(close_button)


        main_box = QVBoxLayout()
        main_box.addWidget(self.table)
        main_box.addLayout(button_box)
        main_box.addWidget(save_button)

        self.setLayout(main_box)

    def file_save(self):
        for _, item in self.model._data.iterrows():
            self.dataframe.loc[self.dataframe['pid'] == item['pid'], 'Takeout_Date'] = item['Takeout_Date']
        # self.dataframe.to_excel(f'bin/{self.date.year}{self.date.month:02d}.xlsx', index = False)
        self.dataframe.to_excel(f'bin/{datetime.now().year}{datetime.now().month:02d}.xlsx', index = False)
        QMessageBox.information(
            self,
            'Message',
            "Data Saved!",
            QMessageBox.StandardButton.Ok,
            QMessageBox.StandardButton.Ok
        )


    def take_out(self):
        self.pid_list = []
        indices = self.table.selectionModel().selectedRows()
        for index in sorted(indices):
            pid = int(self.table.model().index(index.row(), 0).data())
            self.pid_list.append(pid)
        self.model.setTakeout_Date(self.pid_list, datetime.now().replace(microsecond=0))


class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data.copy().reset_index(drop=True)
        self._data.index += 1

    def data(self, index, role):
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, index):
        return self._data.shape[1]
    
    def setTakeout_Date(self, pid, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            self._data.loc[(self._data['pid'].isin(pid) & self._data['Report_Generated'] == True), 'Takeout_Date'] = value
        else:
            return False
        return True



    def headerData(self, section, orientation, role):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._data.columns[section])

            if orientation == Qt.Orientation.Vertical:
                return str(self._data.index[section])


if __name__ == '__main__':
    Path('bin').mkdir(exist_ok=True)
    if not Path(f'bin/{datetime.now().year}{datetime.now().month:02d}.xlsx').exists():
        init_excel_data()
    box_amount, cell_amount, empty_string, full_string = load_settings()
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
