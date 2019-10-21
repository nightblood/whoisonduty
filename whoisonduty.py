# -*- coding: utf-8 -*-

import sys

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from functools import reduce
import datetime
from openpyxl import Workbook
import openpyxl
from openpyxl.styles import Font, colors, Alignment
import random
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side
from openpyxl.comments import Comment
import json
import urllib
from urllib import request
import math


class MainWindow(QWidget):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        MainWindow.setFixedSize(self, 760, 500)
        self.setWindowTitle('自动排班工具')
        self.setGeometry(500, 300, 300, 200)

        self.export_worker = ''
        self.start_dt = ''
        self.end_dt = ''
        self.staff = []
        self.init_data()
        self.init_window()

    def init_data(self):
        self.init_staff()
        self.init_date()
        self.next_holiday = self.get_next_holiday()

    def get_next_holiday(self):
        try:
            url = "http://timor.tech/api/holiday/tts/next"
            resp = request.urlopen(url, data=None, timeout=5)
            data = json.loads(resp.read())
            if data["code"] == 0:
                print(data["tts"])
                return data["tts"]
        except Exception as e:
            print(e)
            return ""

    def init_window(self):
        self.main_layout = QVBoxLayout(self)
        self.l_desc = QLabel('将所有名字填入下方输入框，并以空格隔开。点击【排班】按钮将随机排列名字。')
        self.pte_staff = QPlainTextEdit(self.get_staff_str())
        self.pte_staff.setFont(QFont('Roman times', 14, QFont.Bold))
        self.lb_holiday_desc = QLabel(self.next_holiday)
        self.l_start_dt = QLabel('开始日期')
        self.l_end_dt = QLabel('结束日期')

        self.dte_start = QDateTimeEdit(QDate.currentDate(), self)
        self.dte_start.dateTimeChanged.connect(lambda: self.set_time(1))
        self.dte_start.setDisplayFormat('yyyy年MM月dd日')
        self.dte_start.setCalendarPopup(True)
        self.dte_end = QDateTimeEdit(QDate.currentDate().addDays(7 * 40), self)
        self.dte_end.dateTimeChanged.connect(lambda: self.set_time(2))
        self.dte_end.setDisplayFormat('yyyy年MM月dd日')
        self.dte_end.setCalendarPopup(True)

        self.layout_start_dt = QSplitter(Qt.Horizontal)
        self.layout_end_dt = QSplitter(Qt.Horizontal)

        self.layout_func = QSplitter(Qt.Horizontal)
        self.btn_sort = QPushButton('排班')
        self.btn_export = QPushButton('导出')
        self.layout_func.addWidget(self.btn_sort)
        self.layout_func.addWidget(self.btn_export)

        self.l_weeks = QLabel('共 40 周')

        self.layout_start_dt.addWidget(self.l_start_dt)
        self.layout_start_dt.addWidget(self.dte_start)
        self.layout_end_dt.addWidget(self.l_end_dt)
        self.layout_end_dt.addWidget(self.dte_end)

        self.btn_export.clicked.connect(self.on_export)

        self.btn_sort.clicked.connect(self.on_sort)

        self.main_layout.addWidget(self.l_desc)
        self.main_layout.addWidget(self.pte_staff)
        self.main_layout.addWidget(self.layout_start_dt)
        self.main_layout.addWidget(self.layout_end_dt)
        self.main_layout.addWidget(self.l_weeks)
        self.main_layout.addWidget(self.lb_holiday_desc)
        self.main_layout.addWidget(self.layout_func)

    def set_time(self, btn_idx):
        if btn_idx == 1:
            self.start_dt = self.dte_start.text()
        else:
            self.end_dt = self.dte_end.text()
        print(self.start_dt, self.end_dt)
        self.l_weeks.setText('共' + str(self.get_delta_weeks()) + '周')

    def on_export(self):
        self.staff = self.get_staff_list_from_widget()
        weeks = self.get_delta_weeks()
        if self.staff is None or len(self.staff) == 0:
            QMessageBox.question(self, '提示', '请先输入人员，再进行排班', QMessageBox.Yes)
            return
        if weeks <= 0:
            QMessageBox.question(self, '提示', '日期选择异常，请重新选择日期', QMessageBox.Yes)
            return
        self.export_worker = ExportWorker(self.staff, weeks, self.start_dt)
        self.export_worker.sig_complete.connect(self.export_complete)
        self.export_worker.start()

    def export_complete(self, desc):
        try:
            QMessageBox.question(self, '提示', desc, QMessageBox.Yes)
        except Exception as e:
            print(e)

    def get_delta_weeks(self):
        start = datetime.datetime.strptime(self.start_dt, '%Y年%m月%d日')
        end = datetime.datetime.strptime(self.end_dt, '%Y年%m月%d日')
        weeks = (end - start).days / 7
        if (end - start).days % 7 > 0:
            weeks += 1
        return weeks

    def on_sort(self):
        self.staff = self.get_staff_list_from_widget()
        if self.staff is None or len(self.staff) == 0:
            QMessageBox.question(self, '提示', '请先输入人员，再进行排班', QMessageBox.Yes)
            return
        self.sort_worker = SortWorker(self.staff)
        self.sort_worker.sig_complete.connect(self.sort_complete)
        self.sort_worker.start()

    def sort_complete(self, data):
        self.staff = data.split(' ')
        self.pte_staff.setPlainText(self.get_staff_str())

    def init_staff(self):
        try:
            staff_file = open('staff.txt', 'r', encoding='utf-8')
            data = staff_file.read()
            print(data)
            if data is not None and len(data.strip()) != 0:
                for item in data.strip().split(' '):
                    if len(item) != 0:
                        self.staff.append(item)
            staff_file.close()
        except Exception as e:
            print(e)

    def init_date(self):
        now = datetime.datetime.now()
        self.start_dt = now.strftime('%Y{y}%m{m}%d{d}').format(y='年', m='月', d='日')
        # 40周
        self.end_dt = (now + datetime.timedelta(days=280)).strftime('%Y{y}%m{m}%d{d}').format(y='年', m='月', d='日')

    def get_staff_str(self):
        if self.staff is None or len(self.staff) == 0:
            return ''
        return reduce(lambda x, y: x + ' ' + y, self.staff)

    def get_staff_list_from_widget(self):
        res = []
        str = self.pte_staff.toPlainText()
        if str is None or len(str.strip()) == 0:
            return res
        staffs = str.strip().split(' ')
        print(staffs)
        for item in staffs:
            if len(item) != 0:
                res.append(item)
        return res


class SortWorker(QThread):
    sig_complete = pyqtSignal(str)

    def __init__(self, staff):
        super().__init__()
        self.staff = staff

    def run(self):
        length = len(self.staff)
        res = []
        for idx in range(length):
            item = random.choice(self.staff)
            self.staff.remove(item)
            res.append(item)
        print(res)
        self.sig_complete.emit(reduce(lambda x, y: x + ' ' + y, res))


class ExportWorker(QThread):
    sig_complete = pyqtSignal(str)

    def __init__(self, staffs, weeks, start_dt):
        super().__init__()
        self.staff = staffs
        self.weeks = weeks
        self.start_dt = start_dt

    def run(self):
        try:
            wb = Workbook()
            ws = wb.active

            col_name = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
            first_row = ['日期', '周一', '周二', '周三', '周四', '周五', '周六', '周日']
            ws.row_dimensions[1].height = 40  # 行高
            ws.row_dimensions[2].height = 30  # 行高

            border = Border(left=Side(border_style='thin', color='4f5555'), right=Side(border_style='thin', color='4f5555'),
                            top=Side(border_style='thin', color='4f5555'), bottom=Side(border_style='thin', color='4f5555'))
            pattern_fill = PatternFill("solid", fgColor="feeeed")  # 星期行和日期列背景色
            bg_weekend = PatternFill("solid", fgColor="fedcbd")  # 周末值班人背景色
            bg_holidayd = PatternFill("solid", fgColor="f15b6c")  # 节假日
            bg_holidayd_work = PatternFill("solid", fgColor="87843b")  # 节假日调休

            ws.merge_cells('A1:H1')  # 合并单元格
            ws['A1'] = "值 班 表"
            ws["A1"].fill = PatternFill("solid", fgColor="76becc")
            ws["A1"].border = border
            ws["A1"].alignment = Alignment(horizontal='center', vertical='center')
            ws["A1"].font = Font(u'宋体', size=20, bold=True, italic=False, strike=False, color='000000')

            for col_idx in range(len(col_name)):
                col = col_name[col_idx]
                cell_pos = col + "2"
                ws[cell_pos] = first_row[col_idx]
                ws[cell_pos].alignment = Alignment(horizontal='center', vertical='center')
                ws[cell_pos].fill = pattern_fill
                ws[cell_pos].border = border
                if col_idx == 0:
                    ws.column_dimensions[col].width = 25
                else:
                    ws.column_dimensions[col].width = 15

            row_number = math.ceil(self.weeks) + 2  # +2 ：第一行为【值班表】第二行为【星期】
            # 找到周一的日期，和对应的人
            start_time = datetime.datetime.strptime(self.start_dt, '%Y年%m月%d日')
            delta = start_time.weekday()
            start_time = start_time - datetime.timedelta(days=delta)
            # 找到周一值班的人
            staff_idx = len(self.staff) - delta

            holidays = self.get_holidays(start_time, self.weeks)

            for row_idx in range(3, int(row_number) + 2):  # 从第三行开始排班
                ws.row_dimensions[row_idx].height = 30  # 行高
                date_desc = (start_time + datetime.timedelta(weeks=row_idx - 3)).strftime('%Y{y}%m{m}%d{d}').format(
                    y='年', m='月', d='日') + "-" + (start_time + datetime.timedelta(weeks=row_idx - 3, days=6)).strftime('%m{m}%d{d}').format(
                   m='月', d='日')
                for col_idx in range(len(col_name)):
                    cell_pos = col_name[col_idx] + str(row_idx)
                    bg_cell = None
                    if col_idx == 0:
                        ws[cell_pos] = date_desc
                        ws[cell_pos].fill = pattern_fill
                    else:
                        ws[cell_pos] = self.staff[staff_idx]
                        staff_idx = (staff_idx + 1) % len(self.staff)
                        date_item = (start_time + datetime.timedelta(weeks=row_idx - 3, days=col_idx - 1)).strftime(
                            '%Y{y}%m{m}%d{d}').format(y='-', m='-', d='')
                        comment = date_item
                        if col_idx in (6, 7):
                            bg_cell = bg_weekend
                        if holidays is not None and holidays.__contains__(date_item) and holidays[date_item] is not None and holidays[date_item]["holiday"] is not None:
                            if holidays[date_item]["holiday"]:
                                bg_cell = bg_holidayd
                            else:
                                bg_cell = bg_holidayd_work
                            comment = holidays[date_item]["name"] + "\n" + holidays[date_item]["date"]

                        ws[cell_pos].comment = Comment(comment, 'zlf')
                    if bg_cell is not None:
                        ws[cell_pos].fill = bg_cell
                    ws[cell_pos].alignment = Alignment(horizontal='center', vertical='center')
                    ws[cell_pos].border = border
            file = '值班表' + self.start_dt + '.xlsx'
            wb.save(file)
            self.update_staff_file()
            self.sig_complete.emit('已在当前目录下导出值班表 ' + file)
        except Exception as e:
            print(e)
            print(e.__traceback__.tb_lineno)

            self.sig_complete.emit('异常：' + str(e))

    def get_holidays(self, start_time, weeks):
        try:
            items = start_time.strftime('%Y-%m-%d')

            for i in range(int(weeks * 7) + 1):
                items = items + "," + (start_time + datetime.timedelta(days=i)).strftime('%Y{y}%m{m}%d{d}').format(y='-', m='-', d='')
            print(items)

            url = "http://timor.tech/api/holiday/batch?d=" + items
            resp = request.urlopen(url, data=None, timeout=5)
            data = json.loads(resp.read())
            if data["code"] == 0:
                print(data["holiday"])
                return data["holiday"]
        except Exception as e:
            print(e)
            print(e.__traceback__.tb_lineno)

    def update_staff_file(self):
        try:
            staff_file = open('staff.txt', 'w', encoding='utf-8')
            staff_file.write(reduce(lambda x, y: x + ' ' + y, self.staff))
            staff_file.close()
        except Exception as e:
            print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()

    icon = QIcon()
    icon.addPixmap(QPixmap('icon.ico'), QIcon.Normal, QIcon.Off)
    window.setWindowIcon(icon)

    window.show()
    sys.exit(app.exec_())
