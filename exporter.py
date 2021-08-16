# -*- coding: utf-8 -*-

import sqlite3
from openpyxl import Workbook


def save2xlsx(data):
    print('буду писать в эксель')
    wb = Workbook()
    wb['Sheet'].title = 'Tinkoff Report'
    sheet = wb.active
    # TODO Сюда добавить парсер данных из словаря
    row_count = 2
    col_count = 2
    sheet['A1'] = 'Отчет'
    for ticker in data:
        print(ticker)
        sheet.cell(row=2, column=2).value = ticker
        col_count += 1
        sheet.cell(row=2, column=3).value = ticker
        col_count += 1
    wb.save('report.xlsx')
    wb.close()
    pass
