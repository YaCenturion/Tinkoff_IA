# -*- coding: utf-8 -*-

# import sqlite3
from datetime import datetime
from openpyxl import Workbook


def save2xlsx(data):
    print()
    print('...............................................................')
    print(f'Начинаю писать в report_{datetime.date(datetime.today())}.xlsx')
    wb = Workbook()
    wb['Sheet'].title = 'Tinkoff Report'
    sheet = wb.active
    # Название отчета
    sheet['B1'] = 'Отчет на основе анализа данных Tinkoff Инвестиции'
    # Шапка таблицы
    sheet['B2'] = 'Название'
    sheet['C2'] = 'Тикер'
    sheet['D2'] = 'Кол-во входов'
    sheet['E2'] = 'Купленно всего'
    sheet['F2'] = 'Сейчас в портфеле'
    sheet['G2'] = 'Вложено'
    sheet['H2'] = 'Получено'
    sheet['I2'] = 'Доход'

    sheet['J2'] = 'Время владения'
    sheet['K2'] = 'Прибыль на акцию'
    sheet['L2'] = 'Эффективность %'
    sheet['M2'] = 'gIndex'
    sheet['N2'] = ''
    sheet['O2'] = 'Сейчас вложено'
    sheet['P2'] = 'Средняя цена'

    x = 3
    for ticker in data:
        paper = data[ticker]
        delta = 1 if paper['Time']['Delta'] is None else paper['Time']['Delta']
        print(ticker, ' - обработан и записан в XLSX')
        sheet.cell(row=x, column=2).value = paper['Name']
        sheet.cell(row=x, column=3).value = ticker
        sheet.cell(row=x, column=4).value = paper['OutPosition']
        sheet.cell(row=x, column=5).value = paper['Qty']['PositiveSummary']
        sheet.cell(row=x, column=6).value = paper['Qty']['Summary']
        if paper['OnHands']['Amount'] > 0:
            sheet.cell(row=x, column=7).value = paper['Amount']['SummaryBay']
            sheet.cell(row=x, column=8).value = paper['Amount']['SummarySell'] - paper['OnHands']['Amount']
            sheet.cell(row=x, column=9).value = paper['Amount']['Summary'] - paper['OnHands']['Amount']
        elif paper['OnHands']['Amount'] < 0:
            sheet.cell(row=x, column=7).value = paper['Amount']['SummaryBay'] - paper['OnHands']['Amount']
            sheet.cell(row=x, column=8).value = paper['Amount']['SummarySell']
            sheet.cell(row=x, column=9).value = paper['Amount']['Summary'] - paper['OnHands']['Amount']
        elif paper['OnHands']['Amount'] == 0:
            sheet.cell(row=x, column=7).value = paper['Amount']['SummaryBay']
            sheet.cell(row=x, column=8).value = paper['Amount']['SummarySell']
            sheet.cell(row=x, column=9).value = paper['Amount']['Summary']

        sheet.cell(row=x, column=10).value = delta
        sheet.cell(row=x, column=11).value = round(paper['Amount']['Summary'] / paper['Qty']['PositiveSummary'], 2)
        sheet.cell(row=x, column=12).value = 'в разработке'
        sheet.cell(row=x, column=13).value = 'в разработке'

        sheet.cell(row=x, column=15).value = paper['OnHands']['Amount']
        if paper['Qty']['Summary'] != 0:
            sheet.cell(row=x, column=15).value = paper['OnHands']['Amount']
            sheet.cell(row=x, column=16).value = round(paper['OnHands']['Amount'] * -1 / paper['Qty']['Summary'], 2)
        else:
            sheet.cell(row=x, column=15).value = paper['OnHands']['Amount'] * -1
            sheet.cell(row=x, column=16).value = 0
        x += 1
    wb.save(f'report_{datetime.date(datetime.today())}.xlsx')
    wb.close()
    print()
    print('......................................................')
    print('Все готово!')
