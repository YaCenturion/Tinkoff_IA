# -*- coding: utf-8 -*-

import openpyxl.reader.excel as openxlsx
import sqlite3


def start():
    id_counter = 1
    db_count = 0
    file_pool = ['broker-report-2020-03-01-2020-12-31.xlsx',
                 'broker-report-2021-01-01-2021-08-07.xlsx',
                 'broker-report-2021-08-01-2021-08-18.xlsx']
    for fh in file_pool:
        id_counter, db_count = parsing_xlsx(fh, id_counter, db_count)
        print(f'==> Сделки из {fh} внесены в базу данных (только новые).')
        print()
        print()
    print('Все данные занесены в БД. Всего обработано:', db_count)


def parsing_xlsx(fh, id_counter, db_count):
    wb = openxlsx.load_workbook(filename=fh)
    wb.active = 0
    cell_counter = 1

    while str(wb.active[f'A{cell_counter}'].value).startswith('1.2 Информация о неисполненных сделках') is not True:
        print(f'==> в ячейке A{cell_counter}:', str(wb.active[f'A{cell_counter}'].value))
        deal_row = []
        if str(wb.active[f'A{cell_counter}'].value).isdigit() is True:
            id_num = id_counter
            num_deal = int(wb.active[f'A{cell_counter}'].value)
            num_command = int(wb.active[f'E{cell_counter}'].value)
            _date = str(wb.active[f'H{cell_counter}'].value)
            _time = str(wb.active[f'K{cell_counter}'].value)
            dt = str(_date + ' ' + _time)
            stock_name = str(wb.active[f'Q{cell_counter}'].value)
            deal_type = str(wb.active[f'AA{cell_counter}'].value)
            full_name = str(wb.active[f'AE{cell_counter}'].value)
            ticker = str(wb.active[f'AM{cell_counter}'].value)
            price = str(wb.active[f'AR{cell_counter}'].value).replace(',', '.')
            currency = str(wb.active[f'AV{cell_counter}'].value)
            quantity = int(wb.active[f'BA{cell_counter}'].value)
            amount = str(wb.active[f'BP{cell_counter}'].value).replace(',', '.')
            brokerage = str(wb.active[f'CA{cell_counter}'].value).replace(',', '.')

            deal_row.append((id_num, num_deal, num_command, dt, stock_name, deal_type, full_name,
                             ticker, price, currency, quantity, amount, brokerage))
            print('>> Проверяем на дублирование в базе')
            db_count = check_deal(num_deal, deal_row, db_count)
            id_counter += 1
            deal_row.clear()
        else:
            print(f'---- Данные из XLSX не найден в ячейке A{cell_counter}')
        cell_counter += 1
    return id_counter, db_count


def check_deal(num_deal, deal_row, db_count):
    conn = sqlite3.connect("tinkoffdata.db")
    cursor = conn.cursor()
    if num_deal in cursor.execute("SELECT NumDeals FROM operations"):
        print(f'---- Уже есть сделка с номером: {num_deal}')
    else:
        db_count = insert2sql(deal_row, db_count)
        print(f'>>>> Добавлена сделка {num_deal} в базу данных.')
    print('==============')
    return db_count


def insert2sql(deals, db_count):
    conn = sqlite3.connect("tinkoffdata.db")
    cursor = conn.cursor()

    # TODO Сделать проверку данных на наличие перед добавлением

    cursor.executemany("INSERT INTO operations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", deals)
    conn.commit()
    conn.close()
    db_count += 1
    return db_count


if __name__ == '__main__':
    start()
