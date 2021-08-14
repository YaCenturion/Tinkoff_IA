# from openpyxl import load_workbook
import openpyxl.reader.excel as openxlsx
# from openpyxl import Workbook

import sqlite3


def start():
    # file_name = input('So?^:')
    file_name = 'broker-report-2020-03-01-2020-12-31.xlsx'
    parsing(file_name)


def parsing(fh):
    wb = openxlsx.load_workbook(filename=fh)
    wb.active = 0
    print(wb)
    counter = 1
    id_counter = 1
    all_deal = []
    deal = []
    while str(wb.active[f'A{counter}'].value).startswith('Руководитель') is not True:
        if str(wb.active[f'A{counter}'].value).isdigit() is True:
            id_num = deal.append(id_counter)
            num_deals = deal.append(int(wb.active[f'A{counter}'].value))
            num_command = deal.append(int(wb.active[f'E{counter}'].value))
            _date = str(wb.active[f'H{counter}'].value)
            _time = str(wb.active[f'K{counter}'].value)
            dt = deal.append(_date + ' ' + _time)
            stock_name = deal.append(str(wb.active[f'Q{counter}'].value))
            deal_type = deal.append(str(wb.active[f'AA{counter}'].value))
            full_name = deal.append(str(wb.active[f'AE{counter}'].value))
            ticker = deal.append(str(wb.active[f'AM{counter}'].value))
            price = deal.append(float(str(wb.active[f'AR{counter}'].value).replace(',', '.')))
            currency = deal.append(str(wb.active[f'AV{counter}'].value))
            quantity = deal.append(int(wb.active[f'BA{counter}'].value))
            amount = deal.append(float(str(wb.active[f'BP{counter}'].value).replace(',', '.')))
            brokerage = deal.append(float(str(wb.active[f'CA{counter}'].value).replace(',', '.')))

            # insert2sql(deal)
            id_counter += 1
            cl = '(' + str(deal)[1:-1] + ')'
            print(cl)
            all_deal.append(cl)
            print(deal)
            deal.clear()
        counter += 1
    print(all_deal)
    insert2sql(all_deal)


def insert2sql(deal_row):
    conn = sqlite3.connect("tinkoffdata.db")
    cursor = conn.cursor()
    # TODO проверка транзакции на наличие в базе
    cursor.executemany("INSERT INTO operations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", deal_row)
    conn.close()


if __name__ == '__main__':
    start()
