# from openpyxl import load_workbook
# from openpyxl import Workbook
import openpyxl.reader.excel as openxlsx
import sqlite3


def start():
    # for file_name in ('broker-report-2020-03-01-2020-12-31.xlsx', 'broker-report-2021-01-01-2021-08-07.xlsx')
    # file_name = input('So?^:')
    id_counter = 1
    all_deal = []

    fh = 'broker-report-2020-03-01-2020-12-31.xlsx'
    id_counter, all_deal = parsing_xlsx(fh, id_counter, all_deal)
    print('file added and next deal number is:', id_counter)
    print()
    print()
    fh = 'broker-report-2021-01-01-2021-08-07.xlsx'
    id_counter, all_deal = parsing_xlsx(fh, id_counter, all_deal)
    print('file added and next deal number is:', id_counter)

    print(f'Начинаю запись в базу. Всего сделок {len(all_deal)}')
    insert2sql(all_deal)
    print('base added - finish... Last deal is:', id_counter - 1)


def parsing_xlsx(fh, id_counter, all_deal):
    wb = openxlsx.load_workbook(filename=fh)
    wb.active = 0
    print(wb)
    counter = 1

    while str(wb.active[f'A{counter}'].value).startswith('1.2 Информация о неисполненных сделках') is not True:
        print(f'==> в ячейке A{counter}:', str(wb.active[f'A{counter}'].value))
        if str(wb.active[f'A{counter}'].value).isdigit() is True:
            id_num = id_counter
            num_deals = int(wb.active[f'A{counter}'].value)
            num_command = int(wb.active[f'E{counter}'].value)
            _date = str(wb.active[f'H{counter}'].value)
            _time = str(wb.active[f'K{counter}'].value)
            dt = str(_date + ' ' + _time)
            stock_name = str(wb.active[f'Q{counter}'].value)
            deal_type = str(wb.active[f'AA{counter}'].value)
            full_name = str(wb.active[f'AE{counter}'].value)
            ticker = str(wb.active[f'AM{counter}'].value)
            price = str(wb.active[f'AR{counter}'].value).replace(',', '.')
            currency = str(wb.active[f'AV{counter}'].value)
            quantity = int(wb.active[f'BA{counter}'].value)
            amount = str(wb.active[f'BP{counter}'].value).replace(',', '.')
            brokerage = str(wb.active[f'CA{counter}'].value).replace(',', '.')
            all_deal.append((id_num, num_deals, num_command, dt, stock_name, deal_type, full_name, ticker, price, currency, quantity, amount, brokerage))
            print(id_num, price, ticker)

            id_counter += 1
            # print(all_deal)
            # all_deal.clear()
        else:
            print('==> не найден номер сделки, ищу дальше')
        counter += 1
    return id_counter, all_deal


def insert2sql(deals):
    conn = sqlite3.connect("tinkoffdata.db")
    cursor = conn.cursor()
    # TODO проверка транзакции на наличие в базе
    cursor.executemany("INSERT INTO operations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", deals)
    conn.commit()
    conn.close()
    # print('added', deals)


if __name__ == '__main__':
    start()
