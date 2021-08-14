# from openpyxl import load_workbook
# from openpyxl import Workbook
import openpyxl.reader.excel as openxlsx
import sqlite3


def start():
    id_counter = 1
    all_deal = []
    file_pool = ['broker-report-2020-03-01-2020-12-31.xlsx',
                 'broker-report-2021-01-01-2021-08-07.xlsx']
    for fh in file_pool:
        id_counter, all_deal = parsing_xlsx(fh, id_counter, all_deal)
        print(f'==> File {fh} added and next deal number is:', id_counter)
        print()
        print()
    print(f'Начинаю запись в базу. Всего сделок {len(all_deal)}')
    insert2sql(all_deal)
    print('All deals in DB added. Last deal number is:', id_counter - 1)


def parsing_xlsx(fh, id_counter, all_deal):
    wb = openxlsx.load_workbook(filename=fh)
    wb.active = 0
    print(wb)
    cell_counter = 1

    while str(wb.active[f'A{cell_counter}'].value).startswith('1.2 Информация о неисполненных сделках') is not True:
        print(f'==> в ячейке A{cell_counter}:', str(wb.active[f'A{cell_counter}'].value))
        if str(wb.active[f'A{cell_counter}'].value).isdigit() is True:
            id_num = id_counter
            num_deals = int(wb.active[f'A{cell_counter}'].value)
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
            all_deal.append((id_num, num_deals, num_command, dt, stock_name, deal_type, full_name,
                             ticker, price, currency, quantity, amount, brokerage))
            print(id_num, price, ticker)

            id_counter += 1
            # all_deal.clear()
        else:
            print('==> не найден номер сделки, ищу дальше')
        cell_counter += 1
    return id_counter, all_deal


def check_deal(num_deals, deal_row):
    conn = sqlite3.connect("tinkoffdata.db")
    cursor = conn.cursor()
    # check_sql = "SELECT NumDeals FROM operations"
    cursor.execute("SELECT NumDeals FROM operations")
    if cursor.fetchone():
        insert2sql(deal_row)
    else:
        print(f'Сделка с {num_deals} номером уже есть.')


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
