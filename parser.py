# -*- coding: utf-8 -*-

# import datetime as dtc
# from datetime import datetime
# import re
import sqlite3


# TODO достать из базы
# если в скольки процентах сделок прибегал к усреднению
# если усреднял то сколько раз пока не вышел из бумаги
#
ticker_base = {'GEold': 'GE', 'LUMNold': 'LUMN'}
ArrayFull = dict()


def get_from_db():
    conn = sqlite3.connect("tinkoffdata.db")
    cursor = conn.cursor()
    for value in cursor.execute("SELECT * FROM operations"):
        if value[4] == 'СПБ':  # Берем только сделки с акциями на СПБ
            collect2dict(value)
        else:
            pass
            # print(' = = = = = > > > Не обработано:')
            # print(value)
    else:
        print('---- Больше сделок не нашел.')
    print('==> Все полученные данные из БД обработаны.')
    print(ArrayFull)
    show()


def show():
    for ticker in ArrayFull:
        print(ticker, ArrayFull[ticker]['Amount']['Summary'], ArrayFull[ticker]['Qty']['Summary'])


def ticker_update(ticker):
    if ticker in ticker_base:
        new_ticker = ticker_base[ticker]
    else:
        new_ticker = ticker
    return new_ticker


def collect2dict(value):
    print(value)
    deal_num = value[1]
    # dt = value[3]  # TODO обработать дату-время
    if value[5] in ['Покупка', 'РЕПО 2 Покупка', 'РЕПО 1 Покупка']:
        price_type = -1
        qty_type = 1
        margin = False
    elif value[5] in ['Продажа', 'РЕПО 1 Продажа', 'РЕПО 2 Продажа']:
        price_type = 1
        qty_type = -1
        margin = False
    else:
        price_type = None
        qty_type = None
        margin = False
        print('-------- !!! Обнаружен неизвестный тип сделки', value[5])
    # full_name = value[6]
    ticker = '$' + ticker_update(value[7])
    price = float(value[8]) * price_type
    # currency = value[9]  # Валюта сделки отрицательное значение
    qty = int(value[10])
    brokerage = float(value[12]) * -1
    orig_deal_amount = float(value[11]) * price_type - brokerage
    amount_clear = round(price * qty, 2) - brokerage  # Вычисляем чистую стоимость
    if amount_clear != orig_deal_amount:
        print(f'======================== Подозрение на ошибку в отчете по сделке #{deal_num} / запись: {value[0]}')
        print(amount_clear, ' != ', orig_deal_amount)
    
    # Добавление значений сделки в словарь
    if ticker not in ArrayFull:  # Если ТИКЕР проходит впервые:
        ArrayFull[ticker] = {}
        ArrayFull[ticker]['Amount'] = {}
        ArrayFull[ticker]['Qty'] = {}
        ArrayFull[ticker]['Amount']['Summary'] = amount_clear  # Записываем очищенную сумму из сделки
        ArrayFull[ticker]['Qty']['Summary'] = qty * qty_type  # Записываем количесвенный баланс акций из сделки
        if margin is True:
            ArrayFull[ticker]['MarginTimes'] = 1
        # TODO Добавить учет времени
        
        # Учет количества
        if qty > 0:
            ArrayFull[ticker]['Qty']['PositiveSummary'] = qty  # Записываем количество купленных акций
            ArrayFull[ticker]['Qty']['NegativeSummary'] = 0  # Для акций, которые 1-й раз НЕ шортились
        elif qty < 0:
            ArrayFull[ticker]['Qty']['PositiveSummary'] = 0  # Для акций, которые сразу шортились
            ArrayFull[ticker]['Qty']['NegativeSummary'] = qty  # Записываем количество проданных акций
        
        # Учет суммы сделки
        if amount_clear < 0:  # Если покупка аций (-)
            ArrayFull[ticker]['Amount']['SummaryBay'] = amount_clear
            ArrayFull[ticker]['Amount']['SummarySell'] = 0
        elif amount_clear > 0:  # Если продажа аций (+)
            ArrayFull[ticker]['Amount']['SummaryBay'] = 0
            ArrayFull[ticker]['Amount']['SummarySell'] = amount_clear

    else:  # Если акция уже была и проходит вторично:
        ArrayFull[ticker]['Amount']['Summary'] += amount_clear  # Записываем очищенную сумму из сделки
        ArrayFull[ticker]['Qty']['Summary'] += qty * qty_type  # Записываем количесвенный баланс акций из сделки
        if margin is True and 'MarginTimes' in ArrayFull[ticker]:
            ArrayFull[ticker]['MarginTimes'] += 1
        elif margin is True and 'MarginTimes' not in ArrayFull[ticker]:
            ArrayFull[ticker]['MarginTimes'] = 1
        # Учет изменений в количестве для акций (вторичный)
        if qty > 0:
            ArrayFull[ticker]['Qty']['PositiveSummary'] += qty  # Добавляем количество купленных акций
        elif qty < 0:
            ArrayFull[ticker]['Qty']['NegativeSummary'] += qty  # Добавляем количество проданных акций

        # Учет суммы сделки
        if amount_clear < 0:  # Если покупка аций (-)
            ArrayFull[ticker]['Amount']['SummaryBay'] += amount_clear
        else:  # Если продажа аций (+)
            ArrayFull[ticker]['Amount']['SummarySell'] += amount_clear
    return ArrayFull


if __name__ == '__main__':
    get_from_db()
