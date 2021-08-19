# -*- coding: utf-8 -*-

from datetime import datetime
import sqlite3
import exporter as out


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
            # print(value)
    else:
        print('')
        print('.......................................................')
        print('---- Больше сделок не нашел.')
    print('---- Все полученные данные из БД обработаны.')
    out.save2xlsx(ArrayFull)


def ticker_update(ticker):
    if ticker in ticker_base:
        new_ticker = ticker_base[ticker]
    else:
        new_ticker = ticker
    return new_ticker


def delta(start, stop):
    start = datetime.fromisoformat(str(start))
    stop = datetime.fromisoformat(str(stop))
    return int((stop - start).days)  # Возвращает время в днях


def collect2dict(value):
    # print(value)
    deal_num = value[1]
    deal_dt = datetime.strptime(value[3], "%d.%m.%Y %H:%M:%S")
    if value[5] in ['Покупка']:
        price_type = -1
        qty_type = 1
        margin = False
    elif value[5] in ['Продажа']:
        price_type = 1
        qty_type = -1
        margin = False
    elif value[5] in ['РЕПО 2 Покупка', 'РЕПО 1 Покупка']:
        price_type = -1
        qty_type = 1
        margin = True
    elif value[5] in ['РЕПО 1 Продажа', 'РЕПО 2 Продажа']:
        price_type = 1
        qty_type = -1
        margin = True
    else:
        price_type = None
        qty_type = None
        margin = False
        print('-------- !!! Обнаружен неизвестный тип сделки', value[5])
    name = value[6]
    ticker = '$' + ticker_update(value[7])
    price = float(value[8]) * price_type
    # currency = value[9]  # Валюта сделки отрицательное значение
    qty = int(value[10])
    brokerage = float(value[12]) * -1
    orig_deal_amount = float(value[11]) * price_type - brokerage
    amount_clear = round(price * qty, 2) - brokerage  # Вычисляем чистую стоимость
    if amount_clear != orig_deal_amount:
        print(f'=====//////// Подозрение на ошибку в отчете по сделке #{deal_num} (запись: {value[0]}) :',
              amount_clear, ' != ', orig_deal_amount)

    # Добавление значений сделки в словарь
    if ticker not in ArrayFull:  # Если ТИКЕР проходит впервые:
        ArrayFull[ticker] = {}
        ArrayFull[ticker]['Amount'] = {}
        ArrayFull[ticker]['Qty'] = {}
        ArrayFull[ticker]['Time'] = {}
        ArrayFull[ticker]['OnHands'] = {}
        ArrayFull[ticker]['Name'] = name
        ArrayFull[ticker]['OnHands']['Amount'] = amount_clear
        ArrayFull[ticker]['Amount']['Summary'] = amount_clear  # Записываем очищенную сумму из сделки
        ArrayFull[ticker]['Qty']['Summary'] = qty * qty_type  # Записываем количесвенный баланс акций из сделки
        if margin is True:
            ArrayFull[ticker]['MarginTimes'] = 1

        ArrayFull[ticker]['OutPosition'] = 0

        # Учет времени
        ArrayFull[ticker]['Time']['StartTimer'] = deal_dt
        ArrayFull[ticker]['Time']['Delta'] = 0

        # Учет количества акций
        if qty > 0:
            ArrayFull[ticker]['Qty']['PositiveSummary'] = qty  # Записываем количество купленных акций
            ArrayFull[ticker]['Qty']['NegativeSummary'] = 0  # Для акций, которые 1-й раз НЕ шортились
        elif qty < 0:
            ArrayFull[ticker]['Qty']['PositiveSummary'] = 0  # Для акций, которые сразу в шорт
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

        # Учет текущих акций на руках
        if ArrayFull[ticker]['Qty']['Summary'] == 0:
            # Учет времени
            if ArrayFull[ticker]['Time']['StartTimer'] is not None:
                ArrayFull[ticker]['Time']['Delta'] += delta(ArrayFull[ticker]['Time']['StartTimer'], deal_dt)
                ArrayFull[ticker]['Time']['StartTimer'] = None
            else:
                ArrayFull[ticker]['StartTimer'] = deal_dt
            ArrayFull[ticker]['OutPosition'] += 1
            ArrayFull[ticker]['OnHands']['Amount'] = 0
        else:
            ArrayFull[ticker]['OnHands']['Amount'] += amount_clear

        # Учет суммы сделки
        if amount_clear < 0:  # Если покупка аций (-)
            ArrayFull[ticker]['Amount']['SummaryBay'] += amount_clear
        else:  # Если продажа аций (+)
            ArrayFull[ticker]['Amount']['SummarySell'] += amount_clear
    return ArrayFull


if __name__ == '__main__':
    get_from_db()
