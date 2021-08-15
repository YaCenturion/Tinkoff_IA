ArrayFull = dict()
ticker = '$MSFT'
amount_clear = 94.02
qty = 2
margin = True

ArrayFull[ticker] = {}
ArrayFull[ticker]['Amount'] = amount_clear  # Записываем очищенную сумму из сделки
ArrayFull[ticker]['AmountQty'] = qty * -1

if margin is True:
    if 'MarginTimes' not in ArrayFull[ticker]:
        ArrayFull[ticker]['MarginTimes'] = 1


print(ArrayFull)