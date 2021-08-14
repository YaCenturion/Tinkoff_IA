import sqlite3

conn = sqlite3.connect("tinkoffdata.db")
# DateTime 1970-01-01 00:00:00 UTC
creat_sql = "CREATE TABLE operations (" \
      "id INTEGER PRIMARY KEY , NumDeals NUMERIC, NumCommand NUMERIC, DateTime INTEGER," \
      "StockName TEXT, DealType TEXT, FullName TEXT, ticker TEXT," \
      "price REAL, currency TEXT, quantity INTEGER, amount REAL, brokerage REAL" \
      ")"

cursor = conn.cursor()
cursor.execute(creat_sql)

conn.close()
