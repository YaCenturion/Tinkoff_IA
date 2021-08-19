from datetime import datetime, timedelta


def delta(start, stop):
    start = datetime.fromisoformat(str(start))
    stop = datetime.fromisoformat(str(stop))
    return stop - start


a = datetime.strptime('2020-03-16 10:37:57', "%Y-%m-%d %H:%M:%S")
x = datetime.now()
z = delta(a, x)


b = datetime.strptime('2021-03-16 10:37:57', "%Y-%m-%d %H:%M:%S")
y = datetime.now()
t = delta(a, x)

z += t

print(z)
