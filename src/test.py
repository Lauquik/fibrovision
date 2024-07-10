from datetime import datetime

def days_from_june_first():
    start_date = datetime(2024, 6, 1)
    today_date = datetime.today()
    delta = today_date - start_date

    return delta.days

# print(days_from_june_first())


import pyfiglet

title = pyfiglet.figlet_format("MY Title")
print(title)