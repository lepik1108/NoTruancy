import time
from datetime import date

first_week = input('Enter date which belongs to first educational '
                   'week in semester (default: September 1st of current year (01-09-2014)):')  # привести к формату даты
if first_week == '':
    first_week = date(date.today().year, 9, 1).strftime("%W")

else:
    first_week = date(int(first_week[6:]), int(first_week[3:5]), int(first_week[0:2])).strftime("%W")

week_in_semester = input('Enter number of educational weeks in semester (default: 16):')
if week_in_semester == '':
    week_in_semester = 16

cur_week = date.today().strftime("%W")

print('Номер первой недели в этом году = ', first_week)
print('Номер нынешней недели = ', cur_week)
#print('Колличество недель в семестре = ', cur_week)

n_week = (int(cur_week) - int(first_week))
print('Номер текушей недели = ', n_week)

if n_week > week_in_semester:
    n_week -= week_in_semester
if not n_week % 2:
    week = "Парний"
else:
    week = "Непарний"


def get_week():
    return week


def get_n_week():
    return n_week

