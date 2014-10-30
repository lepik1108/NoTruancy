__author__ = 'lepik'

from datetime import date

first_week = input('Enter date which belongs to first educational '
                   'week in semester (default: September 1st of current year):')
if first_week == '':
    first_week = date(date.today().year,9,1).strftime("%U")
week_in_semester = input('Enter number of educational weeks in semester (default: 18):')
if week_in_semester == '':
    week_in_semester = 18

cur_week = date.today().strftime("%U")

print('Номер первой недели в этом году = ', first_week)
print('Номер нынешней недели = ', cur_week)
print('Колличество недель в семестре = ', cur_week)

n_week = (int(cur_week) - int(first_week))

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