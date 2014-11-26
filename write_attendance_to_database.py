import xlrd
import os

import sql_queries


def read_xls(filename, attendance=[]):
    l = []

    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_index(0)
    num_rows = worksheet.nrows - 1
    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        l.append(row)
        # print (row)

    week = str(l[0][0])[6:-1]
    group = str(l[1][0])[12:-1]

    for i in range(5, len(l)):
        student_att = ({
                       '№': int(str(l[i][0])[7:-2]),
                       'ФИО': str(l[i][1])[6:-1],
                       'группа': group,
                       'неделя': week,
                       'посещаемость': '' + str([str(k)[-2] for k in l[i][2:-4]]),
                       'пропусков всего': (str(l[i][-4])[7:-2]),
                       'по уважительной': (str(l[i][-3])[7:-2]),
                       })
        attendance.append(student_att)
    #print(attendance)
    return attendance


a = []
i = 0
for fn in next(os.walk('./XLS_Received_from_elders/'))[2]:
    print(len(read_xls(('%s/XLS_Received_from_elders/%s' % (os.getcwd(), fn)), a)))

#print((a[i]['ФИО']))
sql_queries.attendance_query(a)

