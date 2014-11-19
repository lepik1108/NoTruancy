__author__ = 'lepik'
import xlrd


def read_xls(filename, send_map):
    l = []
    workbook = xlrd.open_workbook(filename)
    worksheet = workbook.sheet_by_name('Тиждень №11')

    num_rows = worksheet.nrows - 1
    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        l.append(row)
        print (row)
    #print('\n', l)
    week = str(l[0][0])[6:-1]

        # send_map = ({
        #         'Тиждень №': it,
        #         'group_name': groups__group_name,
        #         'course': groups__course,
        #         'spec_name': groups__spec_name,
        #         'elder_name': groups__elder_name,
        #         'elder_mail': groups__elder_mail,
        #     })

fn = './XLS_Received_from_elders/ВБ-081м_6_8.17010201 Системи технічного захисту ' \
     'інформації, автоматизація її обробки.xls'
sm = []
read_xls(fn, sm)