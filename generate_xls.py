__author__ = 'lepik'
import xlwt3 as xlwt
import xlwt_styles
import os
import calc_week
import db_connect
import sql_queries
xlLeft, xlRight, xlCenter = -4131, -4152, -4108

#print((sql_queries.groups_query()))


def generate_empty_tables(n_week=calc_week.get_n_week(), week=calc_week.get_week(), stud_day_num=5):
    print(len(sql_queries.groups_query()))
    for sql_line in (sql_queries.groups_query()):
        group_name = sql_line['group_name']
        course = sql_line['course']
        spec_name = sql_line['spec_name']

        wb = xlwt.Workbook(encoding='cp1251',style_compression=0)
        ws = wb.add_sheet("Тиждень №%i" % n_week, cell_overwrite_ok=False)


        ### Формирование заголовков таблицы
        #По вертикали (первый столбец)
        ws.write_merge(0, 0, 0, 1, "Тиждень №%i (%s)" % (n_week, week), xlwt_styles.style) # Границы мерджа: верх, низ, лево, право
        ws.write_merge(1, 1, 0, 1, "Група %s" % group_name, xlwt_styles.style)
        ws.write_merge(2, 2, 0, 1, 'Навчальна дисципліна, викладач, № заняття(теми), назва теми скорочена',xlwt_styles.style)
        ws.row(2).height_mismatch = True
        ws.row(2).height = 200*20
        ws.write_merge(3, 3, 0, 1, 'Аудиторія',xlwt_styles.style)
        ws.write(4, 0, '№',xlwt_styles.style)
        ws.col(0).width_mismatch = True
        ws.col(0).width = 50*20
        ws.write(4, 1, 'ПІБ студента',xlwt_styles.style)
        ws.col(1).width_mismatch = True
        ws.col(1).width = 500*20          # stud fio cell length
        ws.write_merge(4, 4, 2, 71, '',xlwt_styles.style)

        #По горизонтали (кроме первого столбца)
        i1 = 2
        i3 = 2
        for day in ['Понеділок','Вівторок','Середа','Четвер','П`ятниця','Субота','Неділя']:
            ws.write_merge(0, 0, i1, i1+9, 'число.месяц.год',xlwt_styles.style)
            ws.write_merge(1, 1, i1, i1+9, day ,xlwt_styles.style)
            for n in range(5) :
                ws.write_merge(2, 2, i3, i3+1, 'Предмет №%i (ПIБ преп.)'%(n+1), xlwt_styles.style_v)
                ws.write_merge(3, 3, i3, i3+1, '',xlwt_styles.style)
                i3 += 2
            i1 += 10
        ws.write_merge(0, 0, 72, 75, 'Пропущено занять (годин)', xlwt_styles.style)
        ws.write_merge(1, 1, 72, 73, 'За тиждень',xlwt_styles.style)
        ws.write_merge(1, 1, 74, 75, 'Від початку семестру', xlwt_styles.style)
        ws.write_merge(2, 4, 72, 72, 'Усього',xlwt_styles.style_v)
        ws.write_merge(2, 4, 73, 73, 'З повіжної причини', xlwt_styles.style_v)
        ws.write_merge(2, 4, 74, 74, 'Усього',xlwt_styles.style_v)
        ws.write_merge(2, 4, 75, 75, 'З повіжної причини', xlwt_styles.style_v)
        ws.row(1).height_mismatch = True
        ws.row(1).height = 40*20
        for i in range(2,72):
            ws.col(i).width_mismatch = True
            ws.col(i).width = 30*20
        for i in range(72,76):
            ws.col(i).width_mismatch = True
            ws.col(i).width = 60*20
        n = 1
        k = 5

        con = db_connect.con
        with con:
            cur_s = con.cursor()
            cur_s.execute("SELECT fio FROM students WHERE group_name='%s'" % group_name)
            #cols = [cn[0] for cn in cur_s.description]
            rows_s = cur_s.fetchall()
            for row_s in rows_s:
                print (row_s[0])
                ws.write(k, 0, n ,xlwt_styles.style)
                ws.write(k, 1,(row_s[0]), xlwt_styles.style)
                k += 1
                n += 1
                en_n = k
            print ("Кількість студентів = ", en_n)  # enabled cells borders

            # Добавление формулы подсчета пропущенных часов
            # (Умолчания: "н" - без уважительной причины, "у" - по уважительной причине)
            for i in range(5,en_n):
                ws.write(i, 72, xlwt.Formula('COUNTIF(C%i:BT%i;"%s")+'
                                             'COUNTIF(C%i:BT%i;"%s")' % ((i+1), (i+1), 'н', (i+1), (i+1), 'у')), xlwt_styles.style)
                ws.write(i, 73, xlwt.Formula('COUNTIF(C%i:BT%i;"%s")' % ((i+1), (i+1), 'у')),xlwt_styles.style)

            # Разрешение редактирования полей (для en_n студентов на все учебные дни недели (5, 6, или 7))
            for i in range(5,en_n):
                if stud_day_num == 5:
                    for j in range(2,52):
                        ws.write(i, j, '', xlwt_styles.np_style)
                elif stud_day_num == 6:
                    for j in range(2,62):
                        ws.write(i, j, '', xlwt_styles.np_style)
                elif stud_day_num == 7:
                    for j in range(2,72):
                        ws.write(i, j, '', xlwt_styles.np_style)

            ws.protect = True  #Excell sheet protection
            ws.password = "pass" #Excell sheet password

            wb.save('%s/XLS_Generated/%s_%s_%s.xls'%(os.getcwd(),group_name, course, spec_name))
            attachment = ('%s_%s_%s.xls'%(group_name, course, spec_name))
        #return attachment

generate_empty_tables()


