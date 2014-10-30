__author__ = 'lepik'
import xlwt3 as xlwt
import os
import calc_week
import get_db_data
xlLeft, xlRight, xlCenter = -4131, -4152, -4108


def generate_empty_tables(n_week=calc_week.get_n_week(), week=calc_week.get_week(), con=get_db_data.db_connect()['con'],
                          group_name=get_db_data.db_connect()['group_name'], course=get_db_data.db_connect()['course'],
                          spec_name=get_db_data.db_connect()['spec_name'], stud_day_num=5):
							  
    wb = xlwt.Workbook(encoding='cp1251',style_compression=0)
    ws = wb.add_sheet("Тиждень №%i" % n_week, cell_overwrite_ok=False)

    ### Описание стиля 1(horizontal)
    # перенос по словам, выравнивание
    alignment = xlwt.Alignment()
    alignment.wrap = 1
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT,
    # HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    # шрифт
    font = xlwt.Font()
    font.name = 'Arial Cyr'
    font.bold = True
    # границы
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN  # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    # Создаём стиль с нашими установками
    style = xlwt.XFStyle()
    style.font = font
    style.alignment = alignment
    style.borders = borders

    ### Описание стиля 2(vertical)
    # перенос по словам, выравнивание
    alignment_v = xlwt.Alignment()
    alignment_v.wrap = 1
    alignment_v.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment_v.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment_v.rota = 90
    # шрифт
    font_v = xlwt.Font()
    font_v.name = 'Arial Cyr'
    font_v.bold = True
    # границы
    borders_v = xlwt.Borders()
    borders_v.left = xlwt.Borders.THIN # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders_v.right = xlwt.Borders.THIN
    borders_v.top = xlwt.Borders.THIN
    borders_v.bottom = xlwt.Borders.THIN
    # Создаём стиль с нашими установками
    style_v = xlwt.XFStyle()
    style_v.font = font_v
    style_v.alignment = alignment_v
    style_v.borders = borders_v

    ### Описание стиля 3(Editable cells)
    np_style = xlwt.easyxf("protection: cell_locked false")

    ### Формирование заголовков таблицы
    #По вертикали (первый столбец)
    ws.write_merge(0, 0, 0, 1, "Тиждень №%i (%s)" % (n_week, week), style) # Границы мерджа: верх, низ, лево, право
    ws.write_merge(1, 1, 0, 1, "Група %s" % group_name, style)
    ws.write_merge(2, 2, 0, 1, 'Навчальна дисципліна, викладач, № заняття(теми), назва теми скорочена',style)
    ws.row(2).height_mismatch = True
    ws.row(2).height = 200*20
    ws.write_merge(3, 3, 0, 1, 'Аудиторія',style)
    ws.write(4, 0, '№',style)
    ws.col(0).width_mismatch = True
    ws.col(0).width = 50*20
    ws.write(4, 1, 'ПІБ студента',style)
    ws.col(1).width_mismatch = True
    ws.col(1).width = 500*20          # stud fio cell length
    ws.write_merge(4, 4, 2, 71, '',style)

    #По горизонтали (кроме первого столбца)
    i1 = 2
    i3 = 2
    for day in ['Понеділок','Вівторок','Середа','Четвер','П`ятниця','Субота','Неділя']:
        ws.write_merge(0, 0, i1, i1+9, 'число.месяц.год',style)
        ws.write_merge(1, 1, i1, i1+9, day ,style)
        for n in range(5) :
            ws.write_merge(2, 2, i3, i3+1, 'Предмет №%i (ПIБ преп.)'%(n+1), style_v)
            ws.write_merge(3, 3, i3, i3+1, '',style)
            i3 += 2
        i1 += 10
    ws.write_merge(0, 0, 72, 75, 'Пропущено занять (годин)', style)
    ws.write_merge(1, 1, 72, 73, 'За тиждень',style)
    ws.write_merge(1, 1, 74, 75, 'Від початку семестру', style)
    ws.write_merge(2, 4, 72, 72, 'Усього',style_v)
    ws.write_merge(2, 4, 73, 73, 'З повіжної причини', style_v)
    ws.write_merge(2, 4, 74, 74, 'Усього',style_v)
    ws.write_merge(2, 4, 75, 75, 'З повіжної причини', style_v)
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
    cur_s = con.cursor()
    cur_s.execute("SELECT fio FROM students WHERE group_name='%s'" % group_name)
    #cols = [cn[0] for cn in cur_s.description]
    rows_s = cur_s.fetchall()
    for row_s in rows_s:
        print (row_s[0])
        ws.write(k, 0, n ,style)
        ws.write(k, 1,(row_s[0]), style)
        k += 1
        n += 1
        en_n = k
    print ("Кількість студентів = ", en_n)  # enabled cells borders

    # Добавление формулы подсчета пропущенных часов
    # (Умолчания: "н" - без уважительной причины, "у" - по уважительной причине)
    for i in range(5,en_n):
        ws.write(i, 72, xlwt.Formula('COUNTIF(C%i:BT%i;"%s")+'
                                     'COUNTIF(C%i:BT%i;"%s")' % ((i+1), (i+1), 'н', (i+1), (i+1), 'у')), style)
        ws.write(i, 73, xlwt.Formula('COUNTIF(C%i:BT%i;"%s")' % ((i+1), (i+1), 'у')),style)

    # Разрешение редактирования полей (для en_n студентов на все учебные дни недели (5, 6, или 7))
    for i in range(5,en_n):
        if stud_day_num == 5:
            for j in range(2,52):
                ws.write(i, j, '', np_style)
        elif stud_day_num == 6:
            for j in range(2,62):
                ws.write(i, j, '', np_style)
        elif stud_day_num == 7:
            for j in range(2,72):
                ws.write(i, j, '', np_style)

    ws.protect = True  #Excell sheet protection
    ws.password = "pass" #Excell sheet password
    os.chdir('%s/XLS_Generated/'%os.getcwd())
    wb.save('%s_%s_%s.xls'%(group_name, course, spec_name))
    attachment = ('%s_%s_%s.xls'%(group_name, course, spec_name))
    return attachment


