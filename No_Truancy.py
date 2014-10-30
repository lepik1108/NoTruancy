# -*- coding: cp1251 -*-
#import sip
import sys
import sqlite3 as lite
import xlwt3 as xlwt
#import time
import smtplib, base64
#import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
#from time import gmtime, strftime
from datetime import date
import os

# gmail account
m_user = 'gevandrova.yana' #gmail login
m_pass = 'gevandrova' #gmail password

russian = 'windows-1251'

xlLeft, xlRight, xlCenter = -4131, -4152, -4108

cur_week = date.today().strftime("%U")
first_week = date(date.today().year,9,1).strftime("%U")
print ('Номер нынешней недели = ', cur_week)
print ('Номер первой недели в этом году = ', first_week)
n_week = (int(cur_week) - int(first_week))
if n_week>18 :
    n_week = n_week - 18  
if ( not n_week%2 ):
    week = "Парний"
else :
    week = "Непарний"

#######################################################################################################################
con = lite.connect('main.db') #database file
with con:
    cur = con.cursor()               
    cur.execute(u'SELECT * FROM groups') #table groups
    col_names = [cn[0] for cn in cur.description]
    rows = cur.fetchall()
    for row in rows: 
        group_name = (row[0])
        course = (row[1])
        spec_name = (row[2])
        elder_name = (row[3])
        elder_mail = (row[4])
        print ('\nГрупа ', (group_name), (course), (spec_name),' :')
        wb = xlwt.Workbook(encoding='cp1251',style_compression=0)
        ws = wb.add_sheet("Тиждень №%i"%(n_week),cell_overwrite_ok = False)
        ### Описание стиля 1(horizontal)
        # перенос по словам, выравнивание
        alignment = xlwt.Alignment()
        alignment.wrap = 1
        alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        # шрифт
        font = xlwt.Font()
        font.name = 'Arial Cyr'
        font.bold = True
        # границы
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THIN # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
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
        ws.write_merge(0, 0, 0, 1, "Тиждень №%i (%s)"%(n_week,week),style) # Границы мерджа: верх, низ, лево, право
        ws.write_merge(1, 1, 0, 1, "Група %s"%(group_name),style)
        ws.write_merge(2, 2, 0, 1, 'Навчальна дисципліна, викладач, № заняття(теми), назва теми скорочена',style)
        ws.row(2).height_mismatch = True
        ws.row(2).height = 200*20
        ws.write_merge(3, 3, 0, 1, 'Аудиторія',style)
        ws.write(4, 0, '№',style)
        ws.col(0).width_mismatch = True
        ws.col(0).width = 50*20
        ws.write(4, 1, 'ПІБ студента',style)
        ws.col(1).width_mismatch = True
        ws.col(1).width = 500*20          # stud fio cell lenght
        ws.write_merge(4, 4, 2, 71, '',style)
        #По горизонтали (кроме первого столбца)
        i1=2
        i3=2
        for day in ['Понеділок','Вівторок','Середа','Четвер','П`ятниця','Субота','Неділя']:
            ws.write_merge(0, 0, i1, i1+9, 'число.месяц.год',style)
            ws.write_merge(1, 1, i1, i1+9, day ,style)
            for n in range(5) :
                ws.write_merge(2, 2, i3, i3+1, 'Предмет №%i (ПIБ преп.)'%(n+1) ,style_v)
                ws.write_merge(3, 3, i3, i3+1, '',style)
                i3+=2
            i1+=10
        ws.write_merge(0, 0, 72, 75, 'Пропущено занять (годин)',style)
        ws.write_merge(1, 1, 72, 73, 'За тиждень',style)
        ws.write_merge(1, 1, 74, 75, 'Від початку семестру',style)
        ws.write_merge(2, 4, 72, 72, 'Усього',style_v)
        ws.write_merge(2, 4, 73, 73, 'З повіжної причини',style_v)
        ws.write_merge(2, 4, 74, 74, 'Усього',style_v)
        ws.write_merge(2, 4, 75, 75, 'З повіжної причини',style_v)
        ws.row(1).height_mismatch = True
        ws.row(1).height = 40*20
        for i in range(2,72):
            ws.col(i).width_mismatch = True
            ws.col(i).width = 30*20
        for i in range(72,76):
            ws.col(i).width_mismatch = True
            ws.col(i).width = 60*20
        n=1
        k=5
        cur_s = con.cursor()               
        cur_s.execute(u"SELECT fio FROM students WHERE group_name='%s'"%(group_name))
        cols = [cn[0] for cn in cur_s.description]
        rows_s = cur_s.fetchall()
        for row_s in rows_s:
            print (row_s[0])
            ws.write(k, 0, n ,style)
            ws.write(k, 1,(row_s[0]) ,style)
            k+=1
            n+=1
            en_n=k
        print ("Кількість студентів = ", en_n)  #enabled cells borders
        for i in range(5,en_n):
            n=('н')
            y=('у')
            ws.write(i, 72, xlwt.Formula('COUNTIF(C%i:BT%i;"%s")+COUNTIF(C%i:BT%i;"%s")'%((i+1),(i+1),n, (i+1),(i+1),y)),style)
            ws.write(i, 73, xlwt.Formula('COUNTIF(C%i:BT%i;"%s")'%((i+1),(i+1),y)),style)
        #Разрешение редактирования полей (пока что для en_n студентов на все дни недели)
        for i in range(5,en_n):
            for j in range(2,52):
                ws.write(i, j, '',np_style)

        ws.protect = True  #Excell sheet protection  
        ws.password = "pass" #Excell sheet password   
        os.chdir('%s/XLS_Generated/'%os.getcwd())
        wb.save('%s_%s_%s.xls'%(group_name,course,spec_name))
        attachment = ('%s_%s_%s.xls'%(group_name,course,spec_name))

        fromaddr = 'ОНПУ <%s@gmail.com>'%(m_user)
        toaddr = 'Старостам групп <%s>'%(elder_mail)
        msg = MIMEMultipart()
        msg['Subject'] = 'Журнал посещаемости'
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg_txt =  MIMEText('''Здрвствуйте, ув.%s, староста группы %s. \n\nЗаполните приложенный файл и пришлите обратно не позднее %s.%s.%s.\n\nС уважением, ОНПУ.
        '''%(elder_name, group_name, date.today().year, date.today().month, (date.today().day+7)), 'plain', 'cp1251')
        msg.preamble = ' '
        attach = MIMEApplication(open(attachment, 'rb').read())
        attach.add_header('Content-Disposition', 'attachment', filename=attachment)
        msg.attach(attach)
        msg.attach(msg_txt)
        server = smtplib.SMTP('smtp.gmail.com','587')
        server.set_debuglevel(1);
        server.starttls()
        server.login(m_user,m_pass)
        server.sendmail(fromaddr, toaddr, msg.as_string())
        server.quit()
        os.chdir('..')

con.close()
####################################################################################################################### 

sys.exit()
exit()
               
                
        
