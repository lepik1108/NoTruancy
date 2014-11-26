import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import date
import generate_xls

# gmail account
m_user = 'gevandrova.yana'  # gmail login
m_pass = 'gevandrova'  # gmail password


def send():
    for table in generate_xls.generate_empty_tables():
        group_name = table['group_name']
        course = table['course']
        spec_name = table['spec_name']
        elder_name = table['elder_name']
        elder_mail = table['elder_mail']
        attachment = table['xls_file']

        from_adr = u'ОНПУ <{0:s}@gmail.com>'.format(m_user)
        to_adr = u'Старостам групп <{0:s}>'.format(elder_mail)
        msg = MIMEMultipart()
        msg['Subject'] = 'Журнал посещаемости'
        msg['From'] = from_adr
        msg['To'] = to_adr
        msg_txt = MIMEText(u'''Здрвствуйте, ув. %s, староста группы %s.
        \n\nЗаполните приложенный файл и пришлите
        обратно не позднее %s.%s.%s.\n\nС уважением, ОНПУ.''' % (elder_name,
                                                                 group_name,
                                                                 date.today().year,
                                                                 date.today().month,
                                                                 (date.today().day + 7)), 'plain', 'utf-8')
        msg.preamble = ' '
        attach = MIMEApplication(open(attachment, 'rb').read())
        attach.add_header('Content-Disposition', 'attachment', filename=('%s_%s_%s.xls' % (group_name, course, spec_name)))
        msg.attach(attach)
        msg.attach(msg_txt)
        server = smtplib.SMTP('smtp.gmail.com', "587")
        server.set_debuglevel(1)
        server.starttls()
        server.login(m_user, m_pass)
        server.sendmail(from_adr, to_adr, msg.as_string())
        server.quit()
send()