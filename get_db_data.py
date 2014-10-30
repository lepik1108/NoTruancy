__author__ = 'lepik'
import sqlite3 as lite


def db_connect(db_name='main.db'):
    con = lite.connect(db_name)  # database file
    print('Connected to %s'%db_name)
    with con:
        cur = con.cursor()
        cur.execute(u'SELECT * FROM groups')  # table groups
        # col_names = [cn[0] for cn in cur.description]
        rows = cur.fetchall()
        it = 0
        for row in rows:
            group_name = (row[0])
            course = (row[1])
            spec_name = (row[2])
            elder_name = (row[3])
            elder_mail = (row[4])

            #if it == len(rows):
                #print('Closing DB connection')
                #con.close()
            #else:
            print ('\nГрупа ', group_name, course, spec_name, ':')
            return {'con': con,'group_name': group_name, 'course': course,'spec_name': spec_name,
                    'elder_name': elder_name,'elder_mail': elder_mail}
            #it += 1

