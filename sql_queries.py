__author__ = 'lepik'
import db_connect


def groups_query(con=db_connect.connect()):
    with con:
        cur = con.cursor()
        cur.execute('SELECT * FROM groups')  # table groups
        rows = cur.fetchall()
        it = 1
        send_map = {
            '№': '',
            'group_name': '',
            'course': '',
            'spec_name': '',
            'elder_name': '',
            'elder_mail': '',
        }
        for row in rows:
            group_name = (row[0])
            course = (row[1])
            spec_name = (row[2])
            elder_name = (row[3])
            elder_mail = (row[4])
            send_map.update({
                '№': it,
                'group_name': group_name,
                'course': course,
                'spec_name': spec_name,
                'elder_name': elder_name,
                'elder_mail': elder_mail,
            })
            it += 1
            #print(send_map)
    return send_map


#groups_query()
