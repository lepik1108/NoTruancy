__author__ = 'lepik'
import db_connect


def groups_query(con=db_connect.con):
    with con:
        cur = con.cursor()
        cur.execute('SELECT * FROM groups')  # table groups
        it = 1
        groups_query__send_list = []
        rows = cur.fetchall()
        for row in rows:
            groups__group_name = (row[0])
            groups__course = (row[1])
            groups__spec_name = (row[2])
            groups__elder_name = (row[3])
            groups__elder_mail = (row[4])

            groups_query__send_map = ({
                '№': it,
                'group_name': groups__group_name,
                'course': groups__course,
                'spec_name': groups__spec_name,
                'elder_name': groups__elder_name,
                'elder_mail': groups__elder_mail,
            })
            groups_query__send_list.append(groups_query__send_map)

            it += 1
    return groups_query__send_list


def students_query(group_name, con=db_connect.con):
    with con:
        cur_s = con.cursor()
        cur_s.execute("SELECT fio FROM students WHERE group_name='%s'" % group_name)
        students_query__send_list = []
        print('----------------------------------')
        print(group_name)
        print('----------------------------------')
        rows_s = cur_s.fetchall()
        for row_s in rows_s:
            print(row_s[0])

            students_query__send_map = ({
                'student_fio': row_s[0],
            })
            students_query__send_list.append(students_query__send_map)

    print('----------------------------------')
    print('\n')
    return students_query__send_list
