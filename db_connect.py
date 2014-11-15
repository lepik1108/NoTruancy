__author__ = 'lepik'
import sqlite3 as lite


def connect(db_name='main.db'):
    print('Connect to %s' % db_name)
    con = lite.connect(db_name)  # database file
    return con