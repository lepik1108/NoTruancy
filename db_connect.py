import sqlite3 as lite

db_name = 'main.db'
print('Connect to %s' % db_name)
con = lite.connect(db_name)
