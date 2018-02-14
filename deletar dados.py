import os, sqlite3, pickle

ScriptDir = pickle.load( open('dir.p','rb'))
EstoqueDir = ScriptDir+'estoque'

#conexão com db
DBDir = ScriptDir+'database'
conn = sqlite3.connect(DBDir+'/database.db')
db = conn.cursor()

db.execute('DELETE FROM vendas')

conn.commit()
conn.close()