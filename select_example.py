import sqlite3, pickle, datetime

ScriptDir = pickle.load( open('dir.p','rb'))

#conexão com db
DBDir = ScriptDir+'database'
conn = sqlite3.connect(DBDir+'/database.db')
db = conn.cursor()

DiaIn = '2017-10-04'
DiaFin = '2017-10-05'

# db.execute('''SELECT item, qtd, valor FROM vendas WHERE data BETWEEN ? AND ?''', (DiaIn,DiaFin))
# teste = db.fetchall()

CNPJ = '07.655.260/0005-78'
db.execute('''SELECT vendas.item, vendas.qtd, vendas.valor, est_atual.qtd
            FROM vendas
            JOIN est_atual ON vendas.item = est_atual.item
            WHERE est_atual.cnpj = ?
            AND vendas.data BETWEEN ? AND ?''', (CNPJ,DiaIn,DiaFin))

lista = db.fetchall()

for x in range(len(lista)):
    print(lista[x])

# db.execute('''SELECT DISTINCT item FROM est_atual''')
# codigos_unicos = list(db.fetchall())
#
#
# print(codigos_unicos.__class__)
# for x in range(len(codigos_unicos)):
#     print(len(codigos_unicos[x][0]), codigos_unicos[x][0], sep= ' ')
