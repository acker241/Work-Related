import sqlite3
import os, pickle

#verificação e criação da database
ScriptDir = str(os.path.realpath(__file__))[0:42]
if not os.path.exists(ScriptDir+'database'):
    os.makedirs(ScriptDir+'database')
    print('Pasta Database Criada')
else:
    print('Diretório Database Existe')

if not os.path.exists(ScriptDir+'vendas'):
    os.makedirs(ScriptDir+'vendas')
    print('Pasta Vendas criada')
else:
    print('Diretório Vendas Existe')

if not os.path.exists(ScriptDir+'estoque'):
    os.makedirs(ScriptDir+'estoque')
    print('Diretório Estoque criado')
else:
    print('Diretório Estoque Existe')

pickle.dump(ScriptDir, open("dir.p", "wb"))

DBDir = ScriptDir+'database'
conn = sqlite3.connect(DBDir+'/database.db')

db = conn.cursor()

db.execute('CREATE TABLE itens (cod text, desc text, categoria text, tamanho text)')
db.execute('CREATE TABLE est_atual (item text, qtd real, data_att text, cnpj text)')
db.execute('CREATE TABLE vendas (data text, item text, qtd real, valor text, cnpj text)')

conn.commit()
conn.close()