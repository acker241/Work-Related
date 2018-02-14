import pickle

ScriptDir = pickle.load( open('dir.p','rb'))
EstoqueDir = ScriptDir+'vendas'

arquivo = [('ranking.xlsx', '2017-10-06')]

pickle.dump(arquivo, open(EstoqueDir+'\\vendas.p', 'wb'))

HistEst = pickle.load( open(EstoqueDir+'\\vendas.p', 'rb'))
print(HistEst)