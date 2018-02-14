import pickle, sqlite3, openpyxl

def abrir_estoque(arquivo):
    planilha = openpyxl.load_workbook(arquivo)
    planilha_ativa = planilha.get_sheet_by_name(planilha.get_sheet_names()[0])
    linha_max = planilha_ativa.max_row
    Cods_Descs = []
    for x in range(1,linha_max):
        col_cod = str(planilha_ativa.cell(row=x, column=4).value)
        col_desc = str(planilha_ativa.cell(row=x, column=6).value)
        col_qtd = str(planilha_ativa.cell(row=x, column=10).value)
        if col_cod != 'None' and col_desc != 'None' and (col_cod,col_desc) != ('Código', 'Descrição'):
            if Cods_Descs.count((col_cod, col_desc)) == 0:
                Cods_Descs.append((col_cod, col_desc))
    return Cods_Descs

ScriptDir = pickle.load(open('dir.p','rb'))

DBDir = ScriptDir+'database'
DBConnection = sqlite3.connect(DBDir+'/database.db')
DBCursor = DBConnection.cursor()

DBCursor.execute('''SELECT cod FROM itens''')
DBList = DBCursor.fetchall()

#exceções ao cadastro - criar um dict se detectar alguma string já define categoria
DescExc = [('CONJUNTO','CONJ','C/BERMUDA','C/CAL','C/LEG','C/SAI','C/SHO')]

if len(DBList) == 0:
    print('Cadastro vazio, criando primeiro cadastro')

    #carregar arquivos de estoque para ver quais devemos puxar
    EstArquivos = pickle.load(open(ScriptDir+'estoque/est.p', 'rb'))
    ListaParaCadastro = []
    for x in range(len(EstArquivos)):
        try:
            print('Abrindo',EstArquivos[x][0])
            cods_da_planilha = abrir_estoque(ScriptDir+'estoque/'+EstArquivos[x][0])
            for y in range(len(cods_da_planilha)):
                if ListaParaCadastro.count(cods_da_planilha[y]) == 0:
                    ListaParaCadastro.append(cods_da_planilha[y])
        except IOError:
            print('Arquivo não encontrado')
    ListaPalavras = {}
    PalavrasEx = ['COM','DE','C/','REF','EM']
    for x in range(len(ListaParaCadastro)):
        PalAdd = list(str(ListaParaCadastro[x][1]).split())
        for y in range(len(PalAdd)):
            var_string = str(PalAdd[y]).upper()
            if PalavrasEx.count(var_string) == 0:
                try: ListaPalavras[var_string] += 1
                except KeyError: ListaPalavras[var_string] = 1
            else:
                ListaPalavras[var_string] = 0
    for x in range(len(ListaParaCadastro)):
        PalAdd = list(str(ListaParaCadastro[x][1]).split())
        Peso = 0
        WordCat = ''
        for y in range(len(PalAdd)):
            if ListaPalavras[str(PalAdd[y]).upper()] > Peso:
                Peso = ListaPalavras[str(PalAdd[y]).upper()]
                WordCat = str(PalAdd[y]).upper()
        ListaParaCadastro[x] = (ListaParaCadastro[x][0],ListaParaCadastro[x][1],WordCat)
    CategoriaFin = {}
    for x in range(len(ListaParaCadastro)):
        try: CategoriaFin[ListaParaCadastro[x][2]] += 1
        except KeyError: CategoriaFin[ListaParaCadastro[x][2]] = 1
    ListaChave = list(CategoriaFin.keys())
    for x in range(len(ListaChave)):
        print(ListaChave[x],CategoriaFin[ListaChave[x]],sep=';')
    for x in range(len(ListaParaCadastro)):
        print(ListaParaCadastro[x])