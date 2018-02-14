import pickle, sqlite3, openpyxl

def abrir_estoque(arquivo):
    planilha = openpyxl.load_workbook(arquivo)
    planilha_ativa = planilha.get_sheet_by_name(planilha.get_sheet_names()[0])
    linha_max = planilha_ativa.max_row
    Cods_Descs = []
    for x in range(1,linha_max):
        col_cod = str(planilha_ativa.cell(row=x, column=4).value)
        col_desc = str(planilha_ativa.cell(row=x, column=6).value)
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
DescExc = [(('CONJUNTO','CONJ','C/BERMUDA','C/CAL','C/LEG','C/SAI','C/SHO'),'CONJUNTO'),
           (('CAMISETA','CAMISA','POLO'),'CAMISETA'),(('SHORTS', 'BERMUDA'),'SHORTS'),(('CASACO'),'CASACO'),
           (('JAQUETA'),'JAQUETA'),(('VESTIDO','VSTIDO'),'VESTIDO'),
           (('TÊNIS','TENIS','SAPATO','SAPATILHA','SAPATINHO','CHINELO','SANDALIA','SANDÁLIA','CALÇADO','CALCADO',
             'PAPETE', 'CRAVINHO', 'COMFORT'),'CALCADO'),(('MEIA MALHA','MOLECOTTON','TINTO','TECIDO'),'TECIDO'),
           (('BLUSA','BLUSÃO','BLUSINHA','BLUS'),'BLUSA'),(('BODY'),'BODY'),(('CALÇA','CALCA'),'CALÇA'),
           (('SAIA'),'SAIA'),(('REGATA', 'MACHÃO','MACHAO'),'REGATA'),(('LEGGING'),'LEGGING'),(('MEIA'),'MEIA')]

if len(DBList) == 0:
    print('Cadastro vazio, criando primeiro cadastro')

    #definição das listas utilizadas
    ListaParaCadastro = [] #lista das strings brutas para analisar
    ListaPalavras = {} #lista das palavras separadas para ver as mais comuns
    ListaFinal = [] #lista final para puxar para a DB

    #carregar arquivos de estoque para ver quais devemos puxar
    EstArquivos = pickle.load(open(ScriptDir+'estoque/est.p', 'rb'))

    for x in range(len(EstArquivos)):
        try:
            print('Abrindo',EstArquivos[x][0])
            cods_da_planilha = abrir_estoque(ScriptDir+'estoque/'+EstArquivos[x][0])
            for y in range(len(cods_da_planilha)):
                Entrada = (cods_da_planilha[y][0], cods_da_planilha[y][1],
                           str(cods_da_planilha[y][1]).replace('.',' '))
                if ListaParaCadastro.count(Entrada) == 0:
                    ListaParaCadastro.append(Entrada)
        except IOError:
            print('Arquivo não encontrado')

    TamanhoAntes = len(ListaParaCadastro)

    ref = 0
    while ref < len(ListaParaCadastro):
        if ref < len(ListaParaCadastro):
            for categorias in range(len(DescExc)):
                for palavras in range(len(DescExc[categorias][0])):
                    if isinstance(DescExc[categorias][0],str):
                        String_to_check = DescExc[categorias][0]
                    else:
                        String_to_check = DescExc[categorias][0][palavras]
                    Desc_Ref = str(ListaParaCadastro[ref][1]).upper()
                    if str(Desc_Ref).count(String_to_check) > 0:
                        Cadastro = (ListaParaCadastro[ref][0],ListaParaCadastro[ref][1],DescExc[categorias][1])
                        ListaFinal.append(Cadastro)
                        ListaParaCadastro.pop(ref)
                        ref -= 1
                        break
        if ref == 0 and TamanhoAntes > len(ListaParaCadastro):
            ref = 0
        else: ref +=1

    print(ListaParaCadastro)
    TamanhoFinal = len(ListaParaCadastro)
    print(TamanhoAntes,TamanhoFinal,sep='---')

    PalavrasEx = ['COM','DE','C/','REF','EM']
    for x in range(len(ListaParaCadastro)):
        PalAdd = list(str(ListaParaCadastro[x][2]).split())
        for y in range(len(PalAdd)):
            var_string = str(PalAdd[y]).upper()
            if PalavrasEx.count(var_string) == 0:
                try: ListaPalavras[var_string] += 1
                except KeyError: ListaPalavras[var_string] = 1
            else:
                ListaPalavras[var_string] = 0
