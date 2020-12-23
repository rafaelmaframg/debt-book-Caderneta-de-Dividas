from openpyxl import load_workbook,Workbook
from datetime import date

try:
    arquivo = load_workbook("devedores.xlsx")    
except:
    arquivo = Workbook()
    
linha = "----------------------------------------------------"
planilha1 = arquivo.active

def menu():
    while True:  
        print('\nSelecione a opção desejada:')
        print("[1] - Cadastrar Devedor")
        print("[2] - Listar Devedores")
        print("[3] - Exibir Relatório de pagamentos")
        print('[4] - Deletar dados cliente')
        print("[5] - Sair do Programa")
    
        try:
            op = int(input())
        except:
            print("Opção Invalida")
            menu()
        print(linha)
        if op == 1:
            cadastro()
        elif op == 2:
            leitura()
        elif op == 3:
            relatorio()
        elif op == 4:
            excluiusuario()
            arquivo.save('devedores.xlsx')
        elif op == 5:
            arquivo.save('devedores.xlsx')
            exit()
        
        
        else:
            print("Informe um valor correto 1,2,3 ou 4")
            menu()

def cadastro():
    devedores = ()
    dataatual = date.today()
    data = dataatual.strftime('%d/%m/%Y')
    try: 
        consultados = verifica_registro()
        devedor = str(input("Insira o nome do devedor a ser cadastrado: ")).lower()
        devedor = devedor.capitalize()
        
        if devedor not in consultados:
            divida = str(input('Digite o Valor da Divida: R$'))
            divida = divida.replace(",",".")
            divida = float(divida)
            if divida < 0:
                print(linha)
                print("Insira apenas valores positivos\nOperação cancelada!")
                print(linha)
                menu()
            devedores = divida,devedor,data,divida
            planilha1.append(devedores)
            arquivo.save('devedores.xlsx')
        else:
            max_coluna = planilha1.max_column
            print(linha)
            print("Usuario já cadastrado")
            print(linha)
            print("Informe o tipo de operação financeira:")
            fin = input("[1] Somar a Divida \n[2] Abater Divida:\n")
            if fin == "1":
                somar = str(input("Digite o valor da divida a SOMAR [+]:  "))
                somar = somar.replace(",",".")
                somar = float(somar)  
                if somar < 0:
                    print(linha)
                    print("Insira apenas valores positivos\nOperação cancelada!")
                    print(linha)
                    menu()
                pos = consultados.index(devedor)
                valor = float(planilha1['A'+str(pos+1)].value)+somar
                planilha1['A'+str(pos+1)] = str(valor)
                for i in range (2, max_coluna+2):
                    if planilha1.cell(row=pos+1, column=i).value == None:
                        planilha1.cell(row=pos+1, column=i).value = str(data)
                        planilha1.cell(row=pos+1, column=i+1).value = str(somar)
                print(linha)
                print("Novo Valor de R${} atualizado para {}".format(planilha1['A'+str(pos+1)].value, devedor))
            elif fin == "2":
                abate = str(input("Digite o valor a ser ABATIDO [-]:  "))
                abate = abate.replace(",",".")
                abate = float(abate)
                if abate < 0:
                    print(linha)
                    print("Insira apenas valores positivos\nOperação cancelada!")
                    print(linha)
                    menu()
                pos = consultados.index(devedor)
                valor = float(planilha1['A'+str(pos+1)].value)-abate
                planilha1['A'+str(pos+1)] = str(valor)
                for i in range(2, max_coluna+2):
                    if planilha1.cell(row=pos+1, column=i).value == None:
                        planilha1.cell(row=pos+1, column=i).value = str(data)
                        planilha1.cell(row=pos+1, column=i+1).value = "-"+str(abate)
                print(linha)
                print("Novo Valor de R${} atualizado para {}".format(planilha1['A'+str(pos+1)].value, devedor))
            else:
                print("Digite um valor valido, 1 ou 2")
                menu()
        arquivo.save('devedores.xlsx')     
        print(linha)
        
    except:
        print(linha)
        print("Formato de entrada invalido")
        print(linha)
        menu()

def leitura():
    max_linha= planilha1.max_row
    max_coluna= planilha1.max_column
    for i in range (1, max_linha+1):
        if planilha1.cell(row=i, column=1).value == "0.0" or planilha1.cell(row=i, column=1).value == None:
            continue
          
        print(planilha1.cell(row=i, column=2).value, end="\n")
        print("data:", end=' ')
        for j in range(2, max_coluna+2):
            if planilha1.cell(row=i, column=j).value == None and planilha1.cell(row=i, column=j-1).value != None:
                print((planilha1.cell(row=i, column=j-2).value), end=" ")
        print("R$ {:.2f}\n".format(float(planilha1.cell(row=i, column=1).value)))                  
        print(linha)

def relatorio():
    max_coluna = planilha1.max_column
    try:
        devedor = str(input('Digite o nome do cliente a ser pesquisado:')).lower()
        print(linha)
        devedor = devedor.capitalize()
        registro = verifica_registro()
        pos = registro.index(devedor)
        print("Exibindo resultados para {}".format(devedor))
        for j in range(3, max_coluna+1):
            if planilha1.cell(row=pos+1, column=j).value != None:
                if j%2==0:
                    print("R$: {:.2f}".format(float(planilha1.cell(row=pos+1, column=j).value)))
                else:
                    print(planilha1.cell(row=pos+1, column=j).value, end=" ")
        total = planilha1.cell(row=pos+1, column=1).value
        print(linha)
        print("Resultado Geral: {:.2f}".format(float(total)))   
        print(linha)
    except: 
        print("Usuario invalido")
        menu()
    
def verifica_registro():
    consultados = []
    max_linha= planilha1.max_row
    for i in range(1,max_linha+1):
        consultado = planilha1.cell(row=i, column=2).value        
        consultados.append(consultado)
    return consultados

def excluiusuario(): 
    try: 
        consultados = verifica_registro()
        devedor = str(input("Insira o nome do devedor a ser removido: ")).lower()
        devedor = devedor.capitalize()
        pos = consultados.index(devedor)
        print("Deseja realmente remover {} da lista, você perdera os valores salvos".format(devedor))
        o = int(input("[1] - Sim \n[2] - Não\n"))
    except:
        print('Ocorreu um erro')
        menu()
        
    if o == 1:
        planilha1.delete_rows(pos+1)


