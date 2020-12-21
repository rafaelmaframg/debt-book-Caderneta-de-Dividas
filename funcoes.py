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
        print('[3] - Deletar dados cliente')
        print("[4] - Sair do Programa")
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
        elif op == 4:
            arquivo.save('devedores.xlsx')
            exit()
        elif op == 3:
            excluiusuario()
            arquivo.save('devedores.xlsx')
        else:
            print("Informe um valor correto 1,2,3 ou 4")
            menu()
def leitura(): 
    excluizero()
    max_linha= planilha1.max_row
    for i in range (1, max_linha+1):
        if not planilha1.cell(row=i, column=1).value == None:    
            print(planilha1.cell(row=i, column=1).value, end=" - ")
            print(planilha1.cell(row=i, column=2).value, end=" ")
            print(planilha1.cell(row=i, column=3).value, end=" ")
            print("Ultima Atualização:", end=' ')
            print(planilha1.cell(row=i, column=4).value, end="\n")           
            print(linha)

def verifica_registro():
    consultados = []
    max_linha= planilha1.max_row
    for i in range(1,max_linha+1):
        consultado = planilha1.cell(row=i, column=1).value        
        consultados.append(consultado)
    return consultados

def cadastro():
    devedores = ()
    moeda = "R$"
    dataatual = date.today()
    data = dataatual.strftime('%d/%m/%Y')
    try: 
        consultados = verifica_registro()
        devedor = str(input("Insira o nome do devedor a ser cadastrado: ")).lower()
        devedor = devedor.capitalize()
        
        if devedor not in consultados:
            divida = str(input('Digite o Valor da Divida: R$'))
            divida = divida.replace(",",".")
            devedores = devedor,moeda, divida,data
            planilha1.append(devedores)
            arquivo.save('devedores.xlsx')
        else:
            print(linha)
            print("Usuario já cadastrado")
            print(linha)
            print("Informe o tipo de operação financeira:")
            fin = input("[1] Somar a Divida \n[2] Abater Divida:\n")
            if fin == "1":
                somar = str(input("Digite o valor da divida a SOMAR [+]:  "))
                somar = somar.replace(",",".")  
                pos = consultados.index(devedor)
                valor = float(planilha1['C'+str(pos+1)].value)+float(somar)
                planilha1['C'+str(pos+1)] = str(valor)
                planilha1['D'+str(pos+1)] = str(data)
                print(linha)
                print("Novo Valor de R${} atualizado para {}".format(planilha1['C'+str(pos+1)].value, devedor))
            elif fin == "2":
                abate = str(input("Digite o valor a ser ABATIDO [-]:  "))
                abate = abate.replace(",",".")
                pos = consultados.index(devedor)
                valor = float(planilha1['C'+str(pos+1)].value)-float(abate)
                planilha1['C'+str(pos+1)] = str(valor)
                planilha1['D'+str(pos+1)] = str(data)
                print(linha)
                print("Novo Valor de R${} atualizado para {}".format(planilha1['C'+str(pos+1)].value, devedor))
            else:
                print("Digite um valor valido, 1 ou 2")
                menu()
        excluizero()
        print(linha)
        
    except:
        print(linha)
        print("Formato de entrada invalido")
        print(linha)
        menu()
    
def excluizero():
    for i in range(1, planilha1.max_row+1):
        if planilha1.cell(row=i, column=3).value == "0" or planilha1.cell(row=i, column=3).value == "0.0":
            planilha1.delete_rows(i)
    for i in range(1, planilha1.max_row+1):
        if planilha1.cell(row=i, column=1).value == None:
            planilha1.delete_rows(i)

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


