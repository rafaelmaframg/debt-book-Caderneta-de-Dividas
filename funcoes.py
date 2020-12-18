from openpyxl import load_workbook,Workbook

try:
    arquivo = load_workbook("devedores.xlsx")    
except:
    arquivo = Workbook()
    
linha = "-----------------------------------------------"
planilha1 = arquivo.active

def menu():
    while True:  
        print('\nSelecione a opção desejada:')
        print("[1] - Cadastrar Devedor")
        print("[2] - Listar Devedores")
        print('[3] - Sair do Programa')
        op = int(input())
        print(linha)
        if op == 1:
            cadastro()
        elif op == 2:
            leitura()
        elif op == 3:
            arquivo.save('devedores.xlsx')
            exit()
        else:
            print("Informe um valor correto 1,2 ou 3")

def leitura(): 
    max_linha= planilha1.max_row
    for i in range (2, max_linha+1):
        print(planilha1.cell(row=i, column=1).value, end="-")
        print(planilha1.cell(row=i, column=2).value, end=" ")
        print(planilha1.cell(row=i, column=3).value, end="\n")           
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
    try: 
        consultados = verifica_registro()
        devedor = str(input("Insira o nome do devedor a ser cadastrado: ")).lower()
        devedor = devedor.capitalize()
        
        if devedor not in consultados:
            divida = str(input('Digite o Valor da Divida: R$'))
            divida = divida.replace(",",".")
            devedores = devedor,moeda, divida
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
                print(linha)
                print("Novo Valor de R${} atualizado para {}".format(planilha1['C'+str(pos+1)].value, devedor))
            elif fin == "2":
                abate = str(input("Digite o valor a ser ABATIDO [-]:  "))
                abate = abate.replace(",",".")
                pos = consultados.index(devedor)
                valor = float(planilha1['C'+str(pos+1)].value)-float(abate)
                planilha1['C'+str(pos+1)] = str(valor)
                print(linha)
                print("Novo Valor de R${} atualizado para {}".format(planilha1['C'+str(pos+1)].value, devedor))
            else:
                print("Digite um valor valido, 1 ou 2")
                menu()
        print(linha)
    except:
        print(linha)
        print("Formato de entrada invalido")
        print(linha)
        menu()