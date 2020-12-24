from openpyxl import load_workbook,Workbook
from datetime import date
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

try:
    arquivo = load_workbook("devedores.xlsx")    
except:
    arquivo = Workbook()
    
linha = "----------------------------------------------------"
planilha1 = arquivo.active


def sobre():
    messagebox.showinfo(title="Debt Book",message="Programa gratuito desevolvido para distribuição com foco no pequeno comerciante, com este programa você será capaz de manusear dados sem ter nenhum conhecimento de planilhas.\nPara maiores informações ou sugestões: Rafaelmafra@live.com")


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
        else:
            print("Informe um valor correto 1,2,3 ou 4")
            menu()
def cadastrar():
    clearFrame(quadro2)
    vdev=StringVar()
    lb1["text"] = "Cadastrar Cliente"
    dev = Entry(quadro,textvariable=vdev)
    dev.place(x=10,y=25,width=270,height=20)
    botao=Button(quadro, text="Cadastrar",command=lambda:cadastro(vdev.get()))
    botao.place(x=100,y=50,width=100,height=20)

def cadastro(user):
    devedores = ()
    dataatual = date.today()
    data = dataatual.strftime('%d/%m/%Y')
    try: 
        consultados = verifica_registro()
        devedor = str(user).lower()
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

def excluir():
    vdev=StringVar()
    lb1["text"] = "Deletar Cliente"
    dev = Entry(quadro,textvariable=vdev)
    dev.place(x=10,y=25,width=270,height=20)
    botao=Button(quadro, text="Deletar",command=lambda:excluiusuario(vdev.get()))
    botao.place(x=100,y=50,width=100,height=20)

def excluiusuario(user): 
    try: 
        consultados = verifica_registro()
        devedor = str(user).lower()
        devedor = devedor.capitalize()
        global pos 
        pos = consultados.index(devedor)
        lb2['text']= "Deseja realmente remover {}\nda lista? Você perdera os valores salvos".format(devedor)
        opcao1=Button(quadro3, text='Sim', command=opsim)
        opcao1.place(x=100, y=10)
        opcao2=Button(quadro3, text="Não", command=opnao)
        opcao2.place(x=150,y=10)
    except:
        lb2['text']='Ocorreu um erro, tente novamente!'
        

def clearFrame(frame):
    # destroy all widgets from frame
    for widget in frame.winfo_children():
       widget.destroy()
       

def opsim():      
    planilha1.delete_rows(pos+1)
    arquivo.save('devedores.xlsx')
    clearFrame(quadro3)
    lb2['text'] = "Cliente Excluído da base de dados!!"

def opnao():
    lb2['text']='Operação Cancelada!'
    clearFrame(quadro3)
def sair():
    arquivo.save('devedores.xlsx')
    exit()


app = Tk()

app.title(".::.Debt Book.::. Caderneta de Dívidas")
app.geometry("510x500+300+100")
app.configure(background="#5f9ea0")
quadro = Frame(app,borderwidth="2",bg="#5f9ea0",relief="groove")
quadro.place(x=215,y=100,width=290,height=350)
quadro2 = Frame(quadro,bg="#5f9ea0")
quadro2.place(x=0,y=78,width=285,height=265)
quadro3 = Frame(quadro2,bg="#5f9ea0")
quadro3.place(x=0, y= 50,width=285,height=280)
txt1 = Label(app,text=".::.Bem Vindo Ao Debt Book.::.",bg="#fff", fg="#000",font=('arial black',12))
txt1.place(x=5, y=5, width=500,height=30)
txt2 = Label(app, text='Desenvolvido por Rafael Mafra, programa gratuito, proibido a venda',bg="#5f9ea0",fg="#fff")
txt2.place(x=5, y=450, width=500, height=30)

barrademenus=Menu(app)
app.config(menu=barrademenus)
menusobre=Menu(barrademenus, tearoff=0)
menusobre.add_command(label="Info",command=sobre)
menusobre.add_command(label="Sair",command=app.quit)
barrademenus.add_cascade(label="Sobre",menu=menusobre)
Button(app,text="[1] - Cadastrar Devedor/Dívida ", command=cadastrar).place(x=10,y=100,width=200,height=40)
Button(app,text="[2] - Listar Devedores", command=leitura).place(x=10,y=150,width=200,height=40)
Button(app,text="[3] - Exibir Relatório detalhado", command=relatorio).place(x=10,y=200,width=200,height=40)
Button(app,text="[4] - Deletar Devedor", command=excluir).place(x=10,y=250,width=200,height=40)
Button(app,text="[5] - Sair do Programa", command=sair).place(x=10,y=300,width=200,height=40)



lb1 = Label(quadro, text="Bem Vindo!!", bg="#5f9ea0", font=("Arial",14))
lb1.pack()
lb2 = Label(quadro2, text="", bg="#5f9ea0",font=("arial",12))
lb2.place(x=1,y=5)



app.mainloop()