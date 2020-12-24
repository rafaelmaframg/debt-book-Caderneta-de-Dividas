from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from funcoes import *


app = Tk()

def excluir():
    vdev=StringVar()
    lb1["text"] = "Deletar Usuário"
    dev = Entry(quadro,textvariable=vdev)
    dev.place(x=10,y=25,width=250,height=20)
    botao=Button(quadro, text="Deletar",command=lambda:excluiusuario(vdev.get()))
    botao.place(x=100,y=75,width=100,height=20)
    print(botao)

def excluiusuario(user): 
    try: 
        consultados = verifica_registro()
        devedor = str(user).lower()
        devedor = devedor.capitalize()
        pos = consultados.index(devedor)
        print("Deseja realmente remover {} da lista, você perdera os valores salvos".format(devedor))
        o = int(input("[1] - Sim \n[2] - Não\n"))
    except:
        print('Ocorreu um erro')
        menu()
        
    if o == 1:
        planilha1.delete_rows(pos+1)


def sobre():
    messagebox.showinfo(title="Debt Book",message="Programa gratuito desevolvido para distribuição com foco no pequeno comerciante, com este programa você será capaz de manusear dados sem ter nenhum conhecimento de planilhas.\nPara maiores informações ou sugestões: Rafaelmafra@live.com")

app.title(".::.Debt Book.::. Caderneta de Dívidas")
app.geometry("510x500+300+100")
app.configure(background="#5f9ea0")
quadro = Frame(app,borderwidth="2",bg="#5f9ea0",relief="groove")
quadro.place(x=215,y=100,width=290,height=350)

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
Button(app,text="[1] - Cadastrar Devedor/Dívida ", command=cadastro).place(x=10,y=100,width=200,height=40)
Button(app,text="[2] - Listar Devedores", command=leitura).place(x=10,y=150,width=200,height=40)
Button(app,text="[3] - Exibir Relatório detalhado", command=relatorio).place(x=10,y=200,width=200,height=40)
Button(app,text="[4] - Deletar Devedor", command=excluir).place(x=10,y=250,width=200,height=40)
Button(app,text="[5] - Sair do Programa", command=sair).place(x=10,y=300,width=200,height=40)



lb1 = Label(quadro, text="Bem Vindo!!", bg="#5f9ea0", font=("Arial",13))
lb1.pack()





app.mainloop()