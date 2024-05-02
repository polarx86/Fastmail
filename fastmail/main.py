from tkinter import *
import win32com.client as win32

def enviaremail():
    destino = str(entry_destinatario.get())
    assunto = str(entry_assunto.get())
    mensagem = entry_msg.get()

    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)

    email.To = destino
    email.Subject = assunto
    email.HTMLBody = f'''
    <p>{mensagem}</p>
    '''
    email.Send()
    label_aviso['text'] = 'email enviado!'

cor_letras = '#c402fa'
cor_fundo = '#171617'

janela = Tk()
janela.title("fastmail")
janela.geometry("700x500")
janela.iconphoto(False, PhotoImage(file='icon.png'))
janela.resizable(width=False, height=False)
janela.config(bg=cor_fundo)

label_titulo = Label(janela, width=10, height=2, text='FASTMAIL', font=('Times 30 italic bold'), bg=cor_fundo, fg=cor_letras)
label_titulo.grid(row=0, column=1)

label_destinatario = Label(janela, width=12, height=2, text='DESTINO:', font='Times 15 bold', bg=cor_fundo, fg=cor_letras, anchor='w')
label_destinatario.grid(row=1, column=0, padx=20)

entry_destinatario = Entry(janela, width=30)
entry_destinatario.place(x=175, y=115)

label_assunto = Label(janela, width=12, height=2, text='ASSUNTO:', font='Times 15 bold', bg=cor_fundo, fg=cor_letras, anchor='w')
label_assunto.grid(row=2, column=0)

entry_assunto = Entry(janela, width=30)
entry_assunto.place(x=175, y=167)

label_msg = Label(janela, width=12, height=2, text='MENSAGEM:', font='Times 15 bold', bg=cor_fundo, fg=cor_letras, anchor='w')
label_msg.grid(row=3, column=0)

entry_msg = Entry(janela, width=30)
entry_msg.place(x=175, y=218)

botao = Button(janela,command=enviaremail, width=30, height=2, text='enviar email', font=('Arial 0 bold'), relief='raised', fg='white', bg=cor_letras)
botao.grid(row=4, column=1, padx=0, pady=40)

label_aviso = Label(janela, width=12, height=2, text='', font='Times 20 bold', bg=cor_fundo, fg=cor_letras, anchor='n')
label_aviso.grid(row=5, column=1)

janela.mainloop()
