# %%
import tkinter as tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from tkinter import messagebox
from time import sleep
from pyautogui import click, write, press, PAUSE, hotkey, center
from datetime import datetime
import pyscreeze 

PAUSE = 2

# %% [markdown]
# iconbitmap('caminho/para/ícone.ico'): Define um ícone para a janela (deve ser um arquivo .ico).

# %%
caminho_arquivo = ''

def encontrar_imagem(imagem):
    while not pyscreeze.locateOnScreen(imagem, confidence=0.6):  # reconhecimento de imagem
       sleep(1)
    encontrou = pyscreeze.locateOnScreen(imagem, confidence=0.6)
    click(center(encontrou))

def selecionar_arquivo():
    global caminho_arquivo
    arquivo = askopenfilename(title='Selecione o arquivo excel')
    var_caminhoarquivo.set(arquivo)
    try:
        df = pd.read_excel(arquivo)
        texto = tk.Label(text='Arquivo Selecionado',bg='white')
        texto.place(y=53,x=220,width=150)
        caminho_arquivo = arquivo
    except:
        texto = tk.Label(text='Arquivo não selecionado',bg='white')
        texto.place(y=53,x=220,width=150)

def cancelar():
    janela.destroy()
    
def iniciar():
    resposta = messagebox.askyesno("Alerta", "Está logado no Whatsapp?")
    
    if resposta == True:
        
        base = pd.read_excel(caminho_arquivo)
        sleep(5)
        try:
            for i, linha in base.iterrows():
                name, celular, valor, data, x = linha
                
                encontrar_imagem(r'imagens\pesquisar.png')
                write(celular)
                press('enter')
                sleep(5)
                
                valor = f'{valor:,.2f}'
                valor = valor.replace('.','_').replace(',','.').replace('_',',')
                ano, mes, dia = f'{data}'[0:10].split('-')
                
                encontrar_imagem(r'imagens\mensagem.png')
                write(f'Oi *{name}*, espero que esteja bem!!')
                hotkey('shift','enter')
                write(f'Estou entrando em contato para avisar que a sua fatura de *R$ {valor}* vence em *{dia}-{mes}-{ano[:2]}*.')
                hotkey('shift','enter')
                write('Um bom dia!!')
                
                press('enter')
                press('esc')
                
                ult_cobranca = datetime.now()
                base.loc[i,'Ultima_cobrança'] = f'{ult_cobranca.day}-{ult_cobranca.month}-{ult_cobranca.year}'  
                
                sleep(3)  
        
            base.to_excel(caminho_arquivo, index=False)
            hotkey('alt','space','n')
            messagebox.showinfo("Aviso", "Mensagens Enviadas com sucesso.")
                
        except:
            resposta = messagebox.showerror("Erro", "Alguma coisa deu errado. Feche e Tente denovo.")
            if resposta:
                janela.destroy()
            else:
                janela.destroy()
                
    else:
        pass
    
    
janela = tk.Tk()

#Configuração da janela
janela.title('Cobrança Automática')
janela.geometry('400x150')
janela.configure(bg = 'White')

texto1 = tk.Label(text='Antes de selecionar o arquivo abra o Whatsapp e conecte',bg='white')
texto1.configure(font=('arial', 10,'bold'))
texto1.place(x=10,y=5)

botao_arquivo = tk.Button(text='Selecione o arquivo (.xlsx)',command=selecionar_arquivo)
botao_arquivo.place(x=10,y=50,width=200)
print(botao_arquivo)

var_caminhoarquivo = tk.StringVar()

botao_cancelar = tk.Button(text='Cancelar',command=cancelar)
botao_cancelar.place(y=100,x=220,width=150)

botao_Iniciar = tk.Button(text='Iniciar',command=iniciar)
botao_Iniciar.place(y=100,x=25,width=150)

janela.mainloop()


