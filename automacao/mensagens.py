import pyautogui
from tkinter import messagebox
import tkinter as tk
from automacao.ocr import *
import os
from bot import enviar_mensagens

# Função para iniciar envio
def preparar_envio(campo_planilha, campo_nome_professor, campo_dia_falta, campo_mensagem, campo_imagem):
    arquivo_contatos = campo_planilha.get()
    imagem = campo_imagem.get()
    mensagem_template = campo_mensagem.get("1.0", "end")
    nome_professor = campo_nome_professor.get()
    dia_falta = campo_dia_falta.get()
    print("Mensagem:", mensagem_template)
    
    if len(arquivo_contatos) == 0:
        messagebox.showinfo("Oops!", "Insira uma planilha de contatos para o envio das mensagens!")
        return
    elif len(nome_professor) == 0 or len(dia_falta) == 0:
        messagebox.showinfo("Oops!", "Insira as variáveis para elaborar a mensagem!")
        return
    elif len(imagem.strip()) == 0 and len(mensagem_template.strip()) == 0:
        messagebox.showinfo("Oops!", "Insira uma mensagem ou imagem para o envio!")
        return
    
    if len(imagem) > 0:
        imagem = os.path.normpath(imagem)
    
    # Pressionar Windows para abrir a conversa
    pyautogui.press('win')  
    esperar_elemento('./imagens/menu_iniciar.png')
    
    # Usar o pyautogui para digitar whatsapp
    pyautogui.write('whatsapp')
    esperar_elemento('./imagens/whatsapp_encontrado.png')

    # Pressionar Enter para abrir o app
    pyautogui.press('enter')
    esperar_elemento('./imagens/whatsapp_aberto.png')

    enviar_mensagens(arquivo_contatos, imagem, mensagem_template)


def personalizar_mensagem(nome_professor, dia_falta, campo_mensagem):
    if nome_professor.get() == 'Yasmim':
        mensagem = f"""Olá, tudo bem? Aqui é a instrutora {nome_professor.get()} da Microlins.

Verifiquei que o(a) aluno(a) <nome_aluno> não compareceu à aula de {dia_falta.get()}. Poderia nos informar o motivo da falta pra registro em sistema? Lembrando, para que isso não gere atrasos, deve ser feita a reposição ta bom?

Fico no seu aguardo. 
Obrigado desde já!"""

    else:
        mensagem = f"""Olá, tudo bem? Aqui é o instrutor {nome_professor.get()} da Microlins.

Verifiquei que o(a) aluno(a) <nome_aluno> não compareceu à aula de {dia_falta.get()}. Poderia nos informar o motivo da falta pra registro em sistema? Lembrando, para que isso não gere atrasos, deve ser feita a reposição ta bom?

Fico no seu aguardo. 
Obrigado desde já!"""

    campo_mensagem.delete("1.0", tk.END)
    campo_mensagem.insert("1.0", mensagem)

    return mensagem

# Função para gerar mensagem final substituindo placeholders
def mensagem_final(mensagem_template, nome_aluno):
    mensagem = mensagem_template.replace("<nome_aluno>", nome_aluno)
    return mensagem
          
def gerar_mensagem(tipo_comunicado_var, campo_mensagem, nome_professor, dia_falta):
    tipo_comunicado = tipo_comunicado_var.get()  # Pega a opção selecionada

    if len(nome_professor.get()) == 0 or len(dia_falta.get()) == 0:
        messagebox.showinfo("Oops!", "Insira as variáveis para gerar a mensagem!")
        return

    # Mensagens para cada tipo de comunicado
    if tipo_comunicado == "falta":
        mensagem = personalizar_mensagem(nome_professor, dia_falta, campo_mensagem)
    elif tipo_comunicado == "multirão":
        mensagem = "Estamos organizando um multirão de reposições gratuitas até o dia x para faltas dentro do mês de yx"
    elif tipo_comunicado == "reuniao_de_pais":
        mensagem = "No dia x acontecerá a reunião de pais e professores do curso de x."
    elif tipo_comunicado == "oficina":
        mensagem = "Não perca a oficina de x no dia y."
    elif tipo_comunicado == "formatura":
        mensagem = "A formatura dos alunos acontecerá no dia X."
    elif tipo_comunicado == "feriado":
        mensagem = "Lembrete sobre o feriado do dia x."

    # Exibe a mensagem no campo de texto
    campo_mensagem.delete("1.0", tk.END)  # Limpa a área de texto antes de inserir
    campo_mensagem.insert(tk.END, mensagem)