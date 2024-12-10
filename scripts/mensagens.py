from tkinter import messagebox
import tkinter as tk
from bot import *
from scripts.ocr import *
import os

# Função para iniciar envio
def preparar_envio(campo_planilha, campo_mensagem, campo_imagem):
    arquivo_contatos = campo_planilha.get()
    imagem = campo_imagem.get()
    mensagem_template = campo_mensagem.get("1.0", "end")


    if len(arquivo_contatos) == 0:
        messagebox.showinfo("Oops!", "Insira uma planilha de contatos para o envio das mensagens!")
        return
    
    if len(imagem.strip()) == 0 and len(mensagem_template.strip()) == 0:
        messagebox.showinfo("Oops!", "Insira uma mensagem ou imagem para o envio!")
        return
    
    if len(imagem) > 0:
        imagem = os.path.normpath(imagem)

    enviar_mensagens(arquivo_contatos, imagem, mensagem_template)


def personalizar_mensagem(nome_aluno, nome_educador, mensagem_template):
    mensagem_personalizada = mensagem_template.replace("<nome_aluno>", nome_aluno)
    mensagem_personalizada = mensagem_personalizada.replace("<nome_educador>", nome_educador)

    return mensagem_personalizada
   
def gerar_mensagem(tipo_comunicado_var, campo_mensagem, campo_data, campo_oficina):
    tipo_comunicado = tipo_comunicado_var.get()  # Pega a opção selecionada
    data = campo_data.get()
    oficina = campo_oficina.get()

    # Mensagens para cada tipo de comunicado
    if tipo_comunicado == "falta":
        mensagem = f"""Olá, tudo bem? Aqui é o instrutor <nome_educador> da Microlins.

Verifiquei que o(a) aluno(a) <nome_aluno> não compareceu à aula de {data}. Poderia nos informar o motivo da falta pra registro em sistema? Lembrando, para que isso não gere atrasos, deve ser feita a reposição ta bom?

Fico no seu aguardo. 
Obrigado desde já!"""
        if len(data) == 0:
            messagebox.showinfo("Oops!", "Insira a data da falta para elaborar a mensagem!")
            return
    elif tipo_comunicado == "multirão":
        mensagem = f"Estamos organizando um multirão de reposições gratuitas até o dia {data} para faltas dentro do mês de yx"
        if len(data) == 0:
            messagebox.showinfo("Oops!", "Insira a data do multirão para elaborar a mensagem!")
            return
    elif tipo_comunicado == "reuniao_de_pais":
        mensagem = f"No dia {data} acontecerá a reunião de pais e professores do curso de x."
        if len(data) == 0:
            messagebox.showinfo("Oops!", "Insira a data da reunião para elaborar a mensagem!")
            return
    elif tipo_comunicado == "oficina":
        mensagem = f"Não perca a oficina de {oficina} no dia {data}."
        if len(data) == 0:
            messagebox.showinfo("Oops!", "Insira a data da oficina para elaborar a mensagem!")
            return
        elif len(oficina) == 0:
            messagebox.showinfo("Oops!", "Insira o tema da oficina para elaborar a mensagem!")
            return
    elif tipo_comunicado == "formatura":
        mensagem = f"Convidamos o(a) aluno(a) <nome_aluno> para a nossa formatura, que acontecerá no dia {data}."
        if len(data) == 0:
            messagebox.showinfo("Oops!", "Insira a data da formatura para elaborar a mensagem!")
            return
    elif tipo_comunicado == "feriado":
        mensagem = f"Lembrete sobre o feriado do dia {data}."
        if len(data) == 0:
            messagebox.showinfo("Oops!", "Insira a data do feriado para elaborar a mensagem!")
            return

    # Exibe a mensagem no campo de texto
    campo_mensagem.delete("1.0", tk.END)  # Limpa a área de texto antes de inserir
    campo_mensagem.insert(tk.END, mensagem)