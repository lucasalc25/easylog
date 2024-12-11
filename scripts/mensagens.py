from tkinter import messagebox
import tkinter as tk
from bot import *
from scripts.ocr import *
import os

def gerar_mensagem(tipo_comunicado_var, campo_mensagem, campo_data, campo_hora, campo_tema, campo_data_inicial, campo_data_final):
    tipo_comunicado = tipo_comunicado_var.get()  # Pega a opção selecionada
    data = campo_data.get()

    if len(data) == 0 and tipo_comunicado != "multirão":
        messagebox.showinfo("Oops!", "Insira a data para elaborar a mensagem!")
        return

    # Mensagens para cada tipo de comunicado
    if tipo_comunicado == "falta":
        mensagem = f"""Olá, tudo bem? Aqui é o instrutor <nome_educador> da Microlins.

Verifiquei que o(a) aluno(a) <nome_aluno> não compareceu à aula de {data}. Poderia nos informar o motivo da falta pra registro em sistema? Lembrando, para que isso não gere atrasos, deve ser feita a reposição ta bom?

Fico no seu aguardo. 
Obrigado desde já!"""
    elif tipo_comunicado == "multirão":
        data_inicial = campo_data_inicial.get()
        data_final = campo_data_final.get()
        mensagem = f"Atenção! Estamos organizando um multirão de reposições gratuitas dos dias {data_inicial} a {data_final} para faltas dentro deste mês. Não perca essa chance de deixar seu curso em dia!"
        if len(data_inicial) == 0:
            messagebox.showinfo("Oops!", "Insira a data de início do multirão para elaborar a mensagem!")
            return
        elif len(data_final) == 0:
            messagebox.showinfo("Oops!", "Insira a data de fim do multirão para elaborar a mensagem!")
            return
    elif tipo_comunicado == "reuniao_de_pais":
        mensagem = f"No dia {data} acontecerá a reunião de pais e professores do curso de x."
    elif tipo_comunicado == "oficina":
        tema = campo_tema.get()
        mensagem = f"Anunciamos a oficina de {tema}, com início a partir do dia {data}."
        if len(data) == 0:
            messagebox.showinfo("Oops!", "Insira a data da oficina para elaborar a mensagem!")
            return
        elif len(tema) == 0:
            messagebox.showinfo("Oops!", "Insira o tema da oficina para elaborar a mensagem!")
            return
    elif tipo_comunicado == "formatura":
        hora = campo_hora.get()
        mensagem = f"Convidamos o(a) aluno(a) <nome_aluno> para a nossa formatura, que acontecerá no dia {data} a partir das {hora}."
        if len(hora) == 0:
            messagebox.showinfo("Oops!", "Insira a hora da formatura para elaborar a mensagem!")
            return
    elif tipo_comunicado == "feriado":
        mensagem = f"Lembrete sobre o feriado do dia {data}."

    # Exibe a mensagem no campo de texto
    campo_mensagem.delete("1.0", tk.END)  # Limpa a área de texto antes de inserir
    campo_mensagem.insert(tk.END, mensagem)

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
   
