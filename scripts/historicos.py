from tkinter import messagebox
import tkinter as tk
from scripts.ocr import *
from bot import registrar_ocorrencias

import pyperclip

def gerar_ocorrencia(tipo_ocorrencia_var, campo_titulo, campo_descricao):
    tipo_ocorrencia = tipo_ocorrencia_var.get()  # Pega a opção selecionada

    # Mensagens para cada tipo de comunicado
    if tipo_ocorrencia == "falta":
        titulo = "Falta - <Data>"
        descricao = "Feito contato, aguardando retorno."
    elif tipo_ocorrencia == "multirao":
        titulo = "Multirão de reposição"
        descricao = "Enviado Convite para o(a) aluno(a) para participar do multirão de reposição, explicando o objetivo da iniciativa, a gratuidade e as datas disponíveis. Solicitamos que informassem a data e horário de preferência para organizarmos a programação."
    elif tipo_ocorrencia == "comportamento":
        titulo = "Acompanhamento pedagógico"
        descricao = "Aluno(a) anda conversando bastante em sala, o que acaba gerando atraso em suas aulas e dificuldade do mesmo em realizar os testes, já que ele acaba não prestando atenção em suas aulas teóricas. Na maioria das vezes, ele acaba realizando apenas uma aula em vez de duas."
    elif tipo_ocorrencia == "prova":
        titulo = "Prova - <Módulo>"
        descricao = "Nota: x"
    elif tipo_ocorrencia == "atividades":
        titulo = "Atividades - <Módulo>"
        descricao = "Nota: x"
    elif tipo_ocorrencia == "1_dia_de_aula":
        titulo = "Acompanhamento - 1° dia de aula"
        descricao = "Aluno iniciou seu curso na unidade Manaus Zona Oeste, onde conheceu os professores e a plataforma, realizou a aula inaugural e linha da vida. Foi observado que o aluno possui uma certa dificuldade com..."
    elif tipo_ocorrencia == "plantao":
        titulo = "Plantão de dúvidas"
        descricao = "Foi identificado uma certa dificuldade do aluno(a) em realizar a prova prática, sendo então oferecido o plantão de dúvidas. Em contato com responsável para informar a situação e horários disponíveis, aguardando retorno."

    # Exibe o título no campo de texto
    campo_titulo.delete("1.0", tk.END)  # Limpa a área de texto antes de inserir
    campo_titulo.insert(tk.END, titulo)

    # Exibe a descrição no campo de texto
    campo_descricao.delete("1.0", tk.END)  # Limpa a área de texto antes de inserir
    campo_descricao.insert(tk.END, descricao)

def preparar_registros(campo_planilha, campo_titulo, campo_descricao):
    arquivo_alunos = campo_planilha.get()
    titulo = campo_titulo.get("1.0", tk.END)
    descricao = campo_descricao.get("1.0", tk.END)

    if len(arquivo_alunos) == 0:
        messagebox.showinfo("Oops!", "Insira uma planilha de nomes para o envio das mensagens!")
        return
    elif len(titulo.strip()) == 0:
        messagebox.showinfo("Oops!", "Insira o título da ocorrência!")
        return
    elif len(descricao.strip()) == 0:
        messagebox.showinfo("Oops!", "Insira a descrição da ocorrência!")
        return
    
    messagebox.showinfo("Aviso!", "Certifique-se de que o HUB esteja aberto e atrás do easyLog!")

    registrar_ocorrencias(arquivo_alunos, titulo, descricao)

