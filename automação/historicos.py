import time
import pyautogui
from tkinter import messagebox
import tkinter as tk
from automação import repetir_tecla
from ocr import *
from planilhas import ler_alunos

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
    
    pyautogui.hotkey('alt','tab')
    time.sleep(1)
    
    esperar_elemento("./imagens/hub_aberto.png")

    registrar_ocorrencias(arquivo_alunos, titulo, descricao)

def registrar_ocorrencias(arquivo_alunos, titulo_ocorrencia, descricao_ocorrencia):
    # Ler os contatos
    alunos = ler_alunos(arquivo_alunos)
    ocorrencias_registradas = 0

    pyautogui.press('alt')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    for aluno in alunos:
        try:
            nome_aluno = aluno['nome']  # Nome do aluno
            pyperclip.copy(nome_aluno)

            esperar_elemento('./imagens/pesquisa_aluno.png')
            pesquisa_aluno = localizar_elemento('./imagens/pesquisa_aluno.png')
            pyautogui.click(pesquisa_aluno)
    
            pyautogui.hotkey('ctrl','a')
            time.sleep(1)

            # Pressionar Enter para enviar a imagem
            pyautogui.press('backspace')
            time.sleep(1)

            pyautogui.hotkey('ctrl','v')
            pyautogui.press('enter')
            pyautogui.press('enter')
            time.sleep(3)

            aluno_existe = verificar_existencia('pesquisa_aluno')
            
            if aluno_existe:
                esperar_elemento('./imagens/aluno_encontrado.png')
                aluno_encontrado = localizar_elemento('./imagens/aluno_encontrado.png')
                pyautogui.doubleClick(aluno_encontrado)
                time.sleep(7)
                
            pyautogui.hotkey('ctrl','tab')
            time.sleep(2)
            
            pyperclip.copy(titulo_ocorrencia)
            repetir_tecla('shift','tab',total_repeticoes=5)
            pyautogui.hotkey('ctrl','v')
            time.sleep(2)

            pyperclip.copy(descricao_ocorrencia)
            repetir_tecla('shift','tab', total_repeticoes=2)
            pyautogui.hotkey('ctrl','v')
            time.sleep(2)

            pyautogui.hotkey('alt','s')
            
            ocorrencias_registradas += 1
            
            time.sleep(4)

        except:
            if ocorrencias_registradas == 0:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, não consegui registrar nenhuma ocorrência :(")
            else:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, só consegui registrar {ocorrencias_registradas} ocorrências :(")