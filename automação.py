import time
import pyautogui
import keyboard
from tkinter import filedialog, messagebox
from ttkbootstrap import ttk
import tkinter as tk
from ocr import *
import pandas as pd
import os
import pyperclip

def esperar_elemento(elemento):
    localizacao = None
    while True:
        time.sleep(2)  # Aguarda 2 segundos antes de tentar novamente

        if elemento == 'menu_iniciar':
            localizacao = pyautogui.locateOnScreen('./imagens/menu_iniciar.png', confidence=0.8)  # Ajuste a confiança, se necessário
        elif elemento == 'whatsapp_encontrado':
            localizacao = pyautogui.locateOnScreen('./imagens/whatsapp_encontrado.png', confidence=0.8)  # Ajuste a confiança, se necessário
        elif elemento == 'whatsapp_aberto':
            time.sleep(2)
            localizacao = pyautogui.locateOnScreen('./imagens/whatsapp_aberto.png')  # Ajuste a confiança, se necessário
        elif elemento == 'nova_conversa':
            localizacao = pyautogui.locateOnScreen('./imagens/nova_conversa.png', confidence=0.8)  # Ajuste a confiança, se necessário
        elif elemento == 'pesquisa':
            localizacao = pyautogui.locateOnScreen('./imagens/campo_pesquisa.png', confidence=0.5)  # Ajuste a confiança, se necessário
        elif elemento == 'mensagem':
            localizacao = pyautogui.locateOnScreen('./imagens/campo_mensagem.png', confidence=0.8)  # Ajuste a confiança, se necessário
        elif elemento == 'anexar':
            localizacao = pyautogui.locateOnScreen('./imagens/anexar.png', confidence=0.8)  # Ajuste a confiança, se necessário
        elif elemento == 'aba_anexar':
            localizacao = pyautogui.locateOnScreen('./imagens/aba_anexar.png', confidence=0.8)  # Ajuste a confiança, se necessário
        if localizacao:
            break
    
    return localizacao


# Função para anexar arquivo
def anexar_planilha(campo_planilha):
    caminho_planilha = filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Arquivo do Excel", "*.xlsx")])
    if caminho_planilha:
        campo_planilha.delete(0, tk.END)
        campo_planilha.insert(0, caminho_planilha)


# Função para anexar imagem
def anexar_imagem(imagem):
    caminho = filedialog.askopenfilename(
        title="Selecione uma imagem",
        filetypes=[("Arquivos de imagem", "*.png;*.jpg;*.jpeg")]
    )
    if caminho:
        imagem.delete(0, tk.END)  # Limpa o conteúdo atual
        imagem.insert(0, caminho)  # Exibe o caminho do arquivo


# Função para ler os contatos de um arquivo TXT
def ler_contatos(caminho_planilha):
    try:
        # Lê a planilha usando pandas
        df = pd.read_excel(caminho_planilha)
        
        # Certifique-se de que os nomes das colunas estão corretos
        contatos = []
        for _, row in df.iterrows():
            contatos.append({
                "nome": row["Nome"],         # Nome do aluno
                "telefone": str(row["Telefone"])  # Número de telefone
            })
        return contatos
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []


# Função para iniciar envio
def preparar_envio(campo_planilha, campo_nome_professor, campo_dia_falta, campo_mensagem, campo_imagem):
    arquivo_contatos = campo_planilha.get()
    imagem = campo_imagem.get()
    mensagem_template = campo_mensagem.get("1.0", "end")
    nome_professor = campo_nome_professor.get()
    dia_falta = campo_dia_falta.get()
    
    if len(arquivo_contatos) == 0:
        messagebox.showinfo("Ops!", "Insira uma planilha de contatos para o envio das mensagens!")
        return
    elif len(nome_professor) == 0 or len(dia_falta) == 0:
        messagebox.showinfo("Ops!", "Insira as variáveis para elaborar a mensagem!")
        return
    elif len(imagem) == 0 and len(mensagem_template) == 0:
        messagebox.showinfo("Ops!", "Insira uma mensagem ou imagem para o envio!")
        return
    
    if len(imagem) > 0:
        imagem = os.path.normpath(imagem)
    
    # Pressionar Windows para abrir a conversa
    pyautogui.press('win')  
    esperar_elemento('menu_iniciar')
    
    # Usar o pyautogui para digitar whatsapp
    pyautogui.write('whatsapp')
    esperar_elemento('whatsapp_encontrado')

    # Pressionar Enter para abrir o app
    pyautogui.press('enter')
    esperar_elemento('whatsapp_aberto')

    enviar_mensagens(arquivo_contatos, imagem, mensagem_template)


def mensagem_para_verificacao(nome_professor, dia_falta, text_mensagem):
    mensagem = f"""Olá, tudo bem? Aqui é o(a) professor(a) {nome_professor.get()} da Microlins.

Verifiquei que o(a) aluno(a) <nome_aluno> não compareceu à aula de {dia_falta.get()}. Poderia nos informar o motivo da falta pra registro em sistema? Lembrando, para que isso não gere atrasos, deve ser feita a reposição ta bom?

Fico no seu aguardo. 
Obrigado desde já!"""

    text_mensagem.insert("1.0", mensagem)

# Função para gerar mensagem final substituindo placeholders
def mensagem_final(mensagem_template, nome_aluno):
    mensagem = mensagem_template.replace("<nome_aluno>", nome_aluno)
    return mensagem
        
# Função para enviar uma imagem para os contatos
def enviar_mensagens(arquivo_contatos, imagem, mensagem_template):
    # Ler os contatos
    contatos = ler_contatos(arquivo_contatos)
    
    for contato in contatos:
        try:
            mensagens_enviadas = 0
            nome_aluno = contato['nome']  # Nome do aluno
            numero_telefone = contato['telefone'].replace("(", "").replace(")", "").replace(" ", "")
            numero_telefone = numero_telefone[:2] + numero_telefone[3:]  # Remove o índice 2 (que é o terceiro número)
            mensagem_personalizada = mensagem_template.replace("<nome_aluno>", nome_aluno)
            pyperclip.copy(mensagem_personalizada)

            # Clicar no botão de alternar
            pyautogui.press('enter')
            time.sleep(2)  
            pyautogui.hotkey('ctrl','n')
            esperar_elemento('nova_conversa')
            
            # Usar o pyautogui para digitar o número de telefone do contato
            pyautogui.write(f'{numero_telefone}')
            time.sleep(1)   
            
            whatsapp_existe = verificar_existencia()

            if whatsapp_existe:
                pyautogui.press('tab')
                pyautogui.press('tab')  
                time.sleep(1)

                pyautogui.press('enter')  
                time.sleep(1)
                
                #Verifica se há imagem
                if len(imagem) > 0:
                    botao_anexar = esperar_elemento('anexar')
                    pyautogui.click(botao_anexar)

                    pyautogui.press('tab')
                    time.sleep(1)
                    pyautogui.press('enter')  
                    time.sleep(2)

                    # Colar o caminho da imagem
                    pyautogui.hotkey('ctrl', 'v')  # Colar o caminho da imagem no campo
                    time.sleep(2)

                    pyautogui.press('enter')

                    esperar_elemento('aba_anexar')

                    if mensagem_template:
                        # Usar o pyautogui para colar a mensagem
                        pyautogui.hotkey('ctrl','v')
                        time.sleep(1)
                    
                    time.sleep(1)
                    # Pressionar Enter para enviar a imagem
                    pyautogui.press('enter')
                    mensagens_enviadas += 1

                else:
                    # Usar o pyautogui para colar a mensagem
                    pyautogui.hotkey('ctrl','v')
                    time.sleep(2)

                    # Clicar no botão de alternar
                    pyautogui.press('enter')
                    mensagens_enviadas += 1
                
            else:
                pyautogui.hotkey('ctrl','a')
                time.sleep(1)

                # Pressionar Enter para enviar a imagem
                pyautogui.press('backspace')
            
            time.sleep(2)  # Aguarde um tempo antes de enviar para o próximo contato
        except:
            if mensagens_enviadas == 0:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, não consegui enviar nenhuma mensagem :(")
            else:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, só consegui enviar {mensagens_enviadas} mensagens :(")
                            


    messagebox.showinfo("Concluído!", "Mensagem enviada para todos os contatos")

def aplicar_comunicado(tipo_comunicado_var, campo_mensagem):
    tipo_comunicado = tipo_comunicado_var.get()  # Pega a opção selecionada

    # Mensagens para cada tipo de comunicado
    if tipo_comunicado == "multirão":
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


