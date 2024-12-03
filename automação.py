import time
import pyautogui
from tkinter import filedialog, messagebox
from ttkbootstrap import ttk
import tkinter as tk
from ocr import *
import pandas as pd
import os
import pyperclip

def localizar_elemento(app, elemento):
    localizacao = None
    while True:
        time.sleep(2)  # Aguarda 2 segundos antes de tentar novamente

        if elemento == 'menu_iniciar':
            localizacao = pyautogui.locateOnScreen('./imagens/menu_iniciar.png', confidence=0.8)  # Ajuste a confiança, se necessário

        if app == 'whatsapp':
            if elemento == 'whatsapp_encontrado':
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

        elif app == 'hub':
            if elemento == 'pesquisa_aluno':
                time.sleep(5)
                localizacao = pyautogui.locateOnScreen('./imagens/pesquisa_aluno.png')  # Ajuste a confiança, se necessário
            elif elemento == 'aluno_encontrado':
                time.sleep(3)
                localizacao = pyautogui.locateOnScreen('./imagens/aluno_encontrado.png')  # Ajuste a confiança, se necessário
        
        if localizacao:
                break
    
    return localizacao


# Função para anexar arquivo
def anexar_planilha(campo_planilha):
    caminho_planilha = filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Arquivos do Excel", "*.xls")])
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
    localizar_elemento('whatsapp', 'menu_iniciar')
    
    # Usar o pyautogui para digitar whatsapp
    pyautogui.write('whatsapp')
    localizar_elemento('whatsapp', 'whatsapp_encontrado')

    # Pressionar Enter para abrir o app
    pyautogui.press('enter')
    localizar_elemento('whatsapp', 'whatsapp_aberto')

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
            nome_aluno = contato['aluno']  # Nome do aluno
            numero_telefone = contato['celular'].replace("(", "").replace(")", "").replace(" ", "")
            numero_telefone = numero_telefone[:2] + numero_telefone[3:]  # Remove o índice 2 (que é o terceiro número)
            mensagem_personalizada = mensagem_template.replace("<nome_aluno>", nome_aluno)
            pyperclip.copy(mensagem_personalizada)

            # Clicar no botão de alternar
            pyautogui.press('enter')
            time.sleep(2)  
            pyautogui.hotkey('ctrl','n')
            localizar_elemento('whatsapp', 'nova_conversa')
            
            # Usar o pyautogui para digitar o número de telefone do contato
            pyautogui.write(f'{numero_telefone}')
            time.sleep(1)   
            
            whatsapp_existe = verificar_existencia('pesquisa_whatsapp')

            if whatsapp_existe:
                pyautogui.press('tab')
                pyautogui.press('tab')  
                time.sleep(1)

                pyautogui.press('enter')  
                time.sleep(1)
                
                #Verifica se há imagem
                if len(imagem) > 0:
                    botao_anexar = localizar_elemento('whatsapp', 'anexar')
                    pyautogui.click(botao_anexar)

                    pyautogui.press('tab')
                    time.sleep(1)
                    pyautogui.press('enter')  
                    time.sleep(2)

                    # Colar o caminho da imagem
                    pyautogui.hotkey('ctrl', 'v')  # Colar o caminho da imagem no campo
                    time.sleep(2)

                    pyautogui.press('enter')

                    localizar_elemento('whatsapp', 'aba_anexar')

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

# Função para ler os contatos de um arquivo TXT
def ler_contatos(caminho_planilha):
    try:
        # Lê a planilha usando pandas
        df = pd.read_excel(caminho_planilha, header=2)

        # Certifique-se de que os nomes das colunas estão corretos
        contatos = []
        for _, row in df.iterrows():
            contatos.append({
                "aluno": row["Aluno"],         # Nome do aluno
                "celular": str(row["Celular"])  # Número de telefone
            })
        
        print(contatos)
        return contatos
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []
    
def gerar_comunicado(tipo_comunicado_var, campo_mensagem):
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

            pesquisa_aluno = localizar_elemento('hub','pesquisa_aluno')
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
                aluno_encontrado = localizar_elemento('hub','aluno_encontrado')
                pyautogui.doubleClick(aluno_encontrado)
                time.sleep(7)
                
            pyautogui.hotkey('ctrl','tab')
            time.sleep(2)
            
            pyperclip.copy(titulo_ocorrencia)
            for i in range(1,6):
                pyautogui.hotkey('shift','tab')
                time.sleep(0.5)
            pyautogui.hotkey('ctrl','v')
            time.sleep(2)

            pyperclip.copy(descricao_ocorrencia)
            for i in range(1,3):
                pyautogui.hotkey('shift','tab')
                time.sleep(0.5)
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
                
# Função para ler os contatos de um arquivo TXT
def ler_alunos(caminho_planilha):
    try:
        # Lê a planilha usando pandas
        df = pd.read_excel(caminho_planilha)
        
        # Certifique-se de que os nomes das colunas estão corretos
        alunos = []
        for _, row in df.iterrows():
            alunos.append({
                "nome": row["Nome"],         # Nome do aluno
            })
        return alunos
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []