import time
import keyboard
import pyautogui
from tkinter import filedialog, messagebox
import tkinter as tk
import os
import pyperclip
from pathlib import Path
from config import caminhos
from scripts.ocr import esperar_elemento, localizar_elemento, verificar_existencia
from scripts.planilhas import filtrar_faltosos, ler_alunos, ler_contatos

def criar_pastas():
    # Obtém o caminho da pasta Documentos do usuário
    caminho_documentos = Path.home() / "Documents"
    
    # Define o caminho completo da nova pasta
    caminho_pasta_easylog = os.path.join(caminho_documentos, 'EasyLog')
    caminho_pasta_data = os.path.join(caminho_pasta_easylog, 'Data')
    
    pasta_easylog = 'EasyLog'
    pasta_data = 'Data'
    
    # Cria a pasta, se ela não existir
    if not os.path.exists(caminho_pasta_easylog):
        os.makedirs(caminho_pasta_easylog)
        print(f"Pasta '{pasta_easylog}' criada em {caminho_documentos}")
    else:
        print(f"A pasta '{pasta_easylog}' já existe em {caminho_documentos}")
        
    # Cria a pasta, se ela não existir
    if not os.path.exists(caminho_pasta_data):
        os.makedirs(caminho_pasta_data)
        print(f"Pasta '{pasta_data}' criada em {caminho_pasta_easylog}")
    else:
        print(f"A pasta '{pasta_data}' já existe em {caminho_pasta_easylog}")

def repetir_tecla(*teclas, total_repeticoes):
    """
    Pressiona uma ou mais teclas repetidamente.
    
    :param teclas: Teclas a serem pressionadas (pode ser uma ou mais).
    :param total_repeticoes: Número de vezes que as teclas serão pressionadas.
    """
    for _ in range(total_repeticoes):
        if len(teclas) == 1:  # Apenas uma tecla
            pyautogui.press(teclas[0])
        else:  # Mais de uma tecla
            pyautogui.hotkey(*teclas)
        time.sleep(0.2)
    
# Função para anexar arquivo
def anexar_planilha(campo_planilha):
    caminho_planilha = filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Arquivos do Excel", "*.xlsx")])
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

# Função para enviar uma imagem para os contatos
def enviar_mensagens(arquivo_contatos, imagem, mensagem_template):
    from scripts.mensagens import personalizar_mensagem

    messagebox.showinfo("Aviso!", "Certifique-se de que o Whatsapp esteja conectado ao aparelho!")
    
    # Pressionar Windows para abrir a conversa
    pyautogui.press('win')  
    esperar_elemento(caminhos["menu_iniciar"])
    
    # Usar o pyautogui para digitar whatsapp
    pyautogui.write('whatsapp')
    esperar_elemento(caminhos["whatsapp_encontrado"])

    # Pressionar Enter para abrir o app
    pyautogui.press('enter')
    esperar_elemento(caminhos["whatsapp_aberto"])

    # Ler os contatos
    faltosos = ler_contatos(arquivo_contatos)
    mensagens_enviadas = 0
    
    for faltoso in faltosos:
        try:
            nome_aluno = faltoso['Aluno']  # Nome do aluno
            observacao = faltoso['Observacao']

            # Verifica se a coluna "Observação" está preenchida
            if observacao:
                print(f"Observação encontrada para {nome_aluno}. Pulando para o próximo aluno.")
                continue  # Pula para o próximo aluno se houver observação
            
            telefone = faltoso['Contato']
            telefone = telefone[:2] + telefone[3:]  # Remove o índice 2 (que é o terceiro número)
            nome_educador = faltoso['Educador']
            mensagem_personalizada = personalizar_mensagem(nome_aluno, nome_educador, mensagem_template)
            pyperclip.copy(mensagem_personalizada)

            pyautogui.hotkey('ctrl','n')
            esperar_elemento(caminhos["nova_conversa"])
            
            # Usar o pyautogui para digitar o número de telefone do contatoe
            pyautogui.write(f'{telefone}')
            time.sleep(2)   
            
            whatsapp_existe = verificar_existencia('pesquisa_whatsapp')

            if whatsapp_existe:
                pyautogui.press('tab')
                pyautogui.press('tab')  
                time.sleep(1)

                pyautogui.press('enter')  
                time.sleep(1)
                
                #Verifica se há imagem
                if len(imagem) > 0:
                    esperar_elemento(caminhos["anexar"])
                    botao_anexar = localizar_elemento(caminhos["anexar"])
                    pyautogui.click(botao_anexar)

                    pyautogui.press('tab')
                    time.sleep(1)
                    pyautogui.press('enter')  
                    time.sleep(2)

                    # Colar o caminho da imagem
                    pyautogui.hotkey('ctrl', 'v')  # Colar o caminho da imagem no campo
                    time.sleep(2)

                    pyautogui.press('enter')
                    
                    esperar_elemento(caminhos["aba_anexar"])

                    if mensagem_personalizada:
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


def registrar_ocorrencias(arquivo_alunos, data, titulo_ocorrencia, descricao_ocorrencia):
    repeticoes = 0
    descricao_ocorrencia = ''

    while not localizar_elemento(caminhos["hub_aberto"]):
        if repeticoes > 1:
            # Pressiona 'Alt' e 'Tab' duas vezes mantendo 'Alt' pressionado
            pyautogui.keyDown('alt')  # Mantém a tecla 'Alt' pressionada
            repetir_tecla('tab', total_repeticoes=repeticoes)
            pyautogui.keyUp('alt')  # Mantém a tecla 'Alt' pressionada
        else:
            pyautogui.hotkey('alt','tab')

        repeticoes += 1
        time.sleep(2) 

    # Ler os contatos
    alunos = ler_alunos(arquivo_alunos)
    ocorrencias_registradas = 0

    for aluno in alunos:
        try:
            nome_aluno = aluno['Aluno']  # Nome do aluno
            observacao = aluno['Observacao']

            # Verifica se a coluna "Observação" está preenchida
            if not observacao:
                print(f"Observação não encontrada para {nome_aluno}. Pulando para o próximo aluno.")
                continue  # Pula para o próximo aluno se houver observação
            
            print("\nTitulo da ocorrência:", titulo_ocorrencia)
            if titulo_ocorrencia.strip() == f"Falta - {data.strip()}":
                descricao_ocorrencia = observacao
                print("Descrição da ocorrência:",descricao_ocorrencia)

            pyperclip.copy(nome_aluno)

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

            esperar_elemento(caminhos["contratos"])
            pesquisa_aluno = localizar_elemento(caminhos["pesquisa_aluno"])
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
            
            esperar_elemento(caminhos["aluno_encontrado"])
            aluno_encontrado = localizar_elemento(caminhos["aluno_encontrado"])
            pyautogui.doubleClick(aluno_encontrado)
            
            esperar_elemento(caminhos["contrato_aberto"])
                
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
            
            esperar_elemento(caminhos["contratos"])

            pyautogui.press('esc')
            time.sleep(3)

        except:
            if ocorrencias_registradas == 0:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, não consegui registrar nenhuma ocorrência :(")
            else:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, só consegui registrar {ocorrencias_registradas} ocorrências :(")
    
    messagebox.showinfo("Concluído!", "Histórico registrado para todos os alunos!")           

def gerar_faltosos_e_educadores(data_inicial, data_final):
    repeticoes = 0

    while not localizar_elemento(caminhos["hub_aberto"]):
        if repeticoes > 1:
            # Pressiona 'Alt' e 'Tab' duas vezes mantendo 'Alt' pressionado
            pyautogui.keyDown('alt')  # Mantém a tecla 'Alt' pressionada
            repetir_tecla('tab', total_repeticoes=repeticoes)
            pyautogui.keyUp('alt')  # Mantém a tecla 'Alt' pressionada
        else:
            pyautogui.hotkey('alt','tab')

        repeticoes += 1
        time.sleep(2) 

    pyautogui.press('alt')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    repetir_tecla('tab', total_repeticoes=3)
    pyautogui.press('enter')
    time.sleep(1)

    esperar_elemento(caminhos["faltas_por_periodo"])
    keyboard.write(str(data_inicial))
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    keyboard.write(str(data_final))

    pesquisar_faltosos = localizar_elemento(caminhos["pesquisar"])
    pyautogui.click(pesquisar_faltosos)

    esperar_elemento(caminhos["lista_faltosos"])
    exportar_faltosos = localizar_elemento(caminhos["exportar"])
    pyautogui.click(exportar_faltosos)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    caminho_planilhas = Path.home() / "Documents" / "EasyLog" / "Data" / "alunos_e_educadores.xls"
    caminho_planilhas = os.path.normpath(caminho_planilhas)
    campo_nome_planilha = localizar_elemento(caminhos["campo_nome_planilha"])
    pyautogui.click(campo_nome_planilha)
    time.sleep(1)
    pyautogui.write(caminho_planilhas)
    time.sleep(1)
    salvar = localizar_elemento(caminhos["salvar"])
    pyautogui.click(salvar)
    time.sleep(2)
    
    if localizar_elemento(caminhos["substituir_arquivo"]):
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        
def gerar_faltosos(data_inicial, data_final):
    messagebox.showinfo("Atenção!", "Certifique-se de ter feito o login no HUB!")

    gerar_faltosos_e_educadores(data_inicial, data_final)
        
    pyautogui.press('alt')
    time.sleep(1)
    repetir_tecla('tab', total_repeticoes=7)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    
    esperar_elemento(caminhos["presencas_e_faltas"])
    pyautogui.write(data_inicial)
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(data_final)
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('space')
    time.sleep(1)
    visualizar = localizar_elemento(caminhos["visualizar"])
    pyautogui.click(visualizar)
    
    esperar_elemento(caminhos["visu_presencas_e_faltas"])
    pyautogui.press('alt')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    repetir_tecla('tab', total_repeticoes=3)
    pyautogui.press('right')
    time.sleep(1)
    repetir_tecla('tab', total_repeticoes=5)
    pyautogui.press('enter')
    time.sleep(1)
    
    esperar_elemento(caminhos["opcoes_exportacao"])
    pyautogui.press('enter')
    time.sleep(1)
   
    caminho_planilha = Path.home() / "Documents" / "EasyLog" / "Data" / "faltosos.xls"
    caminho_planilha = os.path.normpath(caminho_planilha)
    campo_nome_planilha = localizar_elemento(caminhos["campo_nome_planilha"])
    pyautogui.click(campo_nome_planilha)
    time.sleep(1)
    pyautogui.write(caminho_planilha)
    time.sleep(1)
    salvar = localizar_elemento(caminhos["salvar"])
    pyautogui.click(salvar)
    time.sleep(2)
    
    if localizar_elemento(caminhos["substituir_arquivo"]):
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        
    esperar_elemento(caminhos["abrir_planilha"])
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    messagebox.showinfo("Atenção!", "Planilha de faltosos gerada!")
    
    filtrar_faltosos(caminho_planilha)

     
        
        
    
    