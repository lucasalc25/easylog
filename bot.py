import time
import keyboard
import pyautogui
from tkinter import filedialog, messagebox
import tkinter as tk
import os
import pyperclip
from pathlib import Path
import pandas as pd
from config import caminhos
from scripts.ocr import esperar_elemento, localizar_elemento, verificar_existencia
from scripts.planilhas import filtrar_faltosos_do_mes, filtrar_faltosos_do_dia, ler_registros, ler_faltosos_dia

def criar_pastas():
    # Obtém o caminho da pasta Documentos do usuário
    caminho_documentos = Path.home() / "Documents"
    
    # Define o caminho completo da nova pasta
    caminho_pasta_easylog = os.path.join(caminho_documentos, 'EasyLog')
    
    # Cria a pasta, se ela não existir
    if not os.path.exists(caminho_pasta_easylog):
        os.makedirs(caminho_pasta_easylog)
        print(f"\nPasta Easylog criada em {caminho_documentos}")
    else:
        print(f"\nPasta Easylog localizada em {caminho_documentos}")
        
        
def excluir_arquivo(caminho_arquivo):
    # Verifica se o arquivo existe antes de tentar removê-lo
    if os.path.exists(caminho_arquivo):
        print(f"Removendo arquivo '{caminho_arquivo}'...")
        os.remove(caminho_arquivo)
        print(f"\nArquivo '{caminho_arquivo}' removido com sucesso!")
    else:
        print(f"\nArquivo '{caminho_arquivo}' não encontrado.")

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
        
def procurar_hub():
    repeticoes = 0
    
    if localizar_elemento(caminhos["hub_aberto"]):
        pyautogui.hotkey('alt','tab')
        time.sleep(2) 
    else:             
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
 

def abrir_aba(aba):   
    if aba == 'faltas_por_periodo':
        aba_pedagogico = localizar_elemento(caminhos["aba_pedagogico"])
        pyautogui.click(aba_pedagogico)
        esperar_elemento(caminhos["abrir_faltas_por_periodo"])
        abrir_faltas_por_periodo = localizar_elemento(caminhos["abrir_faltas_por_periodo"])
        pyautogui.click(abrir_faltas_por_periodo)
        esperar_elemento(caminhos["faltas_por_periodo"])
    elif aba == 'presencas_e_faltas':
        aba_relatorios = localizar_elemento(caminhos["aba_relatorios"])
        pyautogui.click(aba_relatorios)
        esperar_elemento(caminhos["relatorios_pedagogico"])
        relatorios_pedagogico = localizar_elemento(caminhos["relatorios_pedagogico"])
        pyautogui.click(relatorios_pedagogico)
        esperar_elemento(caminhos["abrir_presencas_e_faltas"])
        abrir_presencas_e_faltas = localizar_elemento(caminhos["abrir_presencas_e_faltas"])
        pyautogui.click(abrir_presencas_e_faltas)
        esperar_elemento(caminhos["presencas_e_faltas"])
    elif aba == 'controle_pedagogico':
        aba_pedagogico = localizar_elemento(caminhos["aba_pedagogico"])
        pyautogui.click(aba_pedagogico)
        esperar_elemento(caminhos["abrir_controle_pedagogico"])
        abrir_controle_pedagogico = localizar_elemento(caminhos["abrir_controle_pedagogico"])
        pyautogui.click(abrir_controle_pedagogico)
        esperar_elemento(caminhos["controle_pedagogico"])

def exportar_planilha(caminho_destino):
    print(f"Exportando planilha '{caminho_destino}...")
    caminho_destino = os.path.normpath(caminho_destino)
    campo_nome_planilha = localizar_elemento(caminhos["campo_nome_planilha"])
    pyautogui.click(campo_nome_planilha)
    time.sleep(1)
    keyboard.write(caminho_destino)
    time.sleep(1)
    salvar = localizar_elemento(caminhos["salvar"])
    pyautogui.click(salvar)
    time.sleep(2)
    
    if localizar_elemento(caminhos["substituir_arquivo"]):
        pyautogui.press('tab')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)

    print(f"Planilha exportada!")
    
# Função para anexar arquivo
def anexar_planilha(campo_planilha):
    print(f"Anexando planilha...")
    caminho_planilha = filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Arquivos do Excel", "*.xlsx")])
    if caminho_planilha:
        campo_planilha.delete(0, tk.END)
        campo_planilha.insert(0, caminho_planilha)
        print(f"Endereço da planilha anexado no campo")

# Função para anexar imagem
def anexar_imagem(imagem):
    caminho = filedialog.askopenfilename(
        title="Selecione uma imagem",
        filetypes=[("Arquivos de imagem", "*.png;*.jpg;*.jpeg")]
    )
    if caminho:
        imagem.delete(0, tk.END)  # Limpa o conteúdo atual
        imagem.insert(0, caminho)  # Exibe o caminho do arquivo


def registrar_ocorrencias(arquivo_alunos, data, titulo_ocorrencia, descricao_ocorrencia):
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

    # Ler os contatos
    alunos = ler_registros(arquivo_alunos)
    ocorrencias_registradas = 0

    for aluno in alunos:
        try:
            nome_aluno = aluno['Aluno']  # Nome do aluno
            observacao = aluno['Observacao']

            # Verifica se a coluna "Observação" está preenchida
            if not observacao:
                print(f"\nObservação não encontrada para {nome_aluno}. Pulando para o próximo aluno.")
                continue  # Pula para o próximo aluno se houver observação
            
            print("\nTitulo da ocorrência:", titulo_ocorrencia)
            if "<Observação da Planilha>" in descricao_ocorrencia:
                descricao_ocorrencia = observacao
                print("\nDescrição da ocorrência:", descricao_ocorrencia)

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

            botao_salvar = localizar_elemento(caminhos["botao_salvar"])
            pyautogui.click(botao_salvar)
            
            esperar_elemento(caminhos["registro_salvo"])
            ocorrencias_registradas += 1

            pyautogui.press('esc')
            time.sleep(2)

            esperar_elemento(caminhos["contratos"])

            pyautogui.press('esc')
            time.sleep(2)

        except:
            if ocorrencias_registradas == 0:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, não consegui registrar nenhuma ocorrência :(")
            else:
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, só consegui registrar {ocorrencias_registradas} ocorrências :(")

    messagebox.showinfo("Concluído!", "Histórico registrado para todos os alunos!")           

def gerar_planilha_com_educadores(data_falta, filtro_educador):
    abrir_aba("faltas_por_periodo")

    keyboard.write(str(data_falta))
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    keyboard.write(str(data_falta))

    pesquisar_faltosos = localizar_elemento(caminhos["pesquisar"])
    pyautogui.click(pesquisar_faltosos)

    esperar_elemento(caminhos["lista_faltosos"])
    nome_filtro_presencas = localizar_elemento(caminhos["nome_filtro_presencas"])
    pyautogui.rightClick(nome_filtro_presencas)
    esperar_elemento(caminhos["editor_de_filtro"])
    editor_de_filtro = localizar_elemento(caminhos["editor_de_filtro"])
    pyautogui.rightClick(editor_de_filtro)

    esperar_elemento(caminhos["construtor_filtro"])
    if filtro_educador != "Geral":
        nome_filtro_presencas = localizar_elemento(caminhos["nome_filtro_presencas"])
        pyautogui.click(nome_filtro_presencas)
        repetir_tecla("tab", total_repeticoes=5)
        pyautogui.press('enter')
        time.sleep(1)
        valor_filtro = localizar_elemento(caminhos["valor_filtro"])
        pyautogui.click(valor_filtro)
        time.sleep(1)
        keyboard.write(str(filtro_educador))
        time.sleep(1)
        adicionar_filtro = localizar_elemento(caminhos["adicionar_filtro"])
        pyautogui.click(adicionar_filtro)
        time.sleep(1)
        pyautogui.moveTo(valor_filtro)
        time.sleep(1)
    valor_filtro = localizar_elemento(caminhos["valor_filtro"])
    pyautogui.click(valor_filtro)
    time.sleep(1)
    keyboard.write('0')
    time.sleep(1)
    pyautogui.press('enter')
    botao_ok = localizar_elemento(caminhos["botao_ok"])
    pyautogui.click(botao_ok)
    time.sleep(3)

    esperar_elemento(caminhos["lista_faltosos"])
    exportar_faltosos = localizar_elemento(caminhos["exportar"])
    pyautogui.click(exportar_faltosos)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    caminho_destino = Path.home() / "Documents" / "EasyLog" / "alunos_e_educadores.xls"
    exportar_planilha(caminho_destino)

        
def gerar_faltosos_do_dia(campo_data_inicial, campo_data_final, campo_filtro_educador):
    messagebox.showinfo("Atenção!", "Certifique-se de ter feito o login no HUB!")

    procurar_hub()

    data_falta = campo_data_inicial.get()
    data_falta = data_falta.replace("/", "")
    filtro_educador = campo_filtro_educador.get()

    gerar_planilha_com_educadores(data_falta, filtro_educador)
        
    abrir_aba('presencas_e_faltas')
    
    pyautogui.write(data_falta)
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(data_falta)
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('space')
    time.sleep(1)
    visualizar = localizar_elemento(caminhos["visualizar"])
    pyautogui.click(visualizar)
    
    esperar_elemento(caminhos["visu_presencas_e_faltas"])
    aba_arquivo = localizar_elemento(caminhos["aba_arquivo"])
    pyautogui.click(aba_arquivo)
    exportar_documento = localizar_elemento(caminhos["exportar_documento"])
    pyautogui.moveTo(exportar_documento)
    esperar_elemento(caminhos["documento_excel"])
    documento_excel = localizar_elemento(caminhos["documento_excel"])
    pyautogui.click(documento_excel)
    time.sleep(1)
    
    esperar_elemento(caminhos["opcoes_exportacao"])
    pyautogui.press('enter')
    time.sleep(1)
   
    caminho_destino = Path.home() / "Documents" / "EasyLog" / "faltosos_do_dia.xls"
    exportar_planilha(caminho_destino)
        
    esperar_elemento(caminhos["abrir_planilha"])
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    filtrar_faltosos_do_dia(caminho_destino, data_falta)

    data_falta = data_falta.lower()
    caminho_destino = Path.home() / "Documents" / "EasyLog" / f"faltosos_{data_falta}.xlsx"
    os.startfile(caminho_destino)

    messagebox.showinfo("Atenção!", "Planilha de faltosos do dia gerada!")

def gerar_planilha_com_celulares(data_inicial, data_final):
    abrir_aba("presencas_e_faltas")

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
    aba_arquivo = localizar_elemento(caminhos["aba_arquivo"])
    pyautogui.click(aba_arquivo)
    exportar_documento = localizar_elemento(caminhos["exportar_documento"])
    pyautogui.moveTo(exportar_documento)
    esperar_elemento(caminhos["documento_excel"])
    documento_excel = localizar_elemento(caminhos["documento_excel"])
    pyautogui.click(documento_excel)
    time.sleep(1)
    
    esperar_elemento(caminhos["opcoes_exportacao"])
    pyautogui.press('enter')
    time.sleep(1)
   
    caminho_destino = Path.home() / "Documents" / "EasyLog" / "celulares_faltosos_do_mes.xls"
    exportar_planilha(caminho_destino)
        
    esperar_elemento(caminhos["abrir_planilha"])
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    
def gerar_faltosos_do_mes(campo_data_inicial, campo_data_final):
    from datetime import datetime

    messagebox.showinfo("Atenção!", "Certifique-se de ter feito o login no HUB!")
    procurar_hub()

    data_inicial = campo_data_inicial.get()
    data_final = campo_data_final.get()

    data_inicial = data_inicial.replace("/", "")
    data_final = data_final.replace("/", "")
    
    gerar_planilha_com_celulares(data_inicial, data_final)
    
    procurar_hub()

    aba_pedagogico = localizar_elemento(caminhos["aba_pedagogico"])
    pyautogui.click(aba_pedagogico)

    esperar_elemento(caminhos["abrir_faltas_por_periodo"])
    abrir_faltas_por_periodo = localizar_elemento(caminhos["abrir_faltas_por_periodo"])
    pyautogui.click(abrir_faltas_por_periodo)

    esperar_elemento(caminhos["faltas_por_periodo"])
    keyboard.write(str(data_inicial))
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    keyboard.write(str(data_final))

    pesquisar_faltosos = localizar_elemento(caminhos["pesquisar"])
    pyautogui.click(pesquisar_faltosos)

    esperar_elemento(caminhos["lista_faltosos"])
    abrir_menu_presencas = localizar_elemento(caminhos["abrir_menu_presencas"])
    pyautogui.rightClick(abrir_menu_presencas)
    esperar_elemento(caminhos["editor_de_filtro"])
    editor_de_filtro = localizar_elemento(caminhos["editor_de_filtro"])
    pyautogui.click(editor_de_filtro)

    esperar_elemento(caminhos["construtor_filtro"])
    nome_filtro_presencas = localizar_elemento(caminhos["nome_filtro_presencas"])
    pyautogui.click(nome_filtro_presencas)
    repetir_tecla("tab", total_repeticoes=18)
    pyautogui.press('enter')
    time.sleep(1)
    valor_filtro = localizar_elemento(caminhos["valor_filtro"])
    pyautogui.click(valor_filtro)
    time.sleep(1)
    keyboard.write("Normal")
    time.sleep(1)

    adicionar_filtro = localizar_elemento(caminhos["adicionar_filtro"])
    pyautogui.click(adicionar_filtro)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    nome_filtro_presencas = localizar_elemento(caminhos["nome_filtro_presencas"])
    pyautogui.click(nome_filtro_presencas)
    repetir_tecla("tab", total_repeticoes=15)
    pyautogui.press('enter')
    time.sleep(1)
    valor_filtro = localizar_elemento(caminhos["valor_filtro"])
    pyautogui.click(valor_filtro)
    time.sleep(1)
    keyboard.write("Ativo")
    time.sleep(1)

    adicionar_filtro = localizar_elemento(caminhos["adicionar_filtro"])
    pyautogui.click(adicionar_filtro)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    nome_filtro_presencas = localizar_elemento(caminhos["nome_filtro_presencas"])
    pyautogui.click(nome_filtro_presencas)
    repetir_tecla("tab", total_repeticoes=7)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    repetir_tecla("tab", total_repeticoes=3)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    keyboard.write("0")
    time.sleep(1)
    botao_ok = localizar_elemento(caminhos["botao_ok"])
    pyautogui.click(botao_ok)
    time.sleep(3)

    exportar_faltosos = localizar_elemento(caminhos["exportar"])
    pyautogui.click(exportar_faltosos)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    data_atual = datetime.now()
    mes_atual = data_atual.strftime("%B")
    mes_atual = mes_atual.lower()
    caminho_destino = Path.home() / "Documents" / "EasyLog" / f"faltosos_{mes_atual}.xls"
    exportar_planilha(caminho_destino)
    filtrar_faltosos_do_mes(caminho_destino)

    caminho_destino = Path.home() / "Documents" / "EasyLog" / f"faltosos_{mes_atual}.xlsx"
    os.startfile(caminho_destino)
    messagebox.showinfo("Atenção!", f"Planilha dos faltosos de {mes_atual} gerada!")
    
def gerar_frequencia(campo_dia_da_semana, campo_sala):
    messagebox.showinfo("Atenção!", "Certifique-se de ter feito o login no HUB!")
    procurar_hub()

    dia_da_semana = campo_dia_da_semana.get()
    sala = campo_sala.get()

    abrir_aba("controle_pedagogico")

    ocultar_contrato = localizar_elemento(caminhos["ocultar_contrato"])
    pyautogui.click(ocultar_contrato)

    for i in range(0,2):
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('space')
        time.sleep(0.5)

    pyautogui.press('tab')
    time.sleep(0.5)

    for i in range(0,2):
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('space')
        time.sleep(0.5)

    pyautogui.press('tab')
    time.sleep(0.5)

    for i in range(0,3):
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('space')
        time.sleep(0.5)

    repetir_tecla("tab", total_repeticoes=5)
    pyautogui.press('space')
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)

    for i in range(0,13):
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('space')
        time.sleep(0.5)

    abrir_menu_aluno= localizar_elemento(caminhos["abrir_menu_aluno"])
    pyautogui.rightClick(abrir_menu_aluno)
    esperar_elemento(caminhos["editor_de_filtro"])
    editor_de_filtro = localizar_elemento(caminhos["editor_de_filtro"])
    pyautogui.click(editor_de_filtro)

    esperar_elemento(caminhos["construtor_filtro"])
    # ADICIONAR FILTRO DA SALA
    nome_filtro_aluno = localizar_elemento(caminhos["nome_filtro_aluno"])
    pyautogui.click(nome_filtro_aluno)
    esperar_elemento(caminhos["filtrar_por_sala"])
    filtrar_por_sala = localizar_elemento(caminhos["filtrar_por_sala"])
    pyautogui.click(filtrar_por_sala)
    time.sleep(1)
    valor_filtro = localizar_elemento(caminhos["valor_filtro"])
    pyautogui.click(valor_filtro)
    time.sleep(1)
    keyboard.write(sala)
    time.sleep(1)

    # ADICIONAR FILTRO DO DIA
    adicionar_filtro = localizar_elemento(caminhos["adicionar_filtro"])
    pyautogui.click(adicionar_filtro)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    filtrar_por_dias = localizar_elemento(caminhos["filtrar_por_dias"])
    pyautogui.click(filtrar_por_dias)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    contem = localizar_elemento(caminhos["contem"])
    pyautogui.click(contem)
    time.sleep(1)
    keyboard.write(dia_da_semana)
    time.sleep(1)
    botao_ok = localizar_elemento(caminhos["botao_ok"])
    pyautogui.click(botao_ok)
    time.sleep(3)

    # ORDENAR HORÁRIOS
    abrir_menu_horas = localizar_elemento(caminhos["abrir_menu_horas"])
    pyautogui.rightClick(abrir_menu_horas)
    esperar_elemento(caminhos["ordem_ascendente"])
    ordenar_horas= localizar_elemento(caminhos["ordem_ascendente"])
    pyautogui.click(ordenar_horas)
    time.sleep(1)

    # ORDENAR NOMES
    abrir_menu_aluno = localizar_elemento(caminhos["abrir_menu_aluno"])
    pyautogui.rightClick(abrir_menu_aluno)
    esperar_elemento(caminhos["ordem_ascendente"])
    ordenar_nomes= localizar_elemento(caminhos["ordem_ascendente"])
    pyautogui.click(ordenar_nomes)
    time.sleep(1)

    exportar_frequencia = localizar_elemento(caminhos["exportar"])
    pyautogui.click(exportar_frequencia)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    sala = sala.lower()
    dia_da_semana = dia_da_semana.lower()
    caminho_destino = Path.home() / "Documents" / "EasyLog" / f"alunos_{sala}_{dia_da_semana}.xlsx"
    exportar_planilha(caminho_destino)

    os.startfile(caminho_destino)

    messagebox.showinfo("Atenção!", "Planilha para frequência de alunos gerada!")

# Função para enviar uma imagem para os contatos
def enviar_mensagens(arquivo_contatos, caminho_imagem, mensagem_template):
    from scripts.mensagens import personalizar_mensagem

    messagebox.showinfo("Aviso!", "Certifique-se de que o Whatsapp esteja conectado ao aplicativo!")
    
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
    alunos = ler_faltosos_dia(arquivo_contatos)
    mensagens_enviadas = 0
    
    for aluno in alunos:
        try:
            if aluno['Aluno']:
                nome_aluno = str(aluno['Aluno'])  # Nome do aluno
                if aluno['Observacao']:
                    # Verifica se a coluna "Observação" está preenchida
                    observacao = aluno.get('Observacao')  # Usa get para evitar KeyError
                    if pd.notna(observacao):  # Verifica se a observação não é NaN
                        print(f"Observação encontrada para {nome_aluno}. Pulando para o próximo aluno.")
                        continue  # Pula para o próximo aluno se houver observação
                    
                    if aluno['Educador']:
                        nome_educador = str(aluno['Educador'])
                        if len(nome_aluno) > 0 and len(nome_educador) > 0:
                            print("Personalizando mensagem...")
                            mensagem_personalizada = personalizar_mensagem(nome_aluno, nome_educador, mensagem_template)

                        if aluno['Contato']:
                            telefone = aluno['Contato']
                            telefone = telefone[:2] + telefone[3:]  # Remove o índice 2 (que é o terceiro número)

            pyautogui.hotkey('ctrl','n')
            esperar_elemento(caminhos["nova_conversa"])
            
            # Usar o pyautogui para digitar o número de telefone do contato
            pyautogui.write(f'{telefone}')
            time.sleep(2)   
            
            whatsapp_existe = verificar_existencia('pesquisa_contato')

            if whatsapp_existe:
                print("Whatsapp encontrado")
                pyautogui.press('tab')
                pyautogui.press('tab')  
                time.sleep(1)

                pyautogui.press('enter')  
                time.sleep(1)
                
                #Verifica se há imagem
                if len(caminho_imagem) > 0:
                    esperar_elemento(caminhos["anexar"])
                    botao_anexar = localizar_elemento(caminhos["anexar"])
                    pyautogui.click(botao_anexar)

                    esperar_elemento(caminhos["fotos_e_videos"])
                    pyautogui.press('tab')
                    time.sleep(1)
                    pyautogui.press('enter')
                    
                    esperar_elemento(caminhos["janela_anexar"])
                    # Copia o caminho da imagem
                    pyperclip.copy(caminho_imagem)
                    time.sleep(1)
                    # Cola o caminho da imagem
                    pyautogui.hotkey('ctrl', 'v')  # Colar o caminho da imagem no campo
                    time.sleep(2)

                    pyautogui.press('enter')
                    
                    esperar_elemento(caminhos["aba_anexar"])

                    if mensagem_personalizada:
                        pyperclip.copy(mensagem_personalizada)
                        time.sleep(2)
                        # Usar o pyautogui para colar a mensagem
                        pyautogui.hotkey('ctrl','v')
                        time.sleep(1)
                    
                    time.sleep(1)
                    # Pressionar Enter para enviar a imagem
                    pyautogui.press('enter')
                    mensagens_enviadas += 1

                else:
                    pyperclip.copy(mensagem_personalizada)
                    # Usar o pyautogui para colar a mensagem
                    pyautogui.hotkey('ctrl','v')
                    time.sleep(2)

                    # Clicar no botão de alternar
                    pyautogui.press('enter')
                    mensagens_enviadas += 1
                
            else:
                print("Whatsapp não encontrado")
                pyautogui.hotkey('ctrl','a')
                time.sleep(1)

                # Pressionar Enter para enviar a imagem
                pyautogui.press('backspace')
            
            time.sleep(15)  # Aguarde um tempo antes de enviar para o próximo contato
        except Exception as e:
            if mensagens_enviadas == 0:
                print(f"Erro ao enviar mensagens: {e}")
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, não consegui enviar nenhuma mensagem :(")
                return []
            else:
                print(f"Erro ao enviar mensagens: {e}")
                messagebox.showerror("Oops!", f"Desculpe! Devido a um erro, só consegui enviar {mensagens_enviadas} mensagens :(")
                return []
                            
    messagebox.showinfo("Concluído!", "Mensagem enviada para todos os contatos")
