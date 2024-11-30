import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
import pandas
import time
import pyautogui
import pytesseract
from PIL import Image
import openpyxl
import re

def verifica_existencia():
    
    # Captura a tela e salva como imagem
    screenshot_path = "print.png"
    pyautogui.screenshot(screenshot_path, region=(307, 233, 200, 65))


    # Use o Tesseract para extrair texto da imagem
    # Certifique-se de que o Tesseract esteja instalado e configurado no PATH do sistema
    # Para Windows, configure o caminho abaixo se necessário:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    # Abre a imagem completa
    imagem = Image.open(screenshot_path)

    # Use o Tesseract para extrair texto da imagem
    # Certifique-se de que o Tesseract esteja instalado e configurado no PATH do sistema
    # Para Windows, configure o caminho abaixo se necessário:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    # Usar o pytesseract para extrair o texto da imagem
    texto_extraido = pytesseract.image_to_string(imagem)
    print(texto_extraido.strip())

    # Verificar se há texto na imagem
    if texto_extraido.strip():  # strip() remove espaços extras
        return True
    else:
        return False
        
    
# Função para enviar uma imagem para os contatos
def enviar_mensagens(contatos, mensagem, imagem):
    print(mensagem)
    for contato in contatos:
        #contato = '(92)98422-2186'
        contato = contato.replace("(", "").replace(")", "").replace(" ", "")
        # Remover o terceiro elemento, que é o 9
        contato = contato[:2] + contato[3:]  # Remove o índice 2 (que é o terceiro número)
        
        pyautogui.hotkey('ctrl','n')
        time.sleep(1)
        
        # Usar o pyautogui para digitar o número de telefone do contato
        pyautogui.write(f'{contato}')
        time.sleep(1)   
        
        possui_whatsapp = verifica_existencia()
        time.sleep(1)
        
        if possui_whatsapp:
            # Alternar para abrir conversa com contato
            pyautogui.press('tab')
            pyautogui.press('tab')
            time.sleep(1)

            pyautogui.press('enter')
            time.sleep(1)     
            
            #Verifica se há imagem
            if imagem:
                # Clica no botão anexar
                pyautogui.click(497, 693)
                time.sleep(1)
                
                # Alternar para abrir fotos
                pyautogui.press('tab')
                time.sleep(1) 
                pyautogui.press('enter')
                time.sleep(1)          

                # Clicar para anexar a imagem
                pyautogui.write(r'C:\Users\Suporte\Documents\GitHub\bot-whatsapp\imagem.png')  # Caminho completo da imagem
                time.sleep(1)
                
                pyautogui.press('enter')
                time.sleep(1.5)

                if mensagem:
                    # Usar o pyautogui para digitar a mensagem
                    for frase in mensagem:
                        pyautogui.write(f'{frase}')
                        time.sleep(3)
                    
                        for i in range(1, 5):
                            # Clicar no botão de alternar
                            pyautogui.press('tab')
                        
                        time.sleep(1) 
                
                # Clicar no botão de alternar
                pyautogui.press('tab') 
                time.sleep(1)

                # Pressionar Enter para enviar a imagem
                pyautogui.press('enter')

            else:
                # Usar o pyautogui para digitar o número de telefone do contato
                pyautogui.write(f'{mensagem}')
                time.sleep(6)

                # Clicar no botão de alternar
                pyautogui.press('tab')

                # Clicar no botão de alternar
                pyautogui.press('enter')  
            
        else:
            pyautogui.hotkey('ctrl','a')
            time.sleep(1)

            # Pressionar Enter para enviar a imagem
            pyautogui.press('backspace')
        
        time.sleep(2)  # Aguarde um tempo antes de enviar para o próximo contato

    messagebox.showinfo("Concluído!", "Mensagem enviada para todos os contatos")
            
# Função para ler os contatos de um arquivo TXT
def ler_contatos(arquivo):
    with open(arquivo, "r", encoding="utf-8") as file:
        contatos = file.readlines()
        
    # Limpar os espaços em branco (como '\n') ao redor dos contatos
    return [contato.strip() for contato in contatos]

# Função para iniciar envio
def preparar_envio(campo_planilha, campo_mensagem, campo_imagem):
    caminho_contatos = campo_planilha.get()
    mensagem = campo_mensagem.get("1.0", ttk.END)  # Captura o texto do campo
    imagem = campo_imagem.get()

    # Caminho para o arquivo de contatos e imagem
    arquivo_contatos = caminho_contatos  # Substitua pelo caminho correto

    # Ler os contatos
    contatos = ler_contatos(arquivo_contatos)
    
    # Pressionar Windows para abrir a conversa
    pyautogui.press('win')  
    time.sleep(1)
    
    # Usar o pyautogui para digitar whatsapp
    pyautogui.write('whatsapp')
    time.sleep(1)

    # Pressionar Enter para abrir o app
    pyautogui.press('enter')
    time.sleep(3)

    enviar_mensagens(contatos, mensagem, imagem)

# Função para centralizar a janela
def centralizar_janela(window):
    window.update_idletasks()
    largura_tela = window.winfo_screenwidth()
    altura_tela = window.winfo_screenheight()
    largura_janela = window.winfo_width()
    altura_janela = window.winfo_height()
    x = (largura_tela // 2) - (largura_janela // 2)
    y = (altura_tela // 2) - (altura_janela // 2)
    window.geometry(f"+{x}+{y}")

# Tela Inicial
def exibir_tela_inicial():
    root = ttk.Window(themename="cosmo")
    root.title("Pedagobot")
    root.iconbitmap("./icone.ico")

    # Funções para redirecionar para cada aba
    def abrir_faltosos():
        root.withdraw()  # Esconde a janela de boas-vindas
        exibir_faltosos(root)

    def abrir_comunicados():
        root.withdraw()  # Esconde a janela de boas-vindas
        exibir_comunicados(root)
    
    # Função para exibir a funcionalidade Históricos
    def abrir_historicos():
        messagebox.showinfo("Históricos", "Funcionalidade em desenvolvimento.")

    # Função para exibir a funcionalidade Planilhas
    def abrir_planilhas():
        messagebox.showinfo("Planilhas", "Funcionalidade em desenvolvimento.")

    # Layout da tela inicial
    frame_principal = ttk.Frame(root, padding=20)
    frame_principal.pack(fill=ttk.BOTH, expand=True)

    ttk.Label(
        frame_principal,
        text="Bem-vindo(a) ao Pedagobot",
        font=("Helvetica", 16, "bold")
    ).pack(pady=10)         
        
    ttk.Label(
        frame_principal,
        text="Deixando seu trabalho diário mais rápido e leve",
        font=("Helvetica", 11,)
    ).pack(pady=(10, 20))

    # Botões para cada funcionalidade
    ttk.Button(
        frame_principal, text="Faltosos", 
        command=abrir_faltosos, 
        bootstyle="primary-outline", 
        width=20
    ).pack(pady=10)

    ttk.Button(
        frame_principal, text="Comunicados", 
        command=abrir_comunicados, 
        bootstyle="success-outline", 
        width=20
    ).pack(pady=10) 

    ttk.Button(
        frame_principal, text="Históricos", 
        command=abrir_historicos, 
        bootstyle="info-outline", 
        width=20
    ).pack(pady=10)

    ttk.Button(
        frame_principal, text="Planilhas", 
        command=abrir_planilhas, 
        bootstyle="warning-outline", 
        width=20
    ).pack(pady=10)

    centralizar_janela(root)
    root.mainloop()
    
def ler_contatos(caminho_planilha):
    try:
        wb = openpyxl.load_workbook(caminho_planilha)
        sheet = wb.active

        contatos = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            nome, telefone = row
            contatos.append({
                "nome": nome.strip() if nome else "",
                "telefone": sanitizar_telefone(telefone)
            })
        
        return contatos
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return []

def sanitizar_telefone(telefone):
    if telefone:
        return re.sub(r'\D', '', telefone)  # Remove tudo que não for número
    return ""

# Função para anexar arquivo
def anexar_planilha(txtBox_planilha):
    caminho_planilha = filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Excel files", "*.xlsx")])
    if caminho_planilha:
        txtBox_planilha.delete(0, ttk.END)
        txtBox_planilha.insert(0, caminho_planilha)
        contatos = ler_contatos(caminho_planilha)

        if contatos:
            mostrar_contatos(contatos)
            
def mostrar_contatos(contatos):
    print("Contatos Carregados")
    
    for contato in enumerate(contatos):
        print(f"{contato['nome']} - {contato['telefone']}")


# Função para anexar imagem
def anexar_imagem(imagem):
    caminho = filedialog.askopenfilename(title="Selecione uma imagem", filetypes=[("Arquivos de imagem", "*.png;*.jpg;*.jpeg")]
    )
    if caminho:
        imagem.delete(0, ttk.END)  # Limpa o conteúdo atual
        imagem.insert(0, caminho)  # Exibe o caminho do arquivo

def reexibir_tela_inicial(window, tela_inicial):
    window.withdraw()
    tela_inicial.deiconify()

# Tela Faltosos
def exibir_faltosos(tela_inicial):
    root = ttk.Window(themename="cosmo")
    root.title("Faltosos")
    root.iconbitmap("./icone.ico")

    # Layout da aba Faltosos
    frame_principal = ttk.Frame(root, padding=20)
    frame_principal.pack(fill=ttk.BOTH, expand=True)

    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame_principal, text="Planilha de contatos:", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=ttk.X, pady=5)
    entry_planilha = ttk.Entry(frame_contatos)
    entry_planilha.pack(side=ttk.LEFT, fill=ttk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:anexar_planilha(entry_planilha), bootstyle="success").pack(side=ttk.RIGHT, padx=5)

    # Campos para dados adicionais
    frame_dados = ttk.Labelframe(frame_principal, text="Informações Adicionais:", padding=5, bootstyle="primary")
    frame_dados.pack(fill=ttk.X, pady=5)

    ttk.Label(frame_dados, text="Professor:").pack(side=ttk.LEFT, padx=5)
    entry_professor = ttk.Entry(frame_dados, width=20)
    entry_professor.pack(side=ttk.LEFT, padx=5)

    ttk.Label(frame_dados, text="Dia da Falta:").pack(side=ttk.LEFT, padx=5)
    entry_dia_falta = ttk.Entry(frame_dados, width=15)
    entry_dia_falta.pack(side=ttk.LEFT, padx=5)

    # Mensagem padrão
    mensagem_padrao = """Olá, tudo bem? Aqui é o professor {nome_professor} da Microlins.

Verifiquei que {nome_aluno} não compareceu à aula do dia {data}. Para que isso não gere atrasos no contrato, você precisa nos informar o motivo desta falta para registro em nosso sistema, tá bom? Depois é só marcar sua reposição.

Fico no seu aguardo. 
Obrigado desde já!"""

     # Campo para digitação do modelo de mensagem
    frame_texto = ttk.Labelframe(frame_principal, text="Modelo da mensagem (opcional):", padding=5, bootstyle="primary")
    frame_texto.pack(fill=ttk.BOTH, pady=5, expand=True)
    text_mensagem = ttk.Text(frame_texto, height=7, wrap="word")
    text_mensagem.insert("1.0", mensagem_padrao)
    text_mensagem.pack(fill=ttk.BOTH, padx=5, expand=True)

    # Campo para anexar imagem
    frame_imagem = ttk.Labelframe(frame_principal, text="Imagem (opcional):", padding=5, bootstyle="primary")
    frame_imagem.pack(fill=ttk.X, pady=5)
    entry_imagem = ttk.Entry(frame_imagem)
    entry_imagem.pack(side=ttk.LEFT, fill=ttk.X, expand=True, padx=5)
    ttk.Button(frame_imagem, text="Anexar", command=lambda:anexar_imagem(entry_imagem)).pack(side=ttk.RIGHT, padx=5)

    # Botão para enviar mensagens
    ttk.Button(frame_principal, text="Enviar", command=lambda:preparar_envio(entry_planilha, text_mensagem, entry_imagem), bootstyle="success-outline").pack(pady=10)

    # Botão para voltar
    ttk.Button(
        frame_principal, text="Voltar", 
        command=lambda: (reexibir_tela_inicial(root, tela_inicial)), 
        bootstyle="danger-outline"
    ).pack(pady=10)

    centralizar_janela(root)
    root.mainloop()

# Tela Comunicados
def exibir_comunicados(tela_inicial):
    root = ttk.Window(themename="cosmo")
    root.title("Comunicados")
    root.iconbitmap("./icone.ico")

    # Layout da aba Comunicados
    frame_principal = ttk.Frame(root, padding=20)
    frame_principal.pack(fill=ttk.BOTH, expand=True)

    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame_principal, text="Planilha de contatos:", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=ttk.X, pady=5)
    entry_planilha = ttk.Entry(frame_contatos)
    entry_planilha.pack(side=ttk.LEFT, fill=ttk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:anexar_planilha(entry_planilha), bootstyle="success").pack(side=ttk.RIGHT, padx=5)
    
    # Campo para mensagem
    ttk.Label(frame_principal, text="Modelo de mensagem:").pack(anchor="w", pady=5)
    text_mensagem = ttk.Text(frame_principal, height=8)
    text_mensagem.pack(fill=ttk.BOTH, pady=5)

    # Botão para enviar mensagens
    ttk.Button(frame_principal, text="Enviar", bootstyle="success-outline").pack(pady=10)

    # Botão para voltar
    ttk.Button(                                                                                                                                          
        frame_principal, text="Voltar", 
        command=lambda: (reexibir_tela_inicial(root, tela_inicial)), 
        bootstyle="danger-outline"
    ).pack(pady=10)

    centralizar_janela(root)
    root.mainloop()

# Função para encerrar o programa ao fechar a janela inicial
def encerrar_programa(root):
    if messagebox.askokcancel("Sair", "Você deseja realmente encerrar o programa?"):
        root.destroy()  # Encerra o programa principal

# Iniciar a aplicação
exibir_tela_inicial()
