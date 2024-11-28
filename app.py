import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
import subprocess
import time
import pytesseract
from PIL import Image
import pyautogui

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
                    # Usar o pyautogui para digitar o número de telefone do contato
                    pyautogui.write(f'{mensagem}')
                    time.sleep(3)
                
                for i in range(1, 5):
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
    with open(arquivo, "r") as file:
        contatos = file.readlines()
    # Limpar os espaços em branco (como '\n') ao redor dos contatos
    return [contato.strip() for contato in contatos]

# Função para anexar contatos
def anexar_contatos():
    caminho = filedialog.askopenfilename(
        title="Selecione um arquivo de contatos",
        filetypes=[("Arquivos de texto", "*.txt")]
    )
    if caminho:
        entry_contatos.delete(0, ttk.END)  # Limpa o conteúdo atual
        entry_contatos.insert(0, caminho)  # Exibe o caminho do arquivo

# Função para anexar imagem
def anexar_imagem():
    caminho = filedialog.askopenfilename(
        title="Selecione uma imagem",
        filetypes=[("Arquivos de imagem", "*.png;*.jpg;*.jpeg")]
    )
    if caminho:
        entry_imagem.delete(0, ttk.END)  # Limpa o conteúdo atual
        entry_imagem.insert(0, caminho)  # Exibe o caminho do arquivo

# Função para iniciar envio
def preparar_envio():
    caminho_contatos = entry_contatos.get()
    mensagem = text_mensagem.get("1.0", ttk.END)  # Captura o texto do campo
    imagem = entry_imagem.get()

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

# Centralizar a janela
def centralizar_janela(window):
    window.update_idletasks()  # Atualiza informações da janela
    largura_tela = window.winfo_screenwidth()
    altura_tela = window.winfo_screenheight()
    largura_janela = window.winfo_width()
    altura_janela = window.winfo_height()
    x = (largura_tela // 2) - (largura_janela // 2)
    y = (altura_tela // 2) - (altura_janela // 2)
    # Define a posição da janela
    window.geometry(f"+{x}+{y}")

# Criar janela principal
root = ttk.Window(themename="cosmo")
root.title("Pedagobot")
root.geometry("450x450")
# Adicionar o ícone (.ico)
root.iconbitmap("./icone.ico")

# Menu superior
menu = ttk.Menu(root)
root.config(menu=menu)

# Abas do menu
menu_mensagens = ttk.Menu(menu, tearoff=0)
menu_mensagens.add_command(label="Mensagens", state=ttk.ACTIVE)
menu.add_cascade(label="Mensagens", menu=menu_mensagens)

menu_historicos = ttk.Menu(menu, tearoff=0)
menu_historicos.add_command(label="Históricos", state=ttk.DISABLED)
menu.add_cascade(label="Históricos", menu=menu_historicos)

menu_coletar = ttk.Menu(menu, tearoff=0)
menu_coletar.add_command(label="Coletar Dados", state=ttk.DISABLED)
menu.add_cascade(label="Coletar Dados", menu=menu_coletar)

# Frame principal
frame_principal = ttk.Frame(root, padding=10)
frame_principal.pack(fill=ttk.BOTH, expand=True)

# Campo para anexação de arquivo txt
frame_contatos = ttk.Labelframe(frame_principal, text="Arquivo de contatos (TXT):", padding=5, bootstyle="primary")
frame_contatos.pack(fill=ttk.X, pady=5)
entry_contatos = ttk.Entry(frame_contatos)
entry_contatos.pack(side=ttk.LEFT, fill=ttk.X, expand=True, padx=5)
ttk.Button(frame_contatos, text="Anexar", command=lambda: anexar_contatos()).pack(side=ttk.RIGHT, padx=5)

# Campo para digitação do modelo de mensagem
frame_texto = ttk.Labelframe(frame_principal, text="Modelo da mensagem:", padding=5, bootstyle="primary")
frame_texto.pack(fill=ttk.BOTH, pady=5, expand=True)
text_mensagem = ttk.Text(frame_texto, height=7, wrap="word")  # Define altura para 7 linhas
text_mensagem.pack(fill=ttk.BOTH, padx=5, expand=True)

# Campo para anexar imagem
frame_imagem = ttk.Labelframe(frame_principal, text="Imagem (opcional):", padding=5, bootstyle="primary")
frame_imagem.pack(fill=ttk.X, pady=5)
entry_imagem = ttk.Entry(frame_imagem)
entry_imagem.pack(side=ttk.LEFT, fill=ttk.X, expand=True, padx=5)
ttk.Button(frame_imagem, text="Anexar", command=anexar_imagem).pack(side=ttk.RIGHT, padx=5)

# Botão para início do envio
frame_enviar = ttk.Frame(frame_principal)
frame_enviar.pack(fill=ttk.X, pady=20)
ttk.Button(frame_enviar, text="Enviar Mensagens", command=preparar_envio, bootstyle="success").pack()

# Configurações finais e centralização
root.update()  # Garante que todos os widgets estão renderizados
centralizar_janela(root)

# Loop principal
root.mainloop()
