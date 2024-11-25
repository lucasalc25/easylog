import pyautogui
import time
import pytesseract
from PIL import Image
import cv2
import numpy as np

def abrir_whatsapp():
    # Pressionar Windows para abrir a conversa
    pyautogui.press('win')  
    time.sleep(1)
    
    # Usar o pyautogui para digitar whatsapp
    pyautogui.write('whatsapp')
    time.sleep(1)

    # Pressionar Enter para abrir o app
    pyautogui.press('enter')
    time.sleep(1)
    
def verifica_existencia():
    
    # Captura a tela e salva como imagem
    screenshot_path = "print.png"
    pyautogui.screenshot(screenshot_path, region=(64, 149, 335, 110))


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

    # Verificar se há texto na imagem
    if texto_extraido.strip():  # strip() remove espaços extras
        return True
    else:
        return False
        
    
# Função para enviar uma imagem para os contatos
def enviar_mensagens(contatos):
    for contato in contatos:
        #contato = '(92)98422-2186'
        contato = contato.replace("(", "").replace(")", "").replace(" ", "")
        # Remover o terceiro elemento, que é o 9
        contato = contato[:2] + contato[3:]  # Remove o índice 2 (que é o terceiro número)
        
        pyautogui.hotkey('ctrl','f')
        time.sleep(1)
        
        # Usar o pyautogui para digitar o número de telefone do contato
        pyautogui.write(f'{contato}')
        time.sleep(1)   
        
        possui_whatsapp = verifica_existencia()
        time.sleep(1)
        
        if possui_whatsapp:
             # Clica no contato encontrado
            pyautogui.click(244, 191) 
            time.sleep(1)
            
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
            time.sleep(1)

            
            for i in range(1, 5):
                # Clicar no botão de alternar
                pyautogui.press('tab')   
            
            time.sleep(1)

            # Pressionar Enter para enviar a imagem
            pyautogui.press('enter')
            
        else:
            pyautogui.hotkey('ctrl','a')
            time.sleep(1)

            # Pressionar Enter para enviar a imagem
            pyautogui.press('backspace')
        
        time.sleep(2)  # Aguarde um tempo antes de enviar para o próximo contato
            

# Função para ler os contatos de um arquivo TXT
def ler_contatos(arquivo):
    with open(arquivo, "r") as file:
        contatos = file.readlines()
    # Limpar os espaços em branco (como '\n') ao redor dos contatos
    return [contato.strip() for contato in contatos]

# Caminho para o arquivo de contatos e imagem
arquivo_contatos = "contatos.txt"  # Substitua pelo caminho correto

# Ler os contatos
contatos = ler_contatos(arquivo_contatos)
print(f"Contatos encontrados: {contatos}")

abrir_whatsapp()

# Esperar o WhatsApp abrir
time.sleep(2)

# Enviar a imagem para os contatos
enviar_mensagens(contatos)
