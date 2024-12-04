import pyautogui
import pytesseract
from PIL import Image
   
def verificar_existencia(elemento):
    if elemento == 'pesquisa_whatsapp':
        coordenadas = (307, 233, 200, 65)
    elif elemento == 'pesquisa_aluno':
        coordenadas = (443,210,737,18)
    elif elemento == 'lista_faltosos':
        coordenadas = (346,472,675,82)
    
    # Captura a tela e salva como imagem
    caminho_print = "./imagens/print_verificaçao.png"
    pyautogui.screenshot(caminho_print, region=coordenadas)


    # Use o Tesseract para extrair texto da imagem
    # Certifique-se de que o Tesseract esteja instalado e configurado no PATH do sistema
    # Para Windows, configure o caminho abaixo se necessário:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    # Abre a imagem completa
    imagem = Image.open(caminho_print)

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
    
import cv2
import numpy as np
import pyautogui
import pytesseract
import time
import keyboard

# Configurar Pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def localizar_elemento(caminho_imagem, threshold=0.8):
    print = pyautogui.screenshot()
    frame = np.array(print)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # Carregar o template
    template = cv2.imread(caminho_imagem, 0)

    # Localizar o template na tela
    result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)

    if max_val >= threshold:
        h, w = template.shape
        center_x = max_loc[0] + w // 2
        center_y = max_loc[1] + h // 2
        return (center_x, center_y)
    
    return False

def esperar_elemento(elemento):
    while not localizar_elemento(elemento):
        if elemento == './imagens/hub_aberto.png':
            pyautogui.hotkey('alt','tab')
            time.sleep(1)
        
        time.sleep(1)
