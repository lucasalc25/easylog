import pyautogui
import pytesseract
from PIL import Image
from PIL import ImageGrab
import cv2
import numpy as np
import time
from config import caminhos

# Configurar Pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   
def verificar_existencia(elemento):
    if elemento == 'pesquisa_contato':
        coordenadas = (307, 233, 200, 65)
    elif elemento == 'pesquisa_aluno':
        coordenadas = (443,210,737,18)
    elif elemento == 'lista_faltosos':
        coordenadas = (375,423,675,82)
    
    try:
        # Captura a tela e salva como imagem
        caminho_print = caminhos["print_verificacao"]
        pyautogui.screenshot(caminho_print, region=coordenadas)

        # Abre a imagem completa
        imagem = Image.open(caminho_print)

        # Usar o pytesseract para extrair o texto da imagem
        texto_extraido = pytesseract.image_to_string(imagem)
        print("Texto extraído da imagem:",texto_extraido)

        # Verificar se há texto na imagem
        if texto_extraido.strip():  # strip() remove espaços extras
            return True
    except Exception as e:
        print(f"Erro no OCR: {e}")

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
        if elemento == caminhos["hub_aberto"]:
            pyautogui.hotkey('alt','tab')
            time.sleep(1)
        
        time.sleep(1)

def encontrar_telefone():
    """
    Captura uma região da tela, detecta e retorna um número de telefone.
    
    Parâmetros:
    - x1, y1, x2, y2: Coordenadas da região da tela a ser analisada.
    
    Retorna:
    - O número de telefone encontrado (string) ou None se não encontrar.
    """
    # Captura a tela e salva como imagem
    caminho_print = caminhos["print_telefones"]
    pyautogui.screenshot(caminho_print, region=(488, 600, 135, 50))
    
    # Abre a imagem completa
    imagem = Image.open(caminho_print)

    # Usar o pytesseract para extrair o texto da imagem
    texto_extraido = pytesseract.image_to_string(imagem)
    print("Texto extraído da imagem:", texto_extraido)

    # Se encontrou telefones, retorna o primeiro corrigido
    if texto_extraido:
        telefone_corrigido = texto_extraido.replace('(02', '(92')  # Corrige erro comum no DDD
        telefone_corrigido = texto_extraido.replace('(62', '(92')  # Corrige erro comum no DDD
        return telefone_corrigido

    return None  # Retorna None se não encontrar telefone
