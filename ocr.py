import pyautogui
import pytesseract
from PIL import Image
   
def verificar_existencia(elemento):
    if elemento == 'pesquisa_whatsapp':
        coordenadas = (307, 233, 200, 65)
    elif elemento == 'pesquisa_aluno':
        coordenadas = (443,210,737,18)
    
    # Captura a tela e salva como imagem
    caminho_print = "./imagens/print.png"
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
