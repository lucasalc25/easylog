import tkinter as tk  # Importando tkinter para acesso a constantes
from ttkbootstrap import Window, ttk  # Importando ttkbootstrap para personalização do tema
from datetime import datetime
import bot
from scripts.historicos import gerar_ocorrencia, preparar_registros
from scripts.mensagens import gerar_mensagem, preparar_envio
from scripts.planilhas import preparar_data_faltosos
import os
import sys

data_atual = datetime.now().strftime("%d/%m/%Y")

def resource_path(relative_path):
    """Obtém o caminho absoluto do recurso, ajustado para executáveis criados com PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        # Para executáveis
        return os.path.join(sys._MEIPASS, relative_path)
    # Para execução direta no código-fonte
    return os.path.join(os.path.abspath("."), relative_path)

# Importar utilitário da pasta scripts
sys.path.append(resource_path('scripts'))  # Adicionar a subpasta 'scripts' ao sys.path

def carregar_recurso(tipo, nome_arquivo):
    """Carrega o recurso com base no tipo e nome."""
    subpasta = {
        'imagem': 'imagens',
        'script': 'scripts',
        'icone': 'imagens'
    }.get(tipo, '')
    return resource_path(os.path.join(subpasta, nome_arquivo))

caminho_icone = carregar_recurso('icone', 'icone.ico')
caminho_aba_anexar = carregar_recurso('imagem', 'aba_anexar.png')
caminho_abrir_planilha = carregar_recurso('imagem', 'abrir_planilha.png')
caminho_aluno_encontrado = carregar_recurso('imagem', 'aluno_encontrado.png')
caminho_anexar = carregar_recurso('imagem', 'anexar.png')
caminho_hub_aberto = carregar_recurso('imagem', 'hub_aberto.png')
caminho_caixa_mensagem = carregar_recurso('imagem', 'caixa_mensagem.png')
caminho_campo_nome_planilha = carregar_recurso('imagem', 'campo_nome_planilha.png')
caminho_campo_pesquisa = carregar_recurso('imagem', 'campo_pesquisa.png')
caminho_contato_inexistente = carregar_recurso('imagem', 'contato_inexistente.png')
caminho_contato = carregar_recurso('imagem', 'contato.png')
caminho_contrato_aberto = carregar_recurso('imagem', 'contrato_aberto.png')
caminho_descricao = carregar_recurso('imagem', 'descricao.png')
caminho_exportar = carregar_recurso('imagem', 'exportar.png')
caminho_faltas_por_periodo = carregar_recurso('imagem', 'faltas_por_periodo.png')
caminho_fotos_e_videos = carregar_recurso('imagem', 'fotos_e_videos.png')
caminho_imagem = carregar_recurso('imagem', 'imagem.png')
caminho_lista_faltosos = carregar_recurso('imagem', 'lista_faltosos.png')
caminho_logo = carregar_recurso('imagem', 'logo.png')
caminho_menu_iniciar = carregar_recurso('imagem', 'menu_iniciar.png')
caminho_nova_conversa = carregar_recurso('imagem', 'nova_conversa.png')
caminho_ocorrencia = carregar_recurso('imagem', 'ocorrencia.png')
caminho_opcoes_exportacao = carregar_recurso('imagem', 'opcoes_exportacao.png')
caminho_pesquisa_aluno = carregar_recurso('imagem', 'pesquisa_aluno.png')
caminho_pesquisa_contato = carregar_recurso('imagem', 'pesquisa_contato.png')
caminho_pesquisar = carregar_recurso('imagem', 'pesquisar.png')
caminho_presencas_e_faltas = carregar_recurso('imagem', 'presencas_e_faltas.png')
caminho_salvar = carregar_recurso('imagem', 'salvar.png')
caminho_substituir_arquivo = carregar_recurso('imagem', 'substituir_arquivo.png')
caminho_visu_presencas_e_faltas = carregar_recurso('imagem', 'visu_presencas_e_faltas.png')
caminho_visualizar = carregar_recurso('imagem', 'visualizar.png')
caminho_whatsapp_aberto = carregar_recurso('imagem', 'whatsapp_aberto.png')
caminho_whatsapp_encontrado = carregar_recurso('imagem', 'whatsapp_encontrado.png')

caminho_bot = carregar_recurso('script', 'bot.py')
caminho_mensagens = carregar_recurso('script', 'mensagens.py')
caminho_historicos = carregar_recurso('script', 'historicos.py')
caminho_planilhas = carregar_recurso('script', 'planilhas.py')
caminho_ocr = carregar_recurso('script', 'ocr.py')
caminho___init__ = carregar_recurso('script', '__init__.py')


# Função para exibir a tela inicial
def exibir_janela_inicial():
    root = Window(themename="cosmo")
    root.title("EasyLog")
    root.iconbitmap(caminho_icone)
    root.resizable(False, False)

    frame_principal = ttk.Frame(root, padding=(70,20))
    frame_principal.pack(fill=tk.BOTH, expand=True)

    # Título e subtítulo
    ttk.Label(frame_principal, text="Bem-vindo(a) ao EasyLog", font=("Helvetica", 16, "bold")).pack(pady=10)
    ttk.Label(frame_principal, text="Tornando seu trabalho mais eficiente", font=("Helvetica", 11)).pack(pady=(10, 20))

    # Botões

    ttk.Button(frame_principal, text="MENSAGENS", command=lambda:abrir_janela(root, "Mensagens"), bootstyle="success-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="HISTÓRICOS",command=lambda:abrir_janela(root, "Históricos"), bootstyle="info-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="PLANILHAS", command=lambda:abrir_janela(root, "Planilhas"), bootstyle="warning-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="SUPORTE", command=lambda:bot.messagebox.showinfo("Aviso", "Funcionalidade em desenvolvimento"), bootstyle="danger-outline", width=20).pack(pady=10)

    # Rodapé com versão
    ttk.Label(frame_principal, text="Versão 1.0", font=("Helvetica", 9)).pack(pady=(20, 0))

    centralizar_janela(root)
    root.mainloop()
    
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

# Função para exibir a janela dinâmica
def abrir_janela(janela_inicial, titulo):
    janela = tk.Toplevel(janela_inicial)
    janela.title(titulo)
    janela.iconbitmap("./imagens/icone.ico")
    janela.resizable(False, False)
    janela.transient(janela_inicial)  # Faz a janela secundária ficar vinculada à principal
    janela.grab_set()  # Bloqueia interações com a janela principal

    frame = ttk.Frame(janela, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # Configurar conteúdo da área central com base no título
    if titulo == "Mensagens":
        frame_mensagens(janela, frame)
    elif titulo == "Históricos":
        frame_historicos(janela,frame)
    elif titulo == "Planilhas":
        frame_planilhas(janela,frame)
    else:
        ttk.Label(frame, text="Conteúdo não configurado.", font=("Helvetica", 12)).pack(pady=10)

# Função para configurar a área de "Mensagens"
def frame_mensagens(janela,frame):
    janela.geometry("500x570")
    centralizar_janela(janela)
    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame, text=" Planilha de contatos: * ", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=tk.X, pady=5)
    campo_planilha = ttk.Entry(frame_contatos)
    campo_planilha.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:bot.anexar_planilha(campo_planilha), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando o frame para o tipo de comunicado
    frame_tipo_mensagem = ttk.Labelframe(frame, text=" Tipo de Mensagem * ", padding=5)
    frame_tipo_mensagem.pack(fill=tk.X, padx=5, pady=5)

    # Variável para armazenar o tipo de comunicado selecionado
    tipo_mensagem_var = tk.StringVar(value="falta")

    # Definindo as opções dos RadioButtons
    opcoes = [
        ("Falta", "falta"),
        ("Multirão", "multirão"),
        ("Reunião de Pais", "reuniao_de_pais"),
        ("Oficina", "oficina"),
        ("Formatura", "formatura"),
        ("Feriado", "feriado")
    ]

    # Colocando os RadioButtons em uma grade
    for index, (texto, valor) in enumerate(opcoes):
        row = index // 3  # Calcula em qual linha deve colocar
        column = index % 3  # Calcula a coluna (de 0 a 2)
        ttk.Radiobutton(frame_tipo_mensagem, text=texto, value=valor, variable=tipo_mensagem_var).grid(row=row, column=column, sticky="w", padx=30, pady=5)

    # Campos para dados adicionais
    frame_variaveis = ttk.Labelframe(frame, text=" Variáveis ", padding=5, bootstyle="primary")
    frame_variaveis.pack(fill=tk.X, pady=5)
    ttk.Label(frame_variaveis, text="Data: ").pack(side=tk.LEFT, padx=(5,0))
    campo_data = ttk.Entry(frame_variaveis, width=12)
    campo_data.pack(side=tk.LEFT, padx=5)
    campo_data.insert(0, data_atual)
    ttk.Label(frame_variaveis, text="Oficina: ").pack(side=tk.LEFT, padx=(5,0))
    campo_oficina = ttk.Entry(frame_variaveis, width=28)
    campo_oficina.pack(side=tk.LEFT, padx=5)
    ttk.Button(frame_variaveis, text="Gerar", command=lambda:gerar_mensagem(tipo_mensagem_var, campo_mensagem, campo_data, campo_oficina)).pack(side=tk.RIGHT, padx=(5,0))

    # Campo para digitação do modelo de mensagem
    frame_texto = ttk.Labelframe(frame, text=" Modelo da mensagem: ", padding=5, bootstyle="primary")
    frame_texto.pack(fill=tk.BOTH, pady=5)
    campo_mensagem = tk.Text(frame_texto, height=7, wrap="word")
    campo_mensagem.pack(fill=tk.BOTH, padx=5)

    # Campo para anexar imagem
    frame_imagem = ttk.Labelframe(frame, text=" Imagem: ", padding=5, bootstyle="primary")
    frame_imagem.pack(fill=tk.X, pady=5)
    campo_imagem = ttk.Entry(frame_imagem)
    campo_imagem.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_imagem, text="Anexar", command=lambda:bot.anexar_imagem(campo_imagem), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando um frame para organizar os botões
    frame_botoes = ttk.Frame(frame)
    frame_botoes.pack(pady=(25,0))

    # Adicionando botões lado a lado usando grid()
    ttk.Button(frame_botoes, text="Enviar", command=lambda:preparar_envio(campo_planilha, campo_mensagem, campo_imagem),bootstyle="success-outline", width=10).grid(row=0, column=0, padx=40)
    ttk.Button(frame_botoes, text="Voltar", command=janela.destroy, bootstyle="danger-outline", width=10).grid(row=0, column=1, padx=40)

# Função para configurar a área de "Históricos"
def frame_historicos(janela, frame):
    janela.geometry("500x540")
    centralizar_janela(janela)

    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame, text=" Planilha de alunos: * ", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=tk.X, pady=5)
    campo_planilha = ttk.Entry(frame_contatos)
    campo_planilha.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:bot.anexar_planilha(campo_planilha), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando o frame para o tipo de comunicado
    frame_comunicado = ttk.Labelframe(frame, text=" Tipo de Ocorrência * ", padding=5, bootstyle="primary")
    frame_comunicado.pack(fill=tk.X, padx=5, pady=5)

    # Variável para armazenar o tipo de comunicado selecionado
    tipo_ocorrencia_var = tk.StringVar(value="falta")

    # Definindo as opções dos RadioButtons
    opcoes = [
        ("Falta", "falta"),
        ("Multirão", "multirao"),
        ("Comportamento", "comportamento"),
        ("Prova", "prova"),
        ("Atividades", "atividades"),
        ("1° dia de aula", "1_dia_de_aula"),
        ("Plantão", "plantao"),

    ]

    # Colocando os RadioButtons em uma grade
    for index, (texto, valor) in enumerate(opcoes):
        row = index // 4  # Calcula em qual linha deve colocar
        column = index % 4  # Calcula a coluna (de 0 a 2)
        ttk.Radiobutton(frame_comunicado, text=texto, value=valor, variable=tipo_ocorrencia_var).grid(row=row, column=column, sticky="w", padx=15, pady=5)

    # Botão para aplicar a seleção e inserir a mensagem
    ttk.Button(frame, text="Gerar Ocorrência", command=lambda:gerar_ocorrencia(tipo_ocorrencia_var, campo_titulo, campo_descricao)).pack(pady=10)

    # Campo para digitação do título da ocorrência
    frame_texto_tipo = ttk.Labelframe(frame, text=" Título: * ", padding=5, bootstyle="primary")
    frame_texto_tipo.pack(fill=tk.BOTH, pady=5)
    campo_titulo = tk.Text(frame_texto_tipo, height=1, wrap="word")
    campo_titulo.pack(fill=tk.BOTH, padx=5)

    # Campo para digitação da descrição da ocorrência
    frame_texto_descricao = ttk.Labelframe(frame, text=" Descrição: * ", padding=5, bootstyle="primary")
    frame_texto_descricao.pack(fill=tk.BOTH, pady=5)
    campo_descricao = tk.Text(frame_texto_descricao, height=7, wrap="word")
    campo_descricao.pack(fill=tk.BOTH, padx=5)

    # Criando um frame para organizar os botões
    frame_botoes = ttk.Frame(frame)
    frame_botoes.pack(pady=(25,0))

    # Adicionando botões lado a lado usando grid()
    ttk.Button(frame_botoes, text="Registrar", command=lambda:preparar_registros(campo_planilha, campo_titulo, campo_descricao),bootstyle="success-outline", width=10).grid(row=0, column=0, padx=40)
    ttk.Button(frame_botoes, text="Voltar", command=janela.destroy, bootstyle="danger-outline", width=10).grid(row=0, column=1, padx=40)


# Função para configurar a área de "Planilhas"
def frame_planilhas(janela, frame):
    janela.geometry("500x500")
    centralizar_janela(janela)
    
    # Campo para anexação de planilha
    frame_planilhas = ttk.Labelframe(frame, text=" Faltas ", padding=5, bootstyle="primary")
    frame_planilhas.pack(fill=tk.X, pady=5)
    ttk.Label(frame_planilhas, text="Data Inicial: * ").pack(side=tk.LEFT, padx=(5,0))
    campo_data_inicial = ttk.Entry(frame_planilhas, width=12)
    campo_data_inicial.pack(side=tk.LEFT, padx=5)
    campo_data_inicial.insert(0, data_atual)
    ttk.Label(frame_planilhas, text="Data Final: * ").pack(side=tk.LEFT, padx=(5,0))
    campo_data_final = ttk.Entry(frame_planilhas, width=12)
    campo_data_final.pack(side=tk.LEFT, padx=5)
    campo_data_final.insert(0, data_atual)

    ttk.Button(frame_planilhas, text="Gerar", command=lambda:preparar_data_faltosos(campo_data_inicial, campo_data_final), bootstyle="primary").pack(side=tk.RIGHT, padx=5)

    ttk.Button(frame, text="Voltar", command=janela.destroy, bootstyle="danger-outline", width=10).pack(pady=10)

# Execução do programa
if __name__ == "__main__":
    bot.criar_pastas()
    exibir_janela_inicial()
