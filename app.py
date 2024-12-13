import tkinter as tk  # Importando tkinter para acesso a constantes
from tkinter import messagebox
from ttkbootstrap import Window, ttk  # Importando ttkbootstrap para personalização do tema
from datetime import datetime, timedelta
from config import caminhos
from bot import anexar_imagem, anexar_planilha, criar_pastas
from scripts.historicos import gerar_ocorrencia, preparar_registros
from scripts.mensagens import gerar_mensagem, preparar_envio
from scripts.planilhas import preparar_alunos_atencao, preparar_data_faltosos

# Obter a data atual
data_atual = datetime.now()

# Subtrair um dia da data atual
data_anterior = (data_atual - timedelta(days=1)).strftime('%d/%m/%Y')
dia_inicio_mes = data_atual.strftime('01/%m/%Y')
data_atual_format = data_atual.strftime('%d/%m/%Y')

# Função para exibir a tela inicial
def exibir_janela_inicial():
    root = Window(themename="cosmo")
    root.title("EasyLog")
    root.iconbitmap(caminhos["icone"])
    root.resizable(False, False)            

    frame_principal = ttk.Frame(root, padding=(70,20))
    frame_principal.pack(fill=tk.BOTH, expand=True)

    # Título e subtítulo
    ttk.Label(frame_principal, text="Bem-vindo(a) ao EasyLog", font=("Helvetica", 16, "bold")).pack(pady=10)
    ttk.Label(frame_principal, text="Tornando seu trabalho mais eficiente", font=("Helvetica", 11)).pack(pady=(10, 20))

    # Botões
    ttk.Button(frame_principal, text="PLANILHAS", command=lambda:abrir_janela(root, "Planilhas"), bootstyle="warning-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="MENSAGENS", command=lambda:abrir_janela(root, "Mensagens"), bootstyle="success-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="HISTÓRICOS",command=lambda:abrir_janela(root, "Históricos"), bootstyle="info-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="SUPORTE", command=lambda:messagebox.showinfo("Aviso", "Funcionalidade em desenvolvimento"), bootstyle="danger-outline", width=20).pack(pady=10)

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
    janela.iconbitmap(caminhos["icone"])
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
    janela.geometry("500x550")
    centralizar_janela(janela)
    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame, text=" Planilha de contatos: * ", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=tk.X, pady=5)
    campo_planilha = ttk.Entry(frame_contatos)
    campo_planilha.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:anexar_planilha(campo_planilha), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando o frame para o tipo de comunicado
    frame_tipo_mensagem = ttk.Labelframe(frame, text=" Tipo de Mensagem * ", padding=5)
    frame_tipo_mensagem.pack(fill=tk.X, padx=5, pady=5)

    # Variável para armazenar o tipo de comunicado selecionado
    tipo_mensagem_var = tk.StringVar(value="falta")

    # Definindo as opções dos RadioButtons
    opcoes = [
        ("Falta", "falta"),
        ("Multirão", "multirão"),
        ("Atenção", "atenção"),              
        ("Reunião de Pais", "reuniao_de_pais"),
        ("Oficina", "oficina"),
        ("Personalizada", "personalizada")
    ]

    # Colocando os RadioButtons em uma grade
    for index, (texto, valor) in enumerate(opcoes):
        row = index // 3  # Calcula em qual linha deve colocar
        column = index % 3  # Calcula a coluna (de 0 a 2)
        ttk.Radiobutton(frame_tipo_mensagem, text=texto, value=valor, variable=tipo_mensagem_var, command=lambda valor=valor: atualizar_campos(valor)).grid(row=row, column=column, sticky="w", padx=30, pady=5)

    # Campos para variáveis
    frame_variaveis = ttk.Labelframe(frame, text=" Variáveis ", padding=5, bootstyle="primary")
    frame_variaveis.pack(fill=tk.X, pady=5)

    ttk.Label(frame_variaveis, text="Data: ").pack(side=tk.LEFT, padx=(5,0))
    campo_data = ttk.Entry(frame_variaveis, width=10)
    campo_data.pack(side=tk.LEFT, padx=(10,5))
    campo_data.insert(0, data_anterior)

    campo_data_inicial = ttk.Entry(frame_variaveis, width=10)
    campo_data_final = ttk.Entry(frame_variaveis, width=10)
    
    campo_tema = ttk.Entry(frame_variaveis, width=29)

    campo_hora = ttk.Entry(frame_variaveis, width=7)

    ttk.Button(frame_variaveis, text="Gerar", command=lambda:gerar_mensagem(tipo_mensagem_var, campo_mensagem, campo_data, campo_hora, campo_tema, campo_data_inicial, campo_data_final)).pack(side=tk.RIGHT, padx=5)

    # Função para atualizar os campos
    def atualizar_campos(valor):
        for widget in frame_variaveis.winfo_children():
            widget.pack_forget()

        campo_data.delete(0, tk.END)  # Limpa a área de texto antes de inserir
        campo_hora.delete(0, tk.END)  # Limpa a área de texto antes de inserir
        campo_tema.delete(0, tk.END)  # Limpa a área de texto antes de inserir

        if valor == "falta":
            ttk.Label(frame_variaveis, text="Data: ").pack(side=tk.LEFT, padx=(5,0))
            campo_data.pack(side=tk.LEFT, padx=5)
            campo_data.insert(0, data_anterior)
        elif valor == "multirão":
            ttk.Label(frame_variaveis, text="Data Inicial: ").pack(side=tk.LEFT, padx=(5,0))
            campo_data_inicial.pack(side=tk.LEFT, padx=5)
            ttk.Label(frame_variaveis, text="Data Final: ").pack(side=tk.LEFT, padx=(5,0))
            campo_data_final.pack(side=tk.LEFT, padx=5)
        elif valor == "reuniao_de_pais":
            ttk.Label(frame_variaveis, text="Data: ").pack(side=tk.LEFT, padx=(5,0))
            campo_data.pack(side=tk.LEFT, padx=5)
            campo_data.insert(0, data_anterior)
        elif valor == "oficina":
            ttk.Label(frame_variaveis, text="Data: ").pack(side=tk.LEFT, padx=(5,0))
            campo_data.pack(side=tk.LEFT, padx=5)
            ttk.Label(frame_variaveis, text="Tema: ").pack(side=tk.LEFT, padx=(5,0))
            campo_tema.pack(side=tk.LEFT, padx=5)
        elif valor == "personalizada" or valor == "atenção":
            ttk.Label(frame_variaveis, text="Não há variáveis para essa opção ").pack(side=tk.LEFT, padx=(5,0))

        ttk.Button(frame_variaveis, text="Gerar", command=lambda:gerar_mensagem(tipo_mensagem_var, campo_mensagem, campo_data, campo_hora, campo_tema, campo_data_inicial, campo_data_final)).pack(side=tk.RIGHT, padx=5)

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
    ttk.Button(frame_imagem, text="Anexar", command=lambda:anexar_imagem(campo_imagem), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando um frame para organizar os botões
    frame_botoes = ttk.Frame(frame)
    frame_botoes.pack(pady=(25,0))

    # Adicionando botões lado a lado usando grid()
    ttk.Button(frame_botoes, text="Voltar", command=janela.destroy, bootstyle="danger-outline", width=10).grid(row=0, column=0, padx=40)
    ttk.Button(frame_botoes, text="Enviar", command=lambda:preparar_envio(campo_planilha, campo_mensagem, campo_imagem),bootstyle="success-outline", width=10).grid(row=0, column=1, padx=40)

# Função para configurar a área de "Históricos"
def frame_historicos(janela, frame):
    janela.geometry("500x540")
    centralizar_janela(janela)

    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame, text=" Planilha de alunos: * ", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=tk.X, pady=5)
    campo_planilha = ttk.Entry(frame_contatos)
    campo_planilha.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:anexar_planilha(campo_planilha), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando o frame para o tipo de comunicado
    frame_comunicado = ttk.Labelframe(frame, text=" Tipo de Ocorrência * ", padding=5, bootstyle="primary")
    frame_comunicado.pack(fill=tk.X, padx=5, pady=5)

    # Variável para armazenar o tipo de comunicado selecionado
    tipo_ocorrencia_var = tk.StringVar(value="falta")

    # Definindo as opções dos RadioButtons
    opcoes = [
        ("Falta", "falta"),
        ("Multirão", "multirao"),
        ("Atenção", "atenção"),
        ("Prova", "prova"),
        ("Atividades", "atividades"),
        ("1° dia de aula", "1_dia_de_aula"),
        ("Personalizada", "personalizada"),
    ]

    # Colocando os RadioButtons em uma grade
    for index, (texto, valor) in enumerate(opcoes):
        row = index // 4  # Calcula em qual linha deve colocar
        column = index % 4  # Calcula a coluna (de 0 a 2)
        ttk.Radiobutton(frame_comunicado, text=texto, value=valor, variable=tipo_ocorrencia_var, command=lambda valor=valor: atualizar_campos(valor)).grid(row=row, column=column, sticky="w", padx=15, pady=5)

    # Campos para variáveis
    frame_variaveis = ttk.Labelframe(frame, text=" Variáveis ", padding=5, bootstyle="primary")
    frame_variaveis.pack(fill=tk.X, pady=5)

    ttk.Label(frame_variaveis, text="Data da falta: ").pack(side=tk.LEFT, padx=(5,0))
    campo_data = ttk.Entry(frame_variaveis, width=10)
    campo_data.pack(side=tk.LEFT, padx=(10,5))
    campo_data.insert(0, data_anterior)

    ttk.Button(frame_variaveis, text="Gerar", command=lambda:gerar_ocorrencia(tipo_ocorrencia_var, campo_data, campo_titulo, campo_descricao)).pack(side=tk.RIGHT, padx=5)

    # Função para atualizar os campos
    def atualizar_campos(valor):
        for widget in frame_variaveis.winfo_children():
            widget.pack_forget()

        campo_data.delete(0, tk.END)  # Limpa a área de texto antes de inserir

        if valor == "falta":
            ttk.Label(frame_variaveis, text="Data da falta: ").pack(side=tk.LEFT, padx=(5,0))
            campo_data.pack(side=tk.LEFT, padx=5)
            campo_data.insert(0, data_anterior)
        else:
            ttk.Label(frame_variaveis, text="Não há variáveis para essa opção ").pack(side=tk.LEFT, padx=(5,0))
    
        ttk.Button(frame_variaveis, text="Gerar", command=lambda:gerar_ocorrencia(tipo_ocorrencia_var, campo_data, campo_titulo, campo_descricao)).pack(side=tk.RIGHT, padx=5)

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
    ttk.Button(frame_botoes, text="Voltar", command=janela.destroy, bootstyle="danger-outline", width=10).grid(row=0, column=0, padx=40)
    ttk.Button(frame_botoes, text="Registrar", command=lambda:preparar_registros(campo_planilha, campo_data, campo_titulo, campo_descricao),bootstyle="success-outline", width=10).grid(row=0, column=1, padx=40)


# Função para configurar a área de "Planilhas"
def frame_planilhas(janela, frame):
    janela.geometry("500x500")
    centralizar_janela(janela)
    
    frame_faltas = ttk.Labelframe(frame, text=" Faltas ", padding=5, bootstyle="primary")
    frame_faltas.pack(fill=tk.X, pady=5)
    ttk.Label(frame_faltas, text="Data Inicial: * ").pack(side=tk.LEFT, padx=(5,0))
    campo_data_inicial = ttk.Entry(frame_faltas, width=12)
    campo_data_inicial.pack(side=tk.LEFT, padx=5)
    campo_data_inicial.insert(0, data_anterior)
    ttk.Label(frame_faltas, text="Data Final: * ").pack(side=tk.LEFT, padx=(5,0))
    campo_data_final = ttk.Entry(frame_faltas, width=12)
    campo_data_final.pack(side=tk.LEFT, padx=5)
    campo_data_final.insert(0, data_anterior)
    ttk.Button(frame_faltas, text="Gerar", command=lambda:preparar_data_faltosos(campo_data_inicial, campo_data_final), bootstyle="primary").pack(side=tk.RIGHT, padx=5)
    
    frame_alunos_atencao = ttk.Labelframe(frame, text=" Alunos em atenção ", padding=5, bootstyle="primary")
    frame_alunos_atencao.pack(fill=tk.X, pady=5)
    ttk.Label(frame_alunos_atencao, text="Data Inicial: * ").pack(side=tk.LEFT, padx=(5,0))
    campo_data_inicial = ttk.Entry(frame_alunos_atencao, width=12)
    campo_data_inicial.pack(side=tk.LEFT, padx=5)
    campo_data_inicial.insert(0, dia_inicio_mes)
    ttk.Label(frame_alunos_atencao, text="Data Final: * ").pack(side=tk.LEFT, padx=(5,0))
    campo_data_final = ttk.Entry(frame_alunos_atencao, width=12)
    campo_data_final.pack(side=tk.LEFT, padx=5)
    campo_data_final.insert(0, data_atual_format)
    ttk.Button(frame_alunos_atencao, text="Gerar", command=lambda:preparar_alunos_atencao(campo_data_inicial, campo_data_final), bootstyle="primary").pack(side=tk.RIGHT, padx=5)

    ttk.Button(frame, text="Voltar", command=janela.destroy, bootstyle="danger-outline", width=10).pack(pady=10)

# Execução do programa      
if __name__ == "__main__":
    criar_pastas()
    exibir_janela_inicial()
