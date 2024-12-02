import tkinter as tk  # Importando tkinter para acesso a constantes
from ttkbootstrap import Window, ttk  # Importando ttkbootstrap para personalização do tema
from automação import *
from ttkbootstrap import Style


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
    if titulo == "Faltosos":
        frame_faltosos(janela, frame)
    elif titulo == "Comunicados":
        frame_comunicados(janela,frame)
    elif titulo == "Históricos":
        frame_historicos(janela,frame)
    elif titulo == "Planilhas":
        frame_planilhas(janela,frame)
    else:
        ttk.Label(frame, text="Conteúdo não configurado.", font=("Helvetica", 12)).pack(pady=10)

# Função para configurar a área de "Faltosos"
def frame_faltosos(janela, frame):
    janela.geometry("500x450")
    centralizar_janela(janela)

    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame, text="Planilha de contatos:", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=tk.X, pady=5)
    campo_planilha = ttk.Entry(frame_contatos)
    campo_planilha.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:anexar_planilha(campo_planilha), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Campos para dados adicionais
    frame_dados = ttk.Labelframe(frame, text="Informações Adicionais:", padding=5, bootstyle="primary")
    frame_dados.pack(fill=tk.X, pady=5)

    ttk.Label(frame_dados, text="Professor:").pack(side=tk.LEFT, padx=(5,0))
    campo_nome_professor = ttk.Entry(frame_dados, width=20)
    campo_nome_professor.pack(side=tk.LEFT, padx=5)

    ttk.Label(frame_dados, text="Dia da Falta:").pack(side=tk.LEFT, padx=(5,0))
    campo_dia_falta = ttk.Entry(frame_dados, width=10)
    campo_dia_falta.pack(side=tk.LEFT, padx=5)
    
    ttk.Button(frame_dados, text="Aplicar", command=lambda:mensagem_para_verificacao(campo_nome_professor, campo_dia_falta, campo_mensagem), bootstyle="primary").pack(side=tk.RIGHT, padx=5)

     # Campo para digitação do modelo de mensagem
    frame_texto = ttk.Labelframe(frame, text="Modelo da mensagem (opcional):", padding=5, bootstyle="primary")
    frame_texto.pack(fill=tk.BOTH, pady=5)
    campo_mensagem = tk.Text(frame_texto, height=7, wrap="word")
    campo_mensagem.pack(fill=tk.BOTH, padx=5)

    # Campo para anexar imagem
    frame_imagem = ttk.Labelframe(frame, text="Imagem (opcional):", padding=5, bootstyle="primary")
    frame_imagem.pack(fill=tk.X, pady=5)
    campo_imagem = ttk.Entry(frame_imagem)
    campo_imagem.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_imagem, text="Anexar", command=lambda:anexar_imagem(campo_imagem), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando um frame para organizar os botões
    frame_botoes = ttk.Frame(frame)
    frame_botoes.pack(pady=10)

    # Adicionando botões lado a lado usando grid()
    ttk.Button(frame_botoes, text="Enviar", command=lambda:preparar_envio(campo_planilha, campo_mensagem, campo_imagem),bootstyle="success-outline").grid(row=0, column=0, padx=40)
    ttk.Button(frame_botoes, text="Voltar", command=janela.destroy, bootstyle="danger-outline").grid(row=0, column=1, padx=40)

# Função para configurar a área de "Comunicados"
def frame_comunicados(janela,frame):
    janela.geometry("500x520")
    centralizar_janela(janela)
    # Campo para anexação de planilha
    frame_contatos = ttk.Labelframe(frame, text="Planilha de contatos:", padding=5, bootstyle="primary")
    frame_contatos.pack(fill=tk.X, pady=5)
    campo_planilha = ttk.Entry(frame_contatos)
    campo_planilha.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_contatos, text="Anexar", command=lambda:anexar_planilha(campo_planilha), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Campo para digitação do modelo de mensagem
    frame_texto = ttk.Labelframe(frame, text="Modelo da mensagem (opcional):", padding=5, bootstyle="primary")
    frame_texto.pack(fill=tk.BOTH, pady=5)
    campo_mensagem = tk.Text(frame_texto, height=7, wrap="word")
    campo_mensagem.pack(fill=tk.BOTH, padx=5)

    # Criando o frame para o tipo de comunicado
    frame_comunicado = ttk.Labelframe(frame, text="Tipo de Comunicado", padding=5)
    frame_comunicado.pack(fill=tk.X, padx=5, pady=5)

    # Variável para armazenar o tipo de comunicado selecionado
    tipo_comunicado_var = tk.StringVar(value="multirão")

    # Definindo as opções dos RadioButtons
    opcoes = [
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
        ttk.Radiobutton(frame_comunicado, text=texto, value=valor, variable=tipo_comunicado_var).grid(row=row, column=column, sticky="w", padx=30, pady=5)

    # Botão para aplicar a seleção e inserir a mensagem
    ttk.Button(frame, text="Gerar mensagem", command=lambda:aplicar_comunicado(tipo_comunicado_var, campo_mensagem)).pack(pady=10)

    # Campo para anexar imagem
    frame_imagem = ttk.Labelframe(frame, text="Imagem (opcional):", padding=5, bootstyle="primary")
    frame_imagem.pack(fill=tk.X, pady=5)
    campo_imagem = ttk.Entry(frame_imagem)
    campo_imagem.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    ttk.Button(frame_imagem, text="Anexar", command=lambda:anexar_imagem(campo_imagem), bootstyle="success").pack(side=tk.RIGHT, padx=5)

    # Criando um frame para organizar os botões
    frame_botoes = ttk.Frame(frame)
    frame_botoes.pack(pady=10)

    # Adicionando botões lado a lado usando grid()
    ttk.Button(frame_botoes, text="Enviar", command=lambda:preparar_envio(campo_planilha, campo_mensagem, campo_imagem),bootstyle="success-outline").grid(row=0, column=0, padx=40)
    ttk.Button(frame_botoes, text="Voltar", command=janela.destroy, bootstyle="danger-outline").grid(row=0, column=1, padx=40)

# Função para configurar a área de "Históricos"
def frame_historicos(janela, frame):
    janela.geometry("500x650")
    centralizar_janela(janela)
    ttk.Label(frame, text="Em desenvolvimento", font=("Helvetica", 14, "bold")).pack(pady=10)


# Função para configurar a área de "Planilhas"
def frame_planilhas(janela, frame):
    janela.geometry("500x500")
    centralizar_janela(janela)
    ttk.Label(frame, text="Gerenciamento de Planilhas", font=("Helvetica", 14, "bold")).pack(pady=10)
    ttk.Label(frame, text="Selecione a planilha desejada:", font=("Helvetica", 10)).pack(pady=(5, 10))
    ttk.Combobox(frame, values=["Planilha 1", "Planilha 2", "Planilha 3"]).pack(pady=5)
    ttk.Button(frame, text="Abrir", bootstyle="primary-outline").pack(pady=10)

# Função para exibir a tela inicial
def exibir_janela_inicial():
    root = Window(themename="cosmo")
    root.title("Pedagobot")
    root.iconbitmap("./imagens/icone.ico")

    frame_principal = ttk.Frame(root, padding=20)
    frame_principal.pack(fill=tk.BOTH, expand=True)

    # Título e subtítulo
    ttk.Label(frame_principal, text="Bem-vindo(a) ao Pedagobot", font=("Helvetica", 16, "bold")).pack(pady=10)
    ttk.Label(frame_principal, text="Tornando seu trabalho mais leve", font=("Helvetica", 11)).pack(pady=(10, 20))

    # Botões
    ttk.Button(frame_principal, text="FALTOSOS", command=lambda:abrir_janela(root, "Faltosos"), bootstyle="primary-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="COMUNICADOS", command=lambda:abrir_janela(root, "Comunicados"), bootstyle="success-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="HISTÓRICOS",command=lambda:messagebox.showinfo("Aviso", "Funcionalidade em desenvolvimento"), bootstyle="info-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="PLANILHAS", command=lambda:messagebox.showinfo("Aviso", "Funcionalidade em desenvolvimento"), bootstyle="warning-outline", width=20).pack(pady=10)

    # Rodapé com versão
    ttk.Label(frame_principal, text="Versão 1.0", font=("Helvetica", 9)).pack(pady=(20, 10))

    centralizar_janela(root)
    root.mainloop()

# Execução do programa
if __name__ == "__main__":
    exibir_janela_inicial()
