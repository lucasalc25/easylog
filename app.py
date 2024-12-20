import tkinter as tk  # Importando tkinter para acesso a constantes
from tkinter import messagebox
from ttkbootstrap import Window, ttk  # Importando ttkbootstrap para personalização do tema
from datetime import datetime, timedelta
from config import caminhos
from bot import anexar_imagem, anexar_planilha, criar_pastas
from scripts.historicos import gerar_ocorrencia, preparar_registros
from scripts.mensagens import gerar_mensagem, preparar_envio
from scripts.planilhas import gerar_planilha

# Obter a data atual
data_atual = datetime.now()

# Subtrair um dia da data atual
data_anterior = (data_atual - timedelta(days=1)).strftime('%d/%m/%Y')
dia_inicio_mes = data_atual.strftime('01/%m/%Y')
data_atual_format = data_atual.strftime('%d/%m/%Y')
mes_atual = data_atual.month - 1

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
    ttk.Label(frame_principal, text="Facilitando sua gestão acadêmica", font=("Helvetica", 11)).pack(pady=(10, 20))

    # Botões
    ttk.Button(frame_principal, text="PLANILHAS", command=lambda:abrir_janela(root, "Planilhas"), bootstyle="success-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="MENSAGENS", command=lambda:abrir_janela(root, "Mensagens"), bootstyle="primary-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="HISTÓRICOS",command=lambda:abrir_janela(root, "Históricos"), bootstyle="info-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="AJUDA", command=lambda:messagebox.showinfo("Aviso", "Funcionalidade em desenvolvimento"), bootstyle="warning-outline", width=20).pack(pady=10)
    ttk.Button(frame_principal, text="FECHAR", command=lambda:root.quit(), bootstyle="danger-outline", width=20).pack(pady=10)

    # Rodapé com versão
    ttk.Label(frame_principal, text="Versão 1.1", font=("Helvetica", 9)).pack(pady=(20, 0))

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
    if titulo == "Planilhas":
        frame_planilhas(janela,frame)
    elif titulo == "Mensagens":
        frame_mensagens(janela, frame)
    elif titulo == "Históricos":
        frame_historicos(janela,frame)
    else:
        ttk.Label(frame, text="Conteúdo não configurado.", font=("Helvetica", 12)).pack(pady=10)

# Função para atualizar as opções de horas das aulas
def mudar_hora_aulas(frame_frequencia, dia_da_semana, sala, hora_aula):
    sala_selecionada = sala.get()
    dia_selecionado = dia_da_semana.get()

    for i, widget in enumerate(frame_frequencia.winfo_children()):
        if i >= 4:    
            widget.pack_forget()

            dia_da_semana.delete(0, tk.END)  # Limpa a área de texto antes de inserir
            sala.delete(0, tk.END)  # Limpa a área de texto antes de inserir
            hora_aula.delete(0, tk.END)  # Limpa a área de texto antes de inserir

            if dia_selecionado != "Sábado":
                if sala_selecionada == "Dinamica 1":
                    ttk.Label(frame_frequencia, text="Hora: * ").pack(side="left", padx=5, pady=(10, 15))
                    hora_aula.pack(side="left", padx=(0, 5), pady=(10, 15))  
                    hora_aula["values"] = ["08:00:00", "09:00:00", "10:00:00", "13:00:00", "14:00:00", "15:00:00", "16:00:00", "17:00:00", "18:00:00", "19:00:00"]
                    hora_aula.current(0)
                elif sala_selecionada == "Dinamica 2":
                    ttk.Label(frame_frequencia, text="Hora: * ").pack(side="left", padx=5, pady=(10, 15))
                    hora_aula.pack(side="left", padx=(0, 5), pady=(10, 15))
                    hora_aula["values"] = ["08:00:00", "10:00:00", "13:00:00", "15:00:00", "17:00:00", "18:00:00"]
                    hora_aula.current(0)
            elif dia_selecionado == "Sábado":
                ttk.Label(frame_frequencia, text="Hora: * ").pack(side="left", padx=5, pady=(10, 15))
                hora_aula.pack(side="left", padx=(0, 5), pady=(10, 15))
                hora_aula["values"] = ["08:00:00", "10:00:00", "12:00:00", "14:00:00", "16:00:00"]
                hora_aula.current(0)


# Função para configurar a área de "Planilhas"
def frame_planilhas(janela, frame):
    janela.geometry("500x500")
    centralizar_janela(janela)
    
    frame_faltas = ttk.Labelframe(frame, text=" Faltas no dia", padding=5, bootstyle="primary")
    frame_faltas.pack(fill=tk.X, pady=(0,5))
    # Adicionando os elementos existentes usando grid
    ttk.Label(frame_faltas, text="Data da falta: *").pack(side=tk.LEFT, padx=(10,0), pady=(10, 15))
    data_falta = ttk.Entry(frame_faltas, width=12)
    data_falta.pack(side=tk.LEFT, padx=5, pady=(10, 15))
    data_falta.insert(0, data_anterior)
    ttk.Label(frame_faltas, text="Educador: *").pack(side=tk.LEFT, padx=(10,0), pady=(10, 15))
    filtro_educador = ttk.Combobox(frame_faltas, values=["Geral", "Lucas", "Linderlly", "Yasmin"], state="readonly", width=10, height=3, justify="center")
    filtro_educador.pack(side=tk.LEFT, padx=5, pady=(10, 15))       
    filtro_educador.current(0)  # Define "Nenhum" como o valor padrão
    ttk.Button(frame_faltas, text="Gerar", command=lambda:gerar_planilha("faltas_do_dia", data_falta, data_falta, filtro_educador, "campo_dia_da_semana", "campo_sala", "campo_hora"), bootstyle="primary").pack(side=tk.RIGHT, padx=10, pady=(10, 15))
    
    frame_alunos_atencao = ttk.Labelframe(frame, text=" Faltas no mês ", padding=5, bootstyle="primary")
    frame_alunos_atencao.pack(fill=tk.X, pady=(0,5))
    ttk.Label(frame_alunos_atencao, text="Data Inicial: * ").pack(side=tk.LEFT, padx=(10,0), pady=(10, 15))
    data_inicial = ttk.Entry(frame_alunos_atencao, width=12)
    data_inicial.pack(side=tk.LEFT, padx=5, pady=(10, 15))
    data_inicial.insert(0, dia_inicio_mes)
    ttk.Label(frame_alunos_atencao, text="Data Final: * ").pack(side=tk.LEFT, padx=(10,0), pady=(10, 15))
    data_final = ttk.Entry(frame_alunos_atencao, width=12)
    data_final.pack(side=tk.LEFT, padx=5, pady=(10, 15))
    data_final.insert(0, data_atual_format)
    ttk.Button(frame_alunos_atencao, text="Gerar", command=lambda:gerar_planilha("faltas_do_mes", data_inicial, data_final, "Geral", "campo_dia_da_semana", "campo_sala", "campo_hora"), bootstyle="primary").pack(side=tk.RIGHT, padx=10, pady=(10, 15))

    frame_frequencia = ttk.Labelframe(frame, text=" Listas de Frequência ", padding=5, bootstyle="primary")
    frame_frequencia.pack(pady=(0, 5), fill="x")
    # Linha de widgets lado a lado
    ttk.Label(frame_frequencia, text="Dia: * ").pack(side="left", padx=5, pady=(10, 15))
    dia_da_semana = ttk.Combobox(frame_frequencia, values=["Segunda-Feira", "Terça-Feira", "Quarta-Feira", "Quinta-Feira", "Sexta-Feira", "Sábado"], state="readonly", width=14, justify="center")
    dia_da_semana.pack(side="left", padx=(0, 5), pady=(10, 15))
    dia_da_semana.current(0)
    # Bind para atualizar horas
    dia_da_semana.bind("<<ComboboxSelected>>", lambda event: mudar_hora_aulas(frame_frequencia, dia_da_semana, sala, hora_aula))
    ttk.Label(frame_frequencia, text="Sala: * ").pack(side="left", padx=5, pady=(10, 15))
    sala = ttk.Combobox(frame_frequencia, values=["Dinamica 1", "Dinamica 2"], state="readonly", width=10, justify="center")
    sala.pack(side="left", padx=(0, 5), pady=(10, 15))
    # Bind para atualizar horas
    sala.bind("<<ComboboxSelected>>", lambda event: mudar_hora_aulas(frame_frequencia, dia_da_semana, sala, hora_aula))
    hora_aula = ttk.Combobox(frame_frequencia, state="readonly", width=9, justify="center")
    # Botão na linha debaixo
    ttk.Button(frame, text="Gerar", command=lambda: gerar_planilha("frequencia", data_inicial, data_final, filtro_educador, dia_da_semana, sala, hora_aula), bootstyle="primary").pack(pady=(10, 15))
    
    ttk.Button(frame, text="Voltar", command=janela.destroy, bootstyle="danger-outline", width=10).pack(pady=10)

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


# Execução do programa      
if __name__ == "__main__":
    criar_pastas()
    exibir_janela_inicial()
