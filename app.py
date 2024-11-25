import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

# Função para selecionar o arquivo de contatos
def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione um arquivo de contatos",
        filetypes=[("Arquivos de texto", "*.txt")]
    )
    if caminho_arquivo:
        entrada_arquivo.config(state="normal")  # Ativa a entrada para inserir o texto
        entrada_arquivo.delete(0, tk.END)      # Limpa o campo atual
        entrada_arquivo.insert(0, caminho_arquivo)  # Insere o caminho do arquivo
        entrada_arquivo.config(state="readonly")  # Desativa a entrada novamente


# Cria a janela principal
janela = tk.Tk()
janela.title("Pedagobot")

# Define o tamanho da janela e centraliza
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
largura_janela = int(largura_tela * 0.4)
altura_janela = int(altura_tela * 0.55)
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2
janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

# Configuração para layout responsivo
janela.grid_rowconfigure(0, weight=1)
janela.grid_columnconfigure(0, weight=1)

# Rótulo para instrução
rotulo_instrucao = tk.Label(janela, text="Selecione uma opção:", font=("Arial", 14))
rotulo_instrucao.grid(row=1, column=0, padx=30, pady=0, sticky="w")

# Caixa de combinação
opcoes = ["Faltosos", "Comunicado"]
combobox = ttk.Combobox(janela, values=opcoes, state="readonly", font=("Arial", 12))  # "readonly" impede edição
combobox.grid(row=2, column=0, padx=(30, 300), pady=(0,5), sticky="w")
combobox.set("Escolha uma opção")  # Define um texto inicial

# Rótulo e botão para selecionar arquivo de contatos
rotulo_contatos = tk.Label(janela, text="Lista de Contatos (arquivo .txt):", font=("Arial", 14))
rotulo_contatos.grid(row=3, column=0, padx=30, pady=(10, 0), sticky="w")
entrada_arquivo = tk.Entry(janela, font=("Arial", 12), width=50)
entrada_arquivo.grid(row=4, column=0, padx=(30, 60), pady=(0, 5), sticky="w")
botao_arquivo = tk.Button(janela, text="Procurar", command=selecionar_arquivo)
botao_arquivo.grid(row=5, column=0, ipadx=7,ipady=3,padx=30, pady=(5, 10), sticky="w")

# Rótulo e caixa de texto para a mensagem
rotulo_mensagem = tk.Label(janela, text="Mensagem:", font=("Arial", 14))
rotulo_mensagem.grid(row=6, column=0, padx=30, pady=(10, 0), sticky="w")
caixa_texto = tk.Text(janela, height=6, font=("Arial", 12), wrap="word", width=30)  # Ajusta a largura
caixa_texto.grid(row=7, column=0, padx=30, pady=(5, 10), sticky="ew")

# Botão para confirmar a seleção
botao = tk.Button(janela, text="Confirmar")
botao.grid(row=8, column=0, ipadx=7,ipady=3,padx=10, pady=(15,20), sticky="")

# Expansão dos elementos (responsividade)
janela.grid_columnconfigure(0, weight=1)

# Inicia o loop principal
janela.mainloop()
