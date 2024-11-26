import tkinter as tk
from tkinter import filedialog, messagebox, Menu

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Mensagens")
        self.root.geometry("450x450")

        # Centralizar a janela
        self.center_window()

        # Criar menu
        self.menu = Menu(root)
        self.root.config(menu=self.menu)

        # Adicionar abas ao menu
        self.menu.add_command(label="Mensagens")
        self.menu.add_command(label="Históricos", state=tk.DISABLED)  # Desabilitado
        self.menu.add_command(label="Coletar Dados", state=tk.DISABLED)  # Desabilitado

        # Frame principal
        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=20)

        # Elementos da aba Mensagens
        self.create_mensagens_tab()

    def center_window(self):
        # Obtém as dimensões da tela
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Calcula a posição x e y para centralizar a janela
        x = (screen_width // 2) - (450 // 2)
        y = (screen_height // 2) - (450 // 2)

        # Define a posição da janela
        self.root.geometry(f"450x450+{x}+{y}")

    def create_mensagens_tab(self):
        self.label_arquivo = tk.Label(self.frame, text="Anexar arquivo .txt com contatos:")
        self.label_arquivo.pack()

        self.txt_arquivo = tk.Text(self.frame, height=1, width=40)
        self.txt_arquivo.pack(pady=5)

        self.btn_anexar = tk.Button(self.frame, text="Escolher arquivo", command=self.load_file)
        self.btn_anexar.pack()

        self.label_tipo_mensagem = tk.Label(self.frame, text="Escolha o tipo da mensagem:")
        self.label_tipo_mensagem.pack(pady=(20,0))

        self.var_tipo = tk.StringVar(value="texto")
        self.radio_texto = tk.Radiobutton(self.frame, text="Texto", variable=self.var_tipo, value="texto", command=self.update_message_input)
        self.radio_imagem = tk.Radiobutton(self.frame, text="Imagem", variable=self.var_tipo, value="imagem", command=self.update_message_input)
        self.radio_texto.pack()
        self.radio_imagem.pack()

        self.input_mensagem = None
        self.txt_imagem = None
        self.input_imagem = None
        self.btn_enviar = None  # Inicializa o botão de enviar como None

        self.update_message_input()  # Atualiza a interface com base na seleção inicial

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            messagebox.showinfo("Arquivo Selecionado", f"Arquivo selecionado: {file_path}")

    def update_message_input(self):
         # Remove os campos anteriores, se existirem
        if self.input_mensagem:
            self.input_mensagem.destroy()
        if self.input_imagem:
            self.input_imagem.destroy()
        if self.btn_enviar:  # Destrói o botão de enviar se existir
            self.btn_enviar.destroy()
        if self.txt_imagem:
            self.txt_imagem.destroy()

        # Adiciona novo campo baseado na seleção
        if self.var_tipo.get() == "texto":
            self.input_mensagem = tk.Text(self.frame, height=10, width=40)
            self.input_mensagem.pack(pady=10)
            self.input_mensagem.insert(tk.INSERT, "Digite o modelo da mensagem")
            
        else:
            self.txt_imagem = tk.Text(self.frame, height=1, width=40)
            self.txt_imagem.pack()
            self.input_imagem = tk.Button(self.frame, text="Escolher imagem", command=self.load_image)
            self.input_imagem.pack(pady=10)

        # Cria o botão de enviar
        self.btn_enviar = tk.Button(self.frame, text="Iniciar Envio", command=self.send_message)
        self.btn_enviar.pack(pady=10)


    def load_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if file_path:
            messagebox.showinfo("Imagem Selecionada", f"Imagem selecionada: {file_path}")

    def send_message(self):
        if self.var_tipo.get() == "texto":
            message = self.input_mensagem.get("1.0", "end-1c")
            messagebox.showinfo("Mensagem Enviada", f"Mensagem enviada: {message}")
        else:
            messagebox.showinfo("Mensagem Enviada", "Imagem enviada com sucesso!")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()