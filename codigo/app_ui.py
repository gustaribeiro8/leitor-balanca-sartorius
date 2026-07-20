import customtkinter as ctk
from tkinter import messagebox
import os
import sys
from datetime import datetime

class AppUI(ctk.CTk):
    def __init__(self, controller):
        """
        Inicializa a interface do usuário (View).
        :param controller: O objeto controlador que gerencia a lógica da aplicação.
        """
        super().__init__()
        self.controller = controller

        # --- CONFIGURAÇÕES DA JANELA ---
        self.title("SISAQUI - Modo Planilha Livre v8.3") 
        self.geometry("750x700")
        self.resizable(False, False)

        caminho_icone = self._resource_path("icone_sartorius.ico")
        try:
            self.iconbitmap(caminho_icone)
        except Exception:
            pass

        self._criar_widgets()
        
        # --- ATALHOS ---
        self.bind('<a>', lambda event: self.controller.capturar_coluna("A"))
        self.bind('<A>', lambda event: self.controller.capturar_coluna("A"))
        self.bind('<b>', lambda event: self.controller.capturar_coluna("B"))
        self.bind('<B>', lambda event: self.controller.capturar_coluna("B"))
        self.bind('<space>', lambda event: self.controller.capturar_coluna("G"))
        
        self.protocol("WM_DELETE_WINDOW", self.controller.on_closing)

    def _resource_path(self, relative_path):
        """Obtém o caminho absoluto para o recurso, lidando com o PyInstaller."""
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        return os.path.join(base_path, relative_path)

    def _criar_widgets(self):
        """Cria e posiciona todos os widgets na janela."""
        # 1. Frame do Arquivo
        frame_arquivo = ctk.CTkFrame(self)
        frame_arquivo.pack(pady=(15, 5), padx=10, fill="x")
        
        lbl_arq = ctk.CTkLabel(frame_arquivo, text="Nome do Ensaio:", font=("Arial", 12, "bold"))
        lbl_arq.pack(side="left", padx=10)

        self.entry_arquivo = ctk.CTkEntry(frame_arquivo, width=220)
        self.entry_arquivo.pack(side="left", padx=5)
        self.entry_arquivo.insert(0, f"ensaio_{datetime.now().strftime('%Y-%m-%d')}") 
        
        self.btn_abrir = ctk.CTkButton(frame_arquivo, text="📂 Abrir no Excel", command=self.controller.abrir_arquivo, width=120)
        self.btn_abrir.pack(side="right", padx=10)

        # 2. Frame de Conexão
        frame_topo = ctk.CTkFrame(self)
        frame_topo.pack(pady=5, padx=10, fill="x")

        lbl_porta = ctk.CTkLabel(frame_topo, text="Porta:", font=("Arial", 11))
        lbl_porta.pack(side="left", padx=(10, 2))

        self.combo_portas = ctk.CTkComboBox(frame_topo, values=["..."], width=100)
        self.combo_portas.pack(side="left", padx=5)
        
        btn_refresh = ctk.CTkButton(frame_topo, text="⟳", width=30, command=self.controller.atualizar_lista_portas)
        btn_refresh.pack(side="left", padx=2)

        self.btn_conexao = ctk.CTkButton(frame_topo, text="Conectar", command=self.controller.alternar_conexao, fg_color="green", width=100)
        self.btn_conexao.pack(side="left", padx=10)
        
        self.btn_tarar = ctk.CTkButton(frame_topo, text="ZERAR BALANÇA", command=self.controller.tarar_balanca, fg_color="#555555", width=120)
        self.btn_tarar.pack(side="right", padx=10)
        
        # 3. Frame do Display de Peso
        frame_display = ctk.CTkFrame(self)
        frame_display.pack(pady=10, padx=20, fill="both", expand=True)

        lbl_titulo = ctk.CTkLabel(frame_display, text="Leitura da Balança", font=("Arial", 14, "bold"), text_color="gray")
        lbl_titulo.pack(pady=(15, 0))

        self.lbl_peso = ctk.CTkLabel(frame_display, text="--- g", font=("Roboto", 70, "bold"))
        self.lbl_peso.pack(pady=10)
        
        self.lbl_status = ctk.CTkLabel(frame_display, text="Aguardando conexão...", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 4. Frame de Ações e Contadores
        frame_acoes = ctk.CTkFrame(self)
        frame_acoes.pack(pady=10, padx=10, fill="x")
        frame_acoes.columnconfigure((0, 1), weight=1)

        self.btn_A = ctk.CTkButton(frame_acoes, text="Capturar para Coluna A", height=50, command=lambda: self.controller.capturar_coluna("A"), fg_color="#2980b9")
        self.btn_A.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.lbl_count_A = ctk.CTkLabel(frame_acoes, text="Nº de Amostras A: 0", font=("Arial", 16, "bold"))
        self.lbl_count_A.grid(row=1, column=0, pady=5)

        self.btn_B = ctk.CTkButton(frame_acoes, text="Capturar para Coluna B", height=50, command=lambda: self.controller.capturar_coluna("B"), fg_color="#e67e22", hover_color="#d35400")
        self.btn_B.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.lbl_count_B = ctk.CTkLabel(frame_acoes, text="Nº de Amostras B: 0", font=("Arial", 16, "bold"))
        self.lbl_count_B.grid(row=1, column=1, pady=5)

        # 5. Textbox de Log
        self.textbox_log = ctk.CTkTextbox(self, height=120)
        self.textbox_log.pack(pady=10, padx=10, fill="x")
        self.log("Dica: As colunas são independentes. Pressione 'A', 'B' ou 'Espaço'.\n")
        
        # Estado inicial dos widgets
        self.set_estado_conectado(False)

    # --- MÉTODOS PÚBLICOS (chamados pelo Controller) ---
    def set_estado_conectado(self, conectado):
        """Habilita/desabilita os widgets com base no estado da conexão."""
        if not self.winfo_exists():
            return

        if conectado:
            self.btn_conexao.configure(text="Desconectar", fg_color="red")
            self.combo_portas.configure(state="disabled")
            self.entry_arquivo.configure(state="disabled")
            self.btn_A.configure(state="normal")
            self.btn_B.configure(state="normal")
            self.btn_tarar.configure(state="normal")
        else:
            self.btn_conexao.configure(text="Conectar", fg_color="green")
            self.combo_portas.configure(state="normal")
            self.entry_arquivo.configure(state="normal")
            self.btn_A.configure(state="disabled")
            self.btn_B.configure(state="disabled")
            self.btn_tarar.configure(state="disabled")
            self.atualizar_peso_display(None, False)

    def atualizar_lista_portas(self, portas):
        """Atualiza a lista de portas no ComboBox."""
        self.combo_portas.configure(values=portas)
        self.combo_portas.set(portas[0] if portas else "Nenhuma")

    def atualizar_peso_display(self, peso_float, estavel):
        """Atualiza o label que exibe o peso."""
        cor_texto = "white" if estavel else "orange"
        texto_peso = f"{peso_float:.6f} g" if peso_float is not None else "--- g"
        self.lbl_peso.configure(text=texto_peso, text_color=cor_texto)

    def atualizar_status(self, texto, cor):
        """Atualiza o label de status da conexão."""
        self.lbl_status.configure(text=texto, text_color=cor)
        
    def atualizar_contadores(self, count_A, count_B):
        """Atualiza os labels que mostram o número de amostras."""
        self.lbl_count_A.configure(text=f"Nº de Amostras A: {count_A}")
        self.lbl_count_B.configure(text=f"Nº de Amostras B: {count_B}")

    def atualizar_contadores_lote(self, progresso_A, progresso_B):
        """Atualiza o contador para mostrar o progresso do lote (ex: 2/4)."""
        self.lbl_count_A.configure(text=f"Nº de Amostras A: {progresso_A}")
        self.lbl_count_B.configure(text=f"Nº de Amostras B: {progresso_B}")

    def log(self, msg):
        """Adiciona uma mensagem ao textbox de log."""
        if self.winfo_exists() and getattr(self, "textbox_log", None) and self.textbox_log.winfo_exists():
            self.textbox_log.insert("end", msg + "\n")
            self.textbox_log.see("end")

    def flash_button(self, coluna):
        """Animação de 'flash' para um botão de captura."""
        if not self.winfo_exists():
            return

        if coluna == "A":
            btn = self.btn_A
            original_color = "#2980b9"
            flash_color = "#5dade2" # Azul mais claro
        else: # Coluna B
            btn = self.btn_B
            original_color = "#e67e22"
            flash_color = "#f5b041" # Laranja mais claro

        if btn and btn.winfo_exists():
            btn.configure(fg_color=flash_color)
            self.after(150, lambda: btn.configure(fg_color=original_color) if btn.winfo_exists() else None)

    def get_nome_ensaio(self):
        """Retorna o nome do ensaio inserido pelo usuário."""
        return self.entry_arquivo.get()

    def get_porta_selecionada(self):
        """Retorna a porta serial selecionada no ComboBox."""
        return self.combo_portas.get()

    @staticmethod
    def show_info(titulo, mensagem):
        messagebox.showinfo(titulo, mensagem)

    @staticmethod
    def show_warning(titulo, mensagem):
        messagebox.showwarning(titulo, mensagem)

    @staticmethod
    def show_error(titulo, mensagem):
        messagebox.showerror(titulo, mensagem)
        
    @staticmethod
    def show_confirmation(titulo, mensagem):
        """Mostra uma caixa de diálogo de confirmação (Sim/Não) e retorna True/False."""
        return messagebox.askyesno(titulo, mensagem, icon='question')

    def mainloop(self):
        super().mainloop()