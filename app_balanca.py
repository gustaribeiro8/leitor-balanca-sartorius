import customtkinter as ctk
import serial
import serial.tools.list_ports
import time
import re
import csv
import threading
import os
import sys
from tkinter import messagebox

# --- CONFIGURAÇÕES ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class AppBalanca(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("SISAQUI - Modo Planilha Livre v8.1") 
        self.geometry("750x700")
        self.resizable(False, False)

        caminho_icone = resource_path("icone_sartorius.ico")
        try:
            self.iconbitmap(caminho_icone)
        except:
            pass

        self.ser = None
        self.monitorando = False
        self.ultimo_peso_valido = None
        
        # --- MEMÓRIA DA PLANILHA (LISTA DE DICIONÁRIOS) ---
        # Estrutura simplificada: [{'A': '100,0005', 'B': ''}, ...]
        self.dados_memoria = [] 
        self.cursor_A = 0 
        self.cursor_B = 0 

        self.criar_interface()
        self.listar_portas_disponiveis()
        
        # Atalhos
        self.bind('<a>', lambda event: self.capturar_coluna("A"))
        self.bind('<A>', lambda event: self.capturar_coluna("A"))
        self.bind('<b>', lambda event: self.capturar_coluna("B"))
        self.bind('<B>', lambda event: self.capturar_coluna("B"))
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def criar_interface(self):
        # 1. Arquivo
        self.frame_arquivo = ctk.CTkFrame(self)
        self.frame_arquivo.pack(pady=(15, 5), padx=10, fill="x")
        
        self.lbl_arq = ctk.CTkLabel(self.frame_arquivo, text="Ensaio:", font=("Arial", 12, "bold"))
        self.lbl_arq.pack(side="left", padx=10)

        self.entry_arquivo = ctk.CTkEntry(self.frame_arquivo, width=200)
        self.entry_arquivo.pack(side="left", padx=5)
        self.entry_arquivo.insert(0, "ensaio_livre") 
        
        self.btn_abrir = ctk.CTkButton(self.frame_arquivo, text="📂 Excel", command=self.abrir_tabela, width=80)
        self.btn_abrir.pack(side="right", padx=10)

        # 2. Conexão
        self.frame_topo = ctk.CTkFrame(self)
        self.frame_topo.pack(pady=5, padx=10, fill="x")

        self.lbl_porta = ctk.CTkLabel(self.frame_topo, text="Porta:", font=("Arial", 11))
        self.lbl_porta.pack(side="left", padx=(10, 2))

        self.combo_portas = ctk.CTkComboBox(self.frame_topo, values=["..."], width=100)
        self.combo_portas.pack(side="left", padx=5)
        self.btn_refresh = ctk.CTkButton(self.frame_topo, text="⟳", width=30, command=self.listar_portas_disponiveis)
        self.btn_refresh.pack(side="left", padx=2)

        self.btn_conexao = ctk.CTkButton(self.frame_topo, text="Conectar", command=self.alternar_conexao, fg_color="green", width=100)
        self.btn_conexao.pack(side="left", padx=10)
        
        self.btn_tarar = ctk.CTkButton(self.frame_topo, text="ZERAR", command=self.comando_tarar, fg_color="#555555", width=80)
        self.btn_tarar.pack(side="right", padx=10)
        self.btn_tarar.configure(state="disabled")

        # 3. Display
        self.frame_display = ctk.CTkFrame(self)
        self.frame_display.pack(pady=10, padx=20, fill="both", expand=True)

        self.lbl_titulo = ctk.CTkLabel(self.frame_display, text="LEITURA REAL-TIME", font=("Arial", 14, "bold"), text_color="gray")
        self.lbl_titulo.pack(pady=(15, 0))

        self.lbl_peso = ctk.CTkLabel(self.frame_display, text="--- g", font=("Roboto", 70, "bold"))
        self.lbl_peso.pack(pady=10)
        
        self.lbl_status = ctk.CTkLabel(self.frame_display, text="Conecte para iniciar", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 4. Ações e Contadores
        self.frame_acoes = ctk.CTkFrame(self)
        self.frame_acoes.pack(pady=10, padx=10, fill="x")
        self.frame_acoes.columnconfigure(0, weight=1)
        self.frame_acoes.columnconfigure(1, weight=1)

        # Botão A
        self.btn_A = ctk.CTkButton(self.frame_acoes, text="CAPTURA COLUNA A", height=50, 
                                   command=lambda: self.capturar_coluna("A"), fg_color="#2980b9")
        self.btn_A.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        # Contador A
        self.lbl_count_A = ctk.CTkLabel(self.frame_acoes, text="Amostras A: 0", font=("Arial", 16, "bold"))
        self.lbl_count_A.grid(row=1, column=0, pady=5)

        # Botão B
        self.btn_B = ctk.CTkButton(self.frame_acoes, text="CAPTURA COLUNA B", height=50, 
                                   command=lambda: self.capturar_coluna("B"), fg_color="#e67e22")
        self.btn_B.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Contador B
        self.lbl_count_B = ctk.CTkLabel(self.frame_acoes, text="Amostras B: 0", font=("Arial", 16, "bold"))
        self.lbl_count_B.grid(row=1, column=1, pady=5)

        self.btn_A.configure(state="disabled")
        self.btn_B.configure(state="disabled")

        # Log
        self.textbox_log = ctk.CTkTextbox(self, height=80)
        self.textbox_log.pack(pady=10, padx=10, fill="x")
        self.textbox_log.insert("0.0", "Dica: Colunas independentes. Pode fazer varios A e depois varios B.\n")

    # --- FUNÇÕES BÁSICAS ---
    def listar_portas_disponiveis(self):
        portas = serial.tools.list_ports.comports()
        lista = [p.device for p in portas] if portas else ["Nenhuma"]
        self.combo_portas.configure(values=lista)
        self.combo_portas.set(lista[0])

    def get_nome_arquivo(self):
        pasta = "dados coletados"
        if not os.path.exists(pasta): os.makedirs(pasta)
        nome = self.entry_arquivo.get().strip() or "dados_sisaqui"
        return os.path.join(pasta, nome + ".csv" if not nome.endswith(".csv") else nome)

    def abrir_tabela(self):
        arquivo = self.get_nome_arquivo()
        if os.path.exists(arquivo): os.startfile(arquivo)
        else: messagebox.showwarning("Aviso", "Arquivo ainda não existe.")

    def alternar_conexao(self):
        if self.ser and self.ser.is_open:
            self.monitorando = False
            time.sleep(0.5)
            self.ser.close()
            self.ser = None
            self.btn_conexao.configure(text="Conectar", fg_color="green")
            self.combo_portas.configure(state="normal")
            self.entry_arquivo.configure(state="normal")
            self.btn_A.configure(state="disabled")
            self.btn_B.configure(state="disabled")
            self.btn_tarar.configure(state="disabled")
            self.lbl_status.configure(text="Desconectado", text_color="gray")
            self.log("Desconectado.")
        else:
            porta = self.combo_portas.get()
            if porta in ["Nenhuma", "..."]: return
            try:
                self.ser = serial.Serial(porta, 1200, bytesize=7, parity='O', stopbits=1, timeout=0.5)
                self.monitorando = True
                threading.Thread(target=self.thread_monitoramento, daemon=True).start()
                self.btn_conexao.configure(text="Desconectar", fg_color="red")
                self.combo_portas.configure(state="disabled")
                self.entry_arquivo.configure(state="disabled")
                self.btn_A.configure(state="normal")
                self.btn_B.configure(state="normal")
                self.btn_tarar.configure(state="normal")
                self.lbl_status.configure(text=f"Conectado em {porta}", text_color="#00FF00")
                self.log(f"Conectado! Pode iniciar as leituras.")
            except Exception as e:
                messagebox.showerror("Erro", f"{e}")

    def thread_monitoramento(self):
        while self.monitorando and self.ser and self.ser.is_open:
            try:
                self.ser.reset_input_buffer()
                self.ser.write(b'\x1bP\r\n') 
                linha = self.ser.readline().decode('ascii', errors='ignore')
                match = re.search(r"[-+]?\s*\d+\.\d+", linha)
                if match:
                    self.ultimo_peso_valido = float(match.group().replace(" ", ""))
                    self.lbl_peso.configure(text=f"{self.ultimo_peso_valido} g")
                time.sleep(0.2)
            except:
                time.sleep(1)

    def comando_tarar(self):
        if self.ser: self.ser.write(b'\x1bf4_\r\n'); self.log("Comando TARA enviado.")

    # --- LÓGICA DE PLANILHA LIVRE (SEM DATA/HORA) ---
    def capturar_coluna(self, coluna):
        if not self.monitorando or self.ultimo_peso_valido is None: return
        
        peso_str = str(self.ultimo_peso_valido).replace('.', ',')
        
        # Define qual linha vamos editar/criar
        if coluna == "A":
            idx = self.cursor_A
        else:
            idx = self.cursor_B

        # Se a linha ainda não existe na memória, cria linhas vazias com apenas A e B
        while len(self.dados_memoria) <= idx:
            self.dados_memoria.append({'A': '', 'B': ''})

        # Atualiza a linha existente
        if coluna == "A":
            self.dados_memoria[idx]['A'] = peso_str
            self.cursor_A += 1
            self.lbl_count_A.configure(text=f"Amostras A: {self.cursor_A}")
        else:
            self.dados_memoria[idx]['B'] = peso_str
            self.cursor_B += 1
            self.lbl_count_B.configure(text=f"Amostras B: {self.cursor_B}")
        
        self.salvar_arquivo_completo()
        self.flash_button(self.btn_A if coluna == "A" else self.btn_B)
        
        # Log visual
        diff = self.cursor_A - self.cursor_B
        if diff == 0:
            self.log(f"✅ {coluna} salvo. (Linhas alinhadas!)")
        else:
            self.log(f"✅ {coluna} salvo.")

    def salvar_arquivo_completo(self):
        arquivo = self.get_nome_arquivo()
        try:
            # Modo 'w' = Write (Sobrescreve o arquivo inteiro)
            with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                # CABEÇALHO LIMPO
                writer.writerow(['Peso Padrao A (g)', 'Peso Cliente B (g)'])
                
                # Percorre a memória e escreve linha a linha
                for i, linha in enumerate(self.dados_memoria):
                    writer.writerow([linha['A'], linha['B']])
                    
                    # Regra de pular linha a cada 4 linhas PREENCHIDAS
                    if (i + 1) % 4 == 0:
                        writer.writerow([])
                        
        except PermissionError:
            messagebox.showerror("Erro", "Excel aberto! Feche para salvar.")

    def flash_button(self, btn):
        c = btn._fg_color
        btn.configure(fg_color="white", text_color="black")
        self.after(100, lambda: btn.configure(fg_color=c, text_color="white"))

    def log(self, msg):
        self.textbox_log.insert("0.0", msg + "\n")

    def on_closing(self):
        self.monitorando = False
        if self.ser and self.ser.is_open: self.ser.close()
        self.destroy()

if __name__ == "__main__":
    app = AppBalanca()
    app.mainloop()