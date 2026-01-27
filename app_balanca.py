import customtkinter as ctk
import serial
import serial.tools.list_ports # Biblioteca para achar as portas
import time
import re
import csv
import threading
import os
import sys
from datetime import datetime
from tkinter import messagebox

# --- CONFIGURAÇÕES VISUAIS ---
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

        self.title("SISAQUI - Sistema de Aquisição Universal") 
        self.geometry("700x650") # Aumentei um pouco para caber o menu de portas
        self.resizable(False, False)

        caminho_icone = resource_path("icone_sartorius.ico")
        try:
            self.iconbitmap(caminho_icone)
        except:
            pass

        self.ser = None
        self.monitorando = False
        self.ultimo_peso_valido = None
        
        self.criar_interface()
        self.listar_portas_disponiveis() # Já busca as portas ao abrir
        
        # Atalhos
        self.bind('<a>', lambda event: self.salvar_medida("Padrao (A)"))
        self.bind('<A>', lambda event: self.salvar_medida("Padrao (A)"))
        self.bind('<b>', lambda event: self.salvar_medida("Cliente (B)"))
        self.bind('<B>', lambda event: self.salvar_medida("Cliente (B)"))
        self.bind('<space>', lambda event: self.salvar_medida("Generico"))

    def criar_interface(self):
        # 1. Painel de Arquivo
        self.frame_arquivo = ctk.CTkFrame(self)
        self.frame_arquivo.pack(pady=(15, 5), padx=10, fill="x")
        
        self.lbl_arq = ctk.CTkLabel(self.frame_arquivo, text="Ensaio:", font=("Arial", 12, "bold"))
        self.lbl_arq.pack(side="left", padx=10)

        self.entry_arquivo = ctk.CTkEntry(self.frame_arquivo, width=200, placeholder_text="Nome do arquivo")
        self.entry_arquivo.pack(side="left", padx=5)
        self.entry_arquivo.insert(0, "ensaio_metrologia") 
        
        self.btn_abrir = ctk.CTkButton(self.frame_arquivo, text="📂 Excel", command=self.abrir_tabela, width=80)
        self.btn_abrir.pack(side="right", padx=10)

        # 2. Painel de Conexão (AGORA COM SELETOR)
        self.frame_topo = ctk.CTkFrame(self)
        self.frame_topo.pack(pady=5, padx=10, fill="x")

        # Label simples
        self.lbl_porta = ctk.CTkLabel(self.frame_topo, text="Porta:", font=("Arial", 11))
        self.lbl_porta.pack(side="left", padx=(10, 2))

        # MENU DE ESCOLHA DA PORTA
        self.combo_portas = ctk.CTkComboBox(self.frame_topo, values=["Procurando..."], width=100)
        self.combo_portas.pack(side="left", padx=5)

        # Botão Atualizar Lista (ícone de refresh improvisado)
        self.btn_refresh = ctk.CTkButton(self.frame_topo, text="⟳", width=30, command=self.listar_portas_disponiveis)
        self.btn_refresh.pack(side="left", padx=2)

        self.btn_conexao = ctk.CTkButton(self.frame_topo, text="Conectar", command=self.alternar_conexao, fg_color="green", width=100)
        self.btn_conexao.pack(side="left", padx=10)
        
        self.btn_tarar = ctk.CTkButton(self.frame_topo, text="ZERAR", command=self.comando_tarar, fg_color="#555555", width=80)
        self.btn_tarar.pack(side="right", padx=10)
        self.btn_tarar.configure(state="disabled")

        # 3. Painel Central (Display)
        self.frame_display = ctk.CTkFrame(self)
        self.frame_display.pack(pady=10, padx=20, fill="both", expand=True)

        self.lbl_titulo = ctk.CTkLabel(self.frame_display, text="LEITURA REAL-TIME", font=("Arial", 14, "bold"), text_color="gray")
        self.lbl_titulo.pack(pady=(15, 0))

        self.lbl_peso = ctk.CTkLabel(self.frame_display, text="--- g", font=("Roboto", 70, "bold"))
        self.lbl_peso.pack(pady=10)
        
        self.lbl_status = ctk.CTkLabel(self.frame_display, text="Selecione a porta e conecte", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 4. Painel de Ação (ABBA)
        self.frame_acoes = ctk.CTkFrame(self)
        self.frame_acoes.pack(pady=10, padx=10, fill="x")
        self.frame_acoes.columnconfigure(0, weight=1)
        self.frame_acoes.columnconfigure(1, weight=1)

        self.btn_A = ctk.CTkButton(self.frame_acoes, text=" PADRÃO (A) ", height=60, 
                                   command=lambda: self.salvar_medida("Padrao (A)"), fg_color="#2980b9")
        self.btn_A.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        
        self.btn_B = ctk.CTkButton(self.frame_acoes, text=" CLIENTE (B) ", height=60, 
                                   command=lambda: self.salvar_medida("Cliente (B)"), fg_color="#e67e22")
        self.btn_B.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        self.btn_A.configure(state="disabled")
        self.btn_B.configure(state="disabled")

        # Log
        self.textbox_log = ctk.CTkTextbox(self, height=100)
        self.textbox_log.pack(pady=10, padx=10, fill="x")
        self.textbox_log.insert("0.0", "Bem-vindo. Selecione a porta COM acima.\n")

    # --- NOVA FUNÇÃO: Listar Portas ---
    def listar_portas_disponiveis(self):
        portas = serial.tools.list_ports.comports()
        lista_portas = []
        for p in portas:
            # p.device é "COM8", "COM10", etc.
            lista_portas.append(p.device)
        
        if not lista_portas:
            lista_portas = ["Nenhuma"]
        
        self.combo_portas.configure(values=lista_portas)
        self.combo_portas.set(lista_portas[0]) # Seleciona a primeira
        self.log(f"Portas encontradas: {', '.join(lista_portas)}\n")

    # --- LÓGICA GERAL ---
    def get_nome_arquivo(self):
        nome_pasta = "dados coletados"
        if not os.path.exists(nome_pasta):
            os.makedirs(nome_pasta)
        nome = self.entry_arquivo.get().strip() or "dados_sisaqui"
        return os.path.join(nome_pasta, nome + ".csv" if not nome.endswith(".csv") else nome)

    def verificar_csv(self):
        arquivo = self.get_nome_arquivo()
        if not os.path.exists(arquivo):
            try:
                with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(['Data', 'Hora', 'Peso (g)', 'Tipo'])
                self.log(f"Arquivo '{arquivo}' criado.")
            except PermissionError:
                messagebox.showerror("Erro", "Feche o arquivo Excel!")

    def abrir_tabela(self):
        arquivo = self.get_nome_arquivo()
        if os.path.exists(arquivo):
            os.startfile(arquivo)
        else:
            messagebox.showwarning("Aviso", "Arquivo ainda não existe.")

    def alternar_conexao(self):
        if self.ser and self.ser.is_open:
            self.monitorando = False
            time.sleep(0.5)
            self.ser.close()
            self.ser = None
            self.btn_conexao.configure(text="Conectar", fg_color="green")
            self.combo_portas.configure(state="normal") # Libera escolha de porta
            self.btn_refresh.configure(state="normal")
            self.lbl_status.configure(text="Desconectado", text_color="gray")
            self.lbl_peso.configure(text="--- g")
            self.entry_arquivo.configure(state="normal")
            self.btn_A.configure(state="disabled")
            self.btn_B.configure(state="disabled")
            self.btn_tarar.configure(state="disabled")
            self.log("Desconectado.")
        else:
            # PEGA A PORTA ESCOLHIDA NO MENU
            porta_escolhida = self.combo_portas.get()
            if porta_escolhida == "Nenhuma" or porta_escolhida == "Procurando...":
                messagebox.showwarning("Aviso", "Nenhuma porta selecionada.")
                return

            try:
                # Usa a variável porta_escolhida em vez de 'COM8' fixo
                self.ser = serial.Serial(porta_escolhida, 1200, bytesize=serial.SEVENBITS, parity=serial.PARITY_ODD, stopbits=1, timeout=0.5)
                self.monitorando = True
                threading.Thread(target=self.thread_monitoramento, daemon=True).start()
                
                self.btn_conexao.configure(text="Desconectar", fg_color="red")
                self.combo_portas.configure(state="disabled") # Trava mudança de porta
                self.btn_refresh.configure(state="disabled")
                
                self.lbl_status.configure(text=f"Monitorando na {porta_escolhida}...", text_color="#00FF00")
                self.entry_arquivo.configure(state="disabled")
                self.btn_A.configure(state="normal")
                self.btn_B.configure(state="normal")
                self.btn_tarar.configure(state="normal")
                self.verificar_csv()
                self.focus()
                self.log(f"Conectado na {porta_escolhida}!\n")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha na conexão com {porta_escolhida}:\n{e}")

    def thread_monitoramento(self):
        while self.monitorando and self.ser and self.ser.is_open:
            try:
                self.ser.reset_input_buffer()
                self.ser.write(b'\x1bP\r\n') 
                linha = self.ser.readline().decode('ascii', errors='ignore')
                match = re.search(r"[-+]?\s*\d+\.\d+", linha)
                if match:
                    peso_str = match.group().replace(" ", "")
                    peso_float = float(peso_str)
                    self.ultimo_peso_valido = peso_float 
                    self.lbl_peso.configure(text=f"{peso_float} g")
                time.sleep(0.2)
            except Exception as e:
                print(f"Erro no monitoramento: {e}")
                time.sleep(1)

    def comando_tarar(self):
        if self.ser:
            self.ser.write(b'\x1bf4_\r\n') 
            self.log("\nComando de TARA enviado.\n")
            time.sleep(0.5)

    def salvar_medida(self, tipo_amostra):
        if not self.monitorando or self.ultimo_peso_valido is None:
            self.log("⚠️ Aguarde leitura estável...")
            return

        arquivo = self.get_nome_arquivo()
        agora = datetime.now()
        
        try:
            with open(arquivo, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow([agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), str(self.ultimo_peso_valido).replace('.', ','), tipo_amostra])
            
            self.log(f"✅ Salvo: {self.ultimo_peso_valido}g como [{tipo_amostra}]")
            if "A" in tipo_amostra:
                self.flash_button(self.btn_A)
            elif "B" in tipo_amostra:
                self.flash_button(self.btn_B)
        except PermissionError:
            messagebox.showerror("Erro", "Excel aberto! Feche para salvar.")

    def flash_button(self, btn):
        cor_original = btn._fg_color
        btn.configure(fg_color="white", text_color="black")
        self.after(100, lambda: btn.configure(fg_color=cor_original, text_color="white"))
    
    def log(self, msg):
        self.textbox_log.insert("end", msg + "\n")
        self.textbox_log.see("end")

if __name__ == "__main__":
    app = AppBalanca()
    app.mainloop()