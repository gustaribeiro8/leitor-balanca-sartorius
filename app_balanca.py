import customtkinter as ctk
import serial
import time
import re
import csv
import threading
import os
from datetime import datetime
from tkinter import messagebox

# --- CONFIGURAÇÕES VISUAIS ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AppBalanca(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da Janela
        self.title("SOMA - Metrologia ABBA v4.0")
        self.geometry("700x600")
        self.resizable(False, False)

        # Variáveis de Controle
        self.ser = None
        self.monitorando = False # Controle da Thread
        self.ultimo_peso_valido = None # Guarda o valor para salvar instantaneamente
        
        # --- LAYOUT ---
        self.criar_interface()
        
        # Atalhos de teclado (ABBA)
        # O 'bind' conecta a tecla física a uma função
        self.bind('<a>', lambda event: self.salvar_medida("Padrão (A)"))
        self.bind('<A>', lambda event: self.salvar_medida("Padrão (A)")) # CapsLock
        self.bind('<b>', lambda event: self.salvar_medida("Objeto (B)"))
        self.bind('<B>', lambda event: self.salvar_medida("Objeto (B)")) # CapsLock
        
        # Mantive o Espaço como genérico (sem tipo especificado ou Padrão, você decide)
        self.bind('<space>', lambda event: self.salvar_medida("Genérico"))

    def criar_interface(self):
        # 1. Painel de Arquivo
        self.frame_arquivo = ctk.CTkFrame(self)
        self.frame_arquivo.pack(pady=(15, 5), padx=10, fill="x")
        
        self.lbl_arq = ctk.CTkLabel(self.frame_arquivo, text="Ensaio:", font=("Arial", 12, "bold"))
        self.lbl_arq.pack(side="left", padx=10)

        self.entry_arquivo = ctk.CTkEntry(self.frame_arquivo, width=200, placeholder_text="Nome do arquivo")
        self.entry_arquivo.pack(side="left", padx=5)
        self.entry_arquivo.insert(0, "calibracao_ABBA")
        
        self.btn_abrir = ctk.CTkButton(self.frame_arquivo, text="📂 Excel", command=self.abrir_tabela, width=80)
        self.btn_abrir.pack(side="right", padx=10)

        # 2. Painel de Conexão e Controle
        self.frame_topo = ctk.CTkFrame(self)
        self.frame_topo.pack(pady=5, padx=10, fill="x")

        self.btn_conexao = ctk.CTkButton(self.frame_topo, text="Conectar COM8", command=self.alternar_conexao, fg_color="green")
        self.btn_conexao.pack(side="left", padx=10)
        
        # Botões de Controle da Balança (Novidade)
        self.btn_tarar = ctk.CTkButton(self.frame_topo, text="ZERAR / TARAR", command=self.comando_tarar, fg_color="#555555", width=120)
        self.btn_tarar.pack(side="right", padx=10)
        self.btn_tarar.configure(state="disabled")

        # 3. Painel Central (Display Real-Time)
        self.frame_display = ctk.CTkFrame(self)
        self.frame_display.pack(pady=10, padx=20, fill="both", expand=True)

        self.lbl_titulo = ctk.CTkLabel(self.frame_display, text="LEITURA EM TEMPO REAL", font=("Arial", 14, "bold"), text_color="gray")
        self.lbl_titulo.pack(pady=(15, 0))

        # O numero grandão
        self.lbl_peso = ctk.CTkLabel(self.frame_display, text="--- g", font=("Roboto", 70, "bold"))
        self.lbl_peso.pack(pady=10)
        
        self.lbl_status = ctk.CTkLabel(self.frame_display, text="Desconectado", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 4. Painel de Ação (Botoes Gigantes para ABBA)
        self.frame_acoes = ctk.CTkFrame(self)
        self.frame_acoes.pack(pady=10, padx=10, fill="x")

        # Grid para dividir os botões
        self.frame_acoes.columnconfigure(0, weight=1)
        self.frame_acoes.columnconfigure(1, weight=1)

        self.btn_A = ctk.CTkButton(self.frame_acoes, text="Gravar PADRÃO (A)", height=60, 
                                   command=lambda: self.salvar_medida("Padrão (A)"), fg_color="#2980b9")
        self.btn_A.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        
        self.btn_B = ctk.CTkButton(self.frame_acoes, text="Gravar OBJETO (B)", height=60, 
                                   command=lambda: self.salvar_medida("Objeto (B)"), fg_color="#e67e22")
        self.btn_B.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        # Começam desativados
        self.btn_A.configure(state="disabled")
        self.btn_B.configure(state="disabled")

        # Área de Log
        self.textbox_log = ctk.CTkTextbox(self, height=100)
        self.textbox_log.pack(pady=10, padx=10, fill="x")
        self.textbox_log.insert("0.0", "Use teclas 'A' e 'B' do teclado para capturar.\n")

    # --- LÓGICA DE ARQUIVO ---
    def get_nome_arquivo(self):
        nome = self.entry_arquivo.get().strip() or "dados_abba"
        return nome + ".csv" if not nome.endswith(".csv") else nome

    def verificar_csv(self):
        arquivo = self.get_nome_arquivo()
        if not os.path.exists(arquivo):
            try:
                with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=';')
                    # NOVA COLUNA: TIPO
                    writer.writerow(['Data', 'Hora', 'Peso (g)', 'Tipo'])
                self.log(f"Arquivo '{arquivo}' criado com sucesso.")
            except PermissionError:
                messagebox.showerror("Erro", "Feche o arquivo Excel!")

    def abrir_tabela(self):
        arquivo = self.get_nome_arquivo()
        if os.path.exists(arquivo):
            os.startfile(arquivo)
        else:
            messagebox.showwarning("Aviso", "Arquivo ainda não existe.")

    # --- LÓGICA DE CONEXÃO E MONITORAMENTO ---
    def alternar_conexao(self):
        if self.ser and self.ser.is_open:
            self.monitorando = False # Para a thread
            time.sleep(0.5) # Dá tempo da thread morrer
            self.ser.close()
            self.ser = None
            
            # Reset Visual
            self.btn_conexao.configure(text="Conectar COM8", fg_color="green")
            self.lbl_status.configure(text="Desconectado", text_color="gray")
            self.lbl_peso.configure(text="--- g")
            self.entry_arquivo.configure(state="normal")
            self.btn_A.configure(state="disabled")
            self.btn_B.configure(state="disabled")
            self.btn_tarar.configure(state="disabled")
            self.log("Desconectado.")
        else:
            try:
                self.ser = serial.Serial('COM8', 1200, bytesize=serial.SEVENBITS, parity=serial.PARITY_ODD, stopbits=1, timeout=0.5)
                self.monitorando = True
                
                # Inicia Thread de Leitura Contínua
                threading.Thread(target=self.thread_monitoramento, daemon=True).start()
                
                # Atualiza Visual
                self.btn_conexao.configure(text="Desconectar", fg_color="red")
                self.lbl_status.configure(text="Monitorando em Tempo Real...", text_color="#00FF00")
                self.entry_arquivo.configure(state="disabled")
                self.btn_A.configure(state="normal")
                self.btn_B.configure(state="normal")
                self.btn_tarar.configure(state="normal")
                
                self.verificar_csv()
                self.focus() # Traz foco para a janela para pegar o teclado
                self.log("Conectado! Iniciando leitura contínua...")
            
            except Exception as e:
                messagebox.showerror("Erro", f"Falha na conexão:\n{e}")

    def thread_monitoramento(self):
        """ Esta função roda em paralelo O TEMPO TODO enquanto conectado """
        while self.monitorando and self.ser and self.ser.is_open:
            try:
                # 1. Limpa e Pede Peso
                self.ser.reset_input_buffer()
                self.ser.write(b'\x1bP\r\n') 
                
                # 2. Lê resposta
                linha = self.ser.readline().decode('ascii', errors='ignore')
                
                # 3. Regex para achar numero
                match = re.search(r"[-+]?\s*\d+\.\d+", linha)
                if match:
                    peso_str = match.group().replace(" ", "")
                    peso_float = float(peso_str)
                    
                    # Guarda na variável global da classe para salvar depois
                    self.ultimo_peso_valido = peso_float 
                    
                    # Atualiza a tela (Thread Safe no CustomTkinter)
                    self.lbl_peso.configure(text=f"{peso_float} g")
                
                time.sleep(0.2) # Taxa de atualização (aprox 5Hz)
                
            except Exception as e:
                print(f"Erro no monitoramento: {e}")
                time.sleep(1)

    # --- COMANDOS DA BALANÇA ---
    def comando_tarar(self):
        if self.ser:
            # Envia comando de Tarar/Zerar (Esc + f4 + _) - Verifique seu manual se é f3 ou f4
            # Na maioria é Esc + T ou Esc + f4
            self.ser.write(b'\x1bf4_\r\n') 
            self.log("Comando TARAR enviado.")
            time.sleep(0.5) # Pausa para a balança estabilizar

    # --- SALVAMENTO ---
    def salvar_medida(self, tipo_amostra):
        # Verifica se estamos conectados e se temos um peso válido lido recentemente
        if not self.monitorando or self.ultimo_peso_valido is None:
            self.log("⚠️ Aguarde leitura estável...")
            return

        arquivo = self.get_nome_arquivo()
        agora = datetime.now()
        
        try:
            with open(arquivo, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow([
                    agora.strftime("%d/%m/%Y"), 
                    agora.strftime("%H:%M:%S"), 
                    str(self.ultimo_peso_valido).replace('.', ','),
                    tipo_amostra # AQUI ENTRA O "A" ou "B"
                ])
            
            # Feedback Visual Rápido
            self.log(f"✅ Salvo: {self.ultimo_peso_valido}g como [{tipo_amostra}]")
            
            # Pisca a cor do botão correspondente
            if "A" in tipo_amostra:
                self.flash_button(self.btn_A)
            elif "B" in tipo_amostra:
                self.flash_button(self.btn_B)

        except PermissionError:
            messagebox.showerror("Erro", "Excel aberto! Feche para salvar.")

    def flash_button(self, btn):
        # Efeitinho visual para confirmar o clique
        cor_original = btn._fg_color
        btn.configure(fg_color="white", text_color="black")
        self.after(100, lambda: btn.configure(fg_color=cor_original, text_color="white"))
    
    def log(self, msg):
        self.textbox_log.insert("0.0", msg + "\n") # Insere no topo

if __name__ == "__main__":
    app = AppBalanca()
    app.mainloop()