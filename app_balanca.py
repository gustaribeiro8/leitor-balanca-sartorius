import customtkinter as ctk
import serial
import time
import re
import csv
import threading
import os  # Necessário para abrir o arquivo no Windows
from datetime import datetime
from tkinter import messagebox

# --- CONFIGURAÇÕES VISUAIS ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AppBalanca(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da Janela
        self.title("SOMA - Leitor Sartorius v2.0")
        self.geometry("600x550") # Aumentei um pouco a altura
        self.resizable(False, False)

        # Variáveis
        self.ser = None
        
        # --- LAYOUT ---
        self.criar_interface()
        
        # Atalho de teclado
        self.bind('<space>', self.comando_capturar_tecla)

    def criar_interface(self):
        # 1. Painel de Arquivo (NOVO: Onde você dá o nome)
        self.frame_arquivo = ctk.CTkFrame(self)
        self.frame_arquivo.pack(pady=(15, 5), padx=10, fill="x")
        
        self.lbl_arq = ctk.CTkLabel(self.frame_arquivo, text="Nome da Tabela:", font=("Arial", 12, "bold"))
        self.lbl_arq.pack(side="left", padx=10)

        self.entry_arquivo = ctk.CTkEntry(self.frame_arquivo, width=250, placeholder_text="Ex: Ensaio_01")
        self.entry_arquivo.pack(side="left", padx=5)
        self.entry_arquivo.insert(0, "dados_coletados") # Nome padrão
        
        # Botão para abrir o Excel direto
        self.btn_abrir_tabela = ctk.CTkButton(self.frame_arquivo, text="📂 Abrir/Ver Tabela", 
                                              command=self.abrir_tabela, fg_color="#D35400", width=120)
        self.btn_abrir_tabela.pack(side="right", padx=10)

        # 2. Painel de Conexão
        self.frame_topo = ctk.CTkFrame(self)
        self.frame_topo.pack(pady=5, padx=10, fill="x")

        self.btn_conectar = ctk.CTkButton(self.frame_topo, text="Conectar COM8", command=self.conectar_serial, fg_color="green")
        self.btn_conectar.pack(side="left", padx=10)

        self.lbl_status = ctk.CTkLabel(self.frame_topo, text="Status: Desconectado", text_color="gray")
        self.lbl_status.pack(side="left", padx=10)

        # 3. Painel Central (Peso)
        self.frame_display = ctk.CTkFrame(self)
        self.frame_display.pack(pady=10, padx=20, fill="both", expand=True)

        self.lbl_titulo_peso = ctk.CTkLabel(self.frame_display, text="INDICAÇÃO ATUAL", font=("Arial", 16))
        self.lbl_titulo_peso.pack(pady=(20, 5))

        self.lbl_peso = ctk.CTkLabel(self.frame_display, text="--- g", font=("Roboto", 60, "bold"))
        self.lbl_peso.pack(pady=10)

        # 4. Painel de Ação
        self.frame_acoes = ctk.CTkFrame(self)
        self.frame_acoes.pack(pady=10, padx=10, fill="x")

        self.btn_capturar = ctk.CTkButton(self.frame_acoes, text="CAPTURAR (Espaço)", height=50, command=self.capturar_peso)
        self.btn_capturar.pack(side="left", padx=10, fill="x", expand=True)
        self.btn_capturar.configure(state="disabled")

        # Área de Log
        self.textbox_log = ctk.CTkTextbox(self, height=80)
        self.textbox_log.pack(pady=10, padx=10, fill="x")
        self.textbox_log.insert("0.0", "v2.0 Iniciada. Defina o nome do arquivo e conecte.\n")

    # --- FUNÇÕES AUXILIARES ---
    def get_nome_arquivo(self):
        nome = self.entry_arquivo.get().strip()
        if not nome: # Se estiver vazio
            nome = "dados_coletados"
        if not nome.endswith(".csv"):
            nome += ".csv"
        return nome

    def abrir_tabela(self):
        """ Abre o arquivo CSV no Excel/programa padrão """
        arquivo = self.get_nome_arquivo()
        if os.path.exists(arquivo):
            try:
                os.startfile(arquivo) # Comando do Windows para abrir arquivo
                self.log(f"Abrindo '{arquivo}'...")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir:\n{e}")
        else:
            messagebox.showwarning("Aviso", "O arquivo ainda não existe.\nCapture pelo menos um peso primeiro!")

    # --- LÓGICA DE SERIAL ---
    def conectar_serial(self):
        try:
            self.ser = serial.Serial(
                port='COM8', baudrate=1200, bytesize=serial.SEVENBITS,
                parity=serial.PARITY_ODD, stopbits=serial.STOPBITS_ONE,
                timeout=1, rtscts=False
            )
            self.lbl_status.configure(text="Status: CONECTADO", text_color="#00FF00")
            self.btn_conectar.configure(state="disabled", fg_color="gray")
            self.btn_capturar.configure(state="normal")
            self.log("Conectado na COM8.")
            
            # Cria o cabeçalho se o arquivo for novo
            self.verificar_csv()

        except Exception as e:
            messagebox.showerror("Erro de Conexão", f"Verifique se o cabo está conectado.\nErro: {e}")

    def verificar_csv(self):
        arquivo = self.get_nome_arquivo()
        if not os.path.exists(arquivo):
            try:
                with open(arquivo, 'w', newline='') as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(['Data', 'Hora', 'Peso (g)'])
                self.log(f"Arquivo '{arquivo}' criado.")
            except PermissionError:
                messagebox.showerror("Erro", f"Feche o arquivo '{arquivo}' no Excel para continuar!")

    def comando_capturar_tecla(self, event):
        if self.ser and self.ser.is_open:
            self.capturar_peso()

    def capturar_peso(self):
        threading.Thread(target=self._rotina_leitura).start()

    def _rotina_leitura(self):
        # Feedback visual que está lendo
        self.btn_capturar.configure(fg_color="#E67E22", text="Lendo...") 
        
        self.ser.reset_input_buffer()
        self.ser.write(b'\x1bP\r\n')
        time.sleep(0.2)
        
        peso_lido = None
        for _ in range(20):
            if self.ser.in_waiting:
                try:
                    raw = self.ser.readline().decode('ascii', errors='ignore')
                    match = re.search(r"[-+]?\s*\d+\.\d+", raw)
                    if match:
                        peso_lido = float(match.group().replace(" ", ""))
                        break
                except:
                    pass
            time.sleep(0.1)

        if peso_lido is not None:
            self.salvar_e_mostrar(peso_lido)
        else:
            self.log("Erro: Sem resposta.")
        
        # Volta a cor do botão
        self.btn_capturar.configure(fg_color="#3B8ED0", text="CAPTURAR (Espaço)")

    def salvar_e_mostrar(self, peso):
        arquivo = self.get_nome_arquivo()
        agora = datetime.now()
        
        try:
            with open(arquivo, 'a', newline='') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow([
                    agora.strftime("%d/%m/%Y"), 
                    agora.strftime("%H:%M:%S"), 
                    str(peso).replace('.', ',')
                ])
            
            self.lbl_peso.configure(text=f"{peso} g")
            self.log(f"Salvo em '{arquivo}': {peso} g")
            
        except PermissionError:
            # AQUI ESTÁ O PULO DO GATO: Se o Excel estiver aberto, ele avisa!
            self.after(0, lambda: messagebox.showerror(
                "Arquivo Aberto!", 
                f"O Excel está segurando o arquivo '{arquivo}'.\n\nFeche a janela do Excel e tente capturar de novo."
            ))

    def log(self, msg):
        self.textbox_log.insert("end", msg + "\n")
        self.textbox_log.see("end")

if __name__ == "__main__":
    app = AppBalanca()
    app.mainloop()