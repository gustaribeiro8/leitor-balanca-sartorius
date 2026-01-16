import customtkinter as ctk
import serial
import time
import re
import csv
import threading
from datetime import datetime
from tkinter import messagebox

# --- CONFIGURAÇÕES VISUAIS ---
ctk.set_appearance_mode("Dark")  # Modo escuro
ctk.set_default_color_theme("blue")  # Tema azul

class AppBalanca(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da Janela
        self.title("SOMA - Leitor Sartorius v1.0")
        self.geometry("600x450")
        self.resizable(False, False)

        # Variáveis de Controle
        self.ser = None
        self.lendo_dados = False
        self.arquivo_saida = 'dados_coletados.csv'

        # --- LAYOUT ---
        self.criar_interface()
        
        # Bind da tecla Espaço para capturar
        self.bind('<space>', self.comando_capturar_tecla)

    def criar_interface(self):
        # 1. Painel Superior (Conexão)
        self.frame_topo = ctk.CTkFrame(self)
        self.frame_topo.pack(pady=10, padx=10, fill="x")

        self.btn_conectar = ctk.CTkButton(self.frame_topo, text="Conectar COM8", command=self.conectar_serial, fg_color="green")
        self.btn_conectar.pack(side="left", padx=10)

        self.lbl_status = ctk.CTkLabel(self.frame_topo, text="Status: Desconectado", text_color="gray")
        self.lbl_status.pack(side="left", padx=10)

        # 2. Painel Central (Mostrador de Peso)
        self.frame_display = ctk.CTkFrame(self)
        self.frame_display.pack(pady=20, padx=20, fill="both", expand=True)

        self.lbl_titulo_peso = ctk.CTkLabel(self.frame_display, text="INDICAÇÃO ATUAL", font=("Arial", 16))
        self.lbl_titulo_peso.pack(pady=(20, 5))

        # O numero grandão
        self.lbl_peso = ctk.CTkLabel(self.frame_display, text="--- g", font=("Roboto", 60, "bold"))
        self.lbl_peso.pack(pady=10)

        # 3. Painel Inferior (Botões de Ação)
        self.frame_acoes = ctk.CTkFrame(self)
        self.frame_acoes.pack(pady=10, padx=10, fill="x")

        self.btn_capturar = ctk.CTkButton(self.frame_acoes, text="CAPTURAR (Espaço)", height=50, command=self.capturar_peso)
        self.btn_capturar.pack(side="left", padx=10, fill="x", expand=True)
        self.btn_capturar.configure(state="disabled") # Começa desativado

        # Área de log (Histórico rápido)
        self.textbox_log = ctk.CTkTextbox(self, height=100)
        self.textbox_log.pack(pady=10, padx=10, fill="x")
        self.textbox_log.insert("0.0", "Sistema iniciado. Conecte a balança.\n")

    # --- LÓGICA DE SERIAL (Back-end) ---
    def conectar_serial(self):
        try:
            # Configurações iguais ao seu script anterior
            self.ser = serial.Serial(
                port='COM8', baudrate=1200, bytesize=serial.SEVENBITS,
                parity=serial.PARITY_ODD, stopbits=serial.STOPBITS_ONE,
                timeout=1, rtscts=False
            )
            self.lbl_status.configure(text="Status: CONECTADO", text_color="#00FF00")
            self.btn_conectar.configure(state="disabled", fg_color="gray")
            self.btn_capturar.configure(state="normal")
            self.log("Porta COM8 aberta com sucesso.")
            
            # Garante que o arquivo CSV existe
            self.inicializar_csv()

        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível conectar:\n{e}")

    def inicializar_csv(self):
        try:
            with open(self.arquivo_saida, 'x', newline='') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['Data', 'Hora', 'Peso (g)'])
        except FileExistsError:
            pass

    def comando_capturar_tecla(self, event):
        if self.ser and self.ser.is_open:
            self.capturar_peso()

    def capturar_peso(self):
        # Roda em Thread separada para não travar a tela
        threading.Thread(target=self._rotina_leitura).start()

    def _rotina_leitura(self):
        self.ser.reset_input_buffer()
        
        # Envia comando PRINT
        self.ser.write(b'\x1bP\r\n')
        time.sleep(0.2)
        
        # Tenta ler (Timeout manual)
        tentativas = 20
        peso_lido = None
        
        for _ in range(tentativas):
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

        # Atualiza a interface (precisa ser feito na thread principal, mas o ctk lida bem com isso)
        if peso_lido is not None:
            self.salvar_e_mostrar(peso_lido)
        else:
            self.log("Erro: Sem resposta da balança.")

    def salvar_e_mostrar(self, peso):
        agora = datetime.now()
        data_str = agora.strftime("%d/%m/%Y")
        hora_str = agora.strftime("%H:%M:%S")
        
        # Salva no CSV
        with open(self.arquivo_saida, 'a', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow([data_str, hora_str, str(peso).replace('.', ',')])
        
        # Atualiza GUI
        self.lbl_peso.configure(text=f"{peso} g")
        self.log(f"Capturado: {peso} g em {hora_str}")
        
        # Feedback visual (piscar botão)
        self.btn_capturar.configure(fg_color="#3B8ED0") # Azul padrão

    def log(self, mensagem):
        self.textbox_log.insert("end", mensagem + "\n")
        self.textbox_log.see("end")

if __name__ == "__main__":
    app = AppBalanca()
    app.mainloop()