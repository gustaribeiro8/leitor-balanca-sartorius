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
        return os.path.join(nome_pasta, nome + ".csv" if not nome.endswith(".csv") else nome)

    def verificar_csv(self):
        arquivo = self.get_nome_arquivo()
        if not os.path.exists(arquivo):
            try:
                with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=';')
                    # Cabecalho com espaco para futuros tipos
                    writer.writerow(['Data', 'Hora', '', 'Padrao (A)', '', 'Cliente (B)', '', 'Generico', ''])
                    writer.writerow(['', '', '', 'Peso (g)', 'Hora', 'Peso (g)', 'Hora', 'Peso (g)', 'Hora'])
                self.log(f"Arquivo '{arquivo}' criado.")
            except PermissionError:
                messagebox.showerror("Erro", "Feche o arquivo Excel!")

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
        if self.ser:
            self.ser.write(b'\x1bf4_\r\n') 
            self.log("\nComando de TARA enviado.\n")
            time.sleep(0.5)

    def salvar_medida(self, tipo_amostra):
        if not self.monitorando or self.ultimo_peso_valido is None:
            self.log("⚠️ Aguarde leitura estavel...")
            return

        arquivo = self.get_nome_arquivo()
        try:
            # Le todos os dados existentes
            dados_por_tipo = {'Padrao (A)': [], 'Cliente (B)': [], 'Generico': []}
            datas = []
            
            if os.path.exists(arquivo):
                with open(arquivo, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f, delimiter=';')
                    linhas = list(reader)
                    
                    # Pula os 2 primeiros cabecalhos
                    if len(linhas) > 2:
                        for linha in linhas[2:]:
                            if len(linha) >= 10:
                                data = linha[0]
                                if data and data not in datas and data != 'Data':
                                    datas.append(data)
                                
                                # Coleta pesos tipo A (indice 3, 4)
                                if linha[3]:
                                    dados_por_tipo['Padrao (A)'].append([linha[3], linha[4]])
                                
                                # Coleta pesos tipo B (indice 6, 7)
                                if linha[6]:
                                    dados_por_tipo['Cliente (B)'].append([linha[6], linha[7]])
                                
                                # Coleta pesos genericos (indice 9, 10)
                                if len(linha) > 9 and linha[9]:
                                    dados_por_tipo['Generico'].append([linha[9], linha[10] if len(linha) > 10 else ''])
            
            # Adiciona nova medida
            data_str = agora.strftime("%d/%m/%Y")
            hora_str = agora.strftime("%H:%M:%S")
            peso_str = str(self.ultimo_peso_valido).replace('.', ',')
            
            if data_str not in datas:
                datas.append(data_str)
            
            dados_por_tipo[tipo_amostra].append([peso_str, hora_str])
            
            # Reescreve o arquivo com os novos dados reorganizados
            with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                # Cabecalhos
                writer.writerow(['Data', 'Hora', '', 'Padrao (A)', '', '', 'Cliente (B)', '', '', 'Generico', ''])
                writer.writerow(['', '', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora'])
                
                # Dados organizados por linha/tipo
                max_medidas = max(len(dados_por_tipo['Padrao (A)']), 
                                 len(dados_por_tipo['Cliente (B)']), 
                                 len(dados_por_tipo['Generico']))
                
                contador_linhas = 0
                for i in range(max_medidas):
                    # Quebra de linha a cada 4 medidas
                    if contador_linhas > 0 and contador_linhas % 4 == 0:
                        writer.writerow([])
                    
                    linha = []
                    
                    # Data e Hora (apenas na primeira linha)
                    if i == 0:
                        linha.extend([datas[-1] if datas else '', hora_str])
                    else:
                        linha.extend(['', ''])
                    
                    linha.append('')  # Coluna vazia de separacao
                    
                    # Tipo A
                    if i < len(dados_por_tipo['Padrao (A)']):
                        linha.extend(dados_por_tipo['Padrao (A)'][i])
                    else:
                        linha.extend(['', ''])
                    
                    linha.append('')  # Coluna vazia de separacao
                    
                    # Tipo B
                    if i < len(dados_por_tipo['Cliente (B)']):
                        linha.extend(dados_por_tipo['Cliente (B)'][i])
                    else:
                        linha.extend(['', ''])
                    
                    linha.append('')  # Coluna vazia de separacao
                    
                    # Tipo Generico
                    if i < len(dados_por_tipo['Generico']):
                        linha.extend(dados_por_tipo['Generico'][i])
                    else:
                        linha.extend(['', ''])
                    
                    writer.writerow(linha)
                    contador_linhas += 1
                
                # Quebra de linha e estatisticas
                writer.writerow([])
                writer.writerow(['Estatisticas:'])
                writer.writerow(['Tipo', 'Quantidade', 'Media (g)', 'Minimo (g)', 'Maximo (g)'])
                
                for tipo in ['Padrao (A)', 'Cliente (B)', 'Generico']:
                    if dados_por_tipo[tipo]:
                        pesos = [float(m[0].replace(',', '.')) for m in dados_por_tipo[tipo]]
                        media = sum(pesos) / len(pesos)
                        minimo = min(pesos)
                        maximo = max(pesos)
                        writer.writerow([tipo, len(pesos), f"{media:.2f}".replace('.', ','), 
                                        f"{minimo:.2f}".replace('.', ','), 
                                        f"{maximo:.2f}".replace('.', ',')])
            
            self.log(f"OK Salvo: {self.ultimo_peso_valido}g como [{tipo_amostra}]")
            if "A" in tipo_amostra:
                self.flash_button(self.btn_A)
            elif "B" in tipo_amostra:
                self.flash_button(self.btn_B)
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