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
from datetime import datetime

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

        messagebox.showinfo("Atenção", "Certifique-se de que a balança está ligada, estável e pronta para a conexão.")

        self.title("SISAQUI - Modo Planilha Livre v8.2") 
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
        self.bind('<space>', lambda event: self.capturar_coluna("G"))
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def criar_interface(self):
        # 1. Arquivo
        self.frame_arquivo = ctk.CTkFrame(self)
        self.frame_arquivo.pack(pady=(15, 5), padx=10, fill="x")
        
        self.lbl_arq = ctk.CTkLabel(self.frame_arquivo, text="Nome do Ensaio:", font=("Arial", 12, "bold"))
        self.lbl_arq.pack(side="left", padx=10)

        self.entry_arquivo = ctk.CTkEntry(self.frame_arquivo, width=220)
        self.entry_arquivo.pack(side="left", padx=5)
        self.entry_arquivo.insert(0, f"ensaio_{datetime.now().strftime('%Y-%m-%d')}") 
        
        self.btn_abrir = ctk.CTkButton(self.frame_arquivo, text="📂 Abrir no Excel", command=self.abrir_tabela, width=120)
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
        
        self.btn_tarar = ctk.CTkButton(self.frame_topo, text="ZERAR BALANÇA", command=self.comando_tarar, fg_color="#555555", width=120)
        self.btn_tarar.pack(side="right", padx=10)
        self.btn_tarar.configure(state="disabled")

        # 3. Display
        self.frame_display = ctk.CTkFrame(self)
        self.frame_display.pack(pady=10, padx=20, fill="both", expand=True)

        self.lbl_titulo = ctk.CTkLabel(self.frame_display, text="Leitura da Balança", font=("Arial", 14, "bold"), text_color="gray")
        self.lbl_titulo.pack(pady=(15, 0))

        self.lbl_peso = ctk.CTkLabel(self.frame_display, text="--- g", font=("Roboto", 70, "bold"))
        self.lbl_peso.pack(pady=10)
        
        self.lbl_status = ctk.CTkLabel(self.frame_display, text="Aguardando conexão...", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 4. Ações e Contadores
        self.frame_acoes = ctk.CTkFrame(self)
        self.frame_acoes.pack(pady=10, padx=10, fill="x")
        self.frame_acoes.columnconfigure(0, weight=1)
        self.frame_acoes.columnconfigure(1, weight=1)

        # Botão A
        self.btn_A = ctk.CTkButton(self.frame_acoes, text="Capturar para Coluna A", height=50, 
                                   command=lambda: self.capturar_coluna("A"), fg_color="#2980b9")
        self.btn_A.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        # Contador A
        self.lbl_count_A = ctk.CTkLabel(self.frame_acoes, text="Nº de Amostras A: 0", font=("Arial", 16, "bold"))
        self.lbl_count_A.grid(row=1, column=0, pady=5)

        # Botão B
        self.btn_B = ctk.CTkButton(self.frame_acoes, text="Capturar para Coluna B", height=50, 
                                   command=lambda: self.capturar_coluna("B"), fg_color="#e67e22")
        self.btn_B.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Contador B
        self.lbl_count_B = ctk.CTkLabel(self.frame_acoes, text="Nº de Amostras B: 0", font=("Arial", 16, "bold"))
        self.lbl_count_B.grid(row=1, column=1, pady=5)

        self.btn_A.configure(state="disabled")
        self.btn_B.configure(state="disabled")

        # Log (maior)
        self.textbox_log = ctk.CTkTextbox(self, height=120)
        self.textbox_log.pack(pady=10, padx=10, fill="x")
        self.textbox_log.insert("end", "Dica: As colunas são independentes. Você pode capturar várias amostras 'A' e depois várias 'B'.\n")

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

    def verificar_csv(self):
        arquivo = self.get_nome_arquivo()
        if not os.path.exists(arquivo):
            try:
                with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=';')
                    # Cabecalho com espaco para futuros tipos (formato com 11 colunas)
                    writer.writerow(['Data', 'Hora', '', 'Padrao (A)', '', '', 'Cliente (B)', '', '', 'Generico', ''])
                    writer.writerow(['', '', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora'])
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
                # Abre a porta serial
                self.ser = serial.Serial(porta, 1200, bytesize=7, parity='O', stopbits=1, timeout=0.5)

                # Testa resposta ao comando PRINT para detectar possivel erro 30
                try:
                    erro30, resp = self.verificar_erro_30()
                except Exception:
                    erro30, resp = (False, '')

                if erro30:
                    msg = (f"ATENÇÃO: Possível 'Erro 30' detectado na balança.\n\n"
                           "Este erro impede a comunicação e a leitura de pesos.\n\n"
                           "Para resolver:\n"
                           "1. Pressione o botão 📄 PRINT (ou ESC) no painel da balança.\n"
                           "2. Se o erro persistir, consulte o manual do aplicativo (manual_app_balanca.pdf) para mais instruções.\n\n"
                           f"Detalhes: {resp}")
                    messagebox.showwarning("Erro 30 Detectado", msg)
                    try:
                        self.ser.close()
                    except:
                        pass
                    self.ser = None
                    self.log(f"Conexão cancelada: erro 30 detectado na {porta}.")
                    return

                # Se o usuario optar por prosseguir (ou nao houve erro), inicializa monitoramento
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
                # Falha na abertura da porta: pode ser erro 30 ou outro problema
                msg = (f"Não foi possível conectar à porta {porta}.\n\n"
                       "Possíveis causas:\n"
                       "1. A balança não está conectada ao computador.\n"
                       "2. Outro software está usando a mesma porta.\n"
                       "3. Ocorreu um 'Erro 30' interno na balança.\n\n"
                       "Solução para 'Erro 30':\n"
                       "1. Aperte o botão 📄 PRINT (ou ESC) na balança.\n"
                       "2. Consulte o manual do aplicativo para mais detalhes.\n\n"
                       f"Erro técnico: {e}")
                messagebox.showerror("Erro de Conexão", msg)

    def thread_monitoramento(self):
        while self.monitorando and self.ser and self.ser.is_open:
            try:
                self.ser.reset_input_buffer()
                self.ser.write(b'\x1bP\r\n') 
                linha = self.ser.readline().decode('ascii', errors='ignore')

                if not self.monitorando: # Checa novamente caso o usuario tenha desconectado
                    break

                match = re.search(r"[-+]?\s*\d+\.\d+", linha)
                if match:
                    peso_str = match.group().replace(" ", "")
                    self.ultimo_peso_valido = float(peso_str)
                    # Formata para 4 casas decimais, que é comum em balanças analíticas
                    self.lbl_peso.configure(text=f"{self.ultimo_peso_valido:.4f} g")
                
                time.sleep(0.2)
            except (serial.SerialException, OSError) as e:
                self.log(f"ERRO: A porta serial foi desconectada ou falhou: {e}")
                self.after(0, self.handle_connection_loss)
                break # Encerra o loop de monitoramento
            except Exception as e:
                self.log(f"ERRO inesperado no monitoramento: {e}")
                time.sleep(1) # Espera um pouco antes de tentar de novo

    def verificar_erro_30(self, timeout=3.0):
        """Envia comando PRINT e verifica se a balanca responde com erro 30 ou nao consegue fornecer leitura.
        Retorna (True, resposta_texto) se detectar erro 30 ou falha de leitura, caso contrario (False, '')
        """
        if not self.ser or not self.ser.is_open:
            return (False, '')

        # Limpa buffer e envia comando de solicitacao de leitura
        try:
            self.ser.reset_input_buffer()
            self.ser.write(b'\x1bP\r\n')
        except Exception:
            return (False, '')

        fim = time.time() + timeout
        linhas = []
        peso_lido = False
        
        while time.time() < fim:
            try:
                if self.ser.in_waiting > 0:
                    raw = self.ser.readline()
                    try:
                        texto = raw.decode('ascii', errors='ignore').strip()
                    except:
                        texto = ''
                    if texto:
                        linhas.append(texto)
                        
                        # Procura padroes de erro 30 explicitamente
                        if re.search(r'(?i)(err(or)?)[[\s:\-]*30', texto) or re.search(r'(?i)\berr\b.*30', texto) or (texto.isdigit() and int(texto) == 30):
                            return (True, f"Codigo 30 detectado: {texto}")
                        
                        # Procura por peso valido (numero com ponto decimal)
                        if re.search(r"[-+]?\s*\d+\.\d+", texto):
                            peso_lido = True
                            return (False, '')  # Leitura OK, sem erro
            except:
                break
        
        # Se saiu do loop sem ler um peso valido, eh indicio de erro 30
        if not peso_lido:
            resposta = "\\n".join(linhas) if linhas else "(sem resposta da balanca)"
            return (True, f"Balanca nao forneceu leitura valida. Posivel erro 30. Respostas: {resposta}")
        
        return (False, '')
    def comando_tarar(self):
        if self.ser:
            self.ser.write(b'\x1bf4_\r\n') 
            self.log("\nComando de TARA enviado.\n")
            time.sleep(0.5)

    def capturar_coluna(self, coluna):
        """Captura medida para a coluna A, B ou generica.
        Converte a letra em tipo_amostra e delega para salvar_medida().
        """
        if coluna == 'A':
            tipo = 'Padrao (A)'
        elif coluna == 'B':
            tipo = 'Cliente (B)'
        else:
            tipo = 'Generico'

        self.log(f"Capturando coluna {coluna} -> {tipo}...")
        # Realiza a gravação e atualiza os contadores na interface
        self.salvar_medida(tipo)
        # Atualiza contadores visuais com as variaveis internas (se existirem)
        try:
            if hasattr(self, 'cursor_A') and hasattr(self, 'cursor_B'):
                self.lbl_count_A.configure(text=f"Nº de Amostras A: {self.cursor_A}")
                self.lbl_count_B.configure(text=f"Nº de Amostras B: {self.cursor_B}")
        except Exception:
            pass

    def salvar_medida(self, tipo_amostra):
        if not self.monitorando or self.ultimo_peso_valido is None:
            self.log("⚠️ Aguarde uma leitura estável...")
            return

        arquivo = self.get_nome_arquivo()
        agora = datetime.now()
        
        # --- Leitura e Preparação dos Dados ---
        dados_por_tipo = {'Padrao (A)': [], 'Cliente (B)': [], 'Generico': []}
        datas = []
        
        try:
            if os.path.exists(arquivo):
                with open(arquivo, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f, delimiter=';')
                    linhas = list(reader)
                    if len(linhas) > 2:
                        for linha in linhas[2:]:
                            if len(linha) >= 10:
                                if linha[0] and linha[0] != 'Data' and linha[0] not in datas: datas.append(linha[0])
                                if linha[3]: dados_por_tipo['Padrao (A)'].append([linha[3], linha[4]])
                                if linha[6]: dados_por_tipo['Cliente (B)'].append([linha[6], linha[7]])
                                if len(linha) > 9 and linha[9]: dados_por_tipo['Generico'].append([linha[9], linha[10] if len(linha) > 10 else ''])

            # Adiciona nova medida
            data_str = agora.strftime("%d/%m/%Y")
            if data_str not in datas: datas.append(data_str)
            
            dados_por_tipo[tipo_amostra].append([str(self.ultimo_peso_valido).replace('.', ','), agora.strftime("%H:%M:%S")])

            # Atualiza contadores
            self.cursor_A = len(dados_por_tipo.get('Padrao (A)', []))
            self.cursor_B = len(dados_por_tipo.get('Cliente (B)', []))
            self.lbl_count_A.configure(text=f"Nº de Amostras A: {self.cursor_A}")
            self.lbl_count_B.configure(text=f"Nº de Amostras B: {self.cursor_B}")

            # --- Escrita Atômica ---
            arquivo_tmp = arquivo + ".tmp"
            with open(arquivo_tmp, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['Data', 'Hora', '', 'Padrao (A)', '', '', 'Cliente (B)', '', '', 'Generico', ''])
                writer.writerow(['', '', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora'])
                
                max_medidas = max(len(v) for v in dados_por_tipo.values())
                for i in range(max_medidas):
                    if i > 0 and i % 4 == 0: writer.writerow([])
                    
                    linha_csv = [datas[-1] if datas and i==0 else '', agora.strftime("%H:%M:%S") if i==0 else '', '']
                    linha_csv.extend(dados_por_tipo['Padrao (A)'][i] if i < len(dados_por_tipo['Padrao (A)']) else ['', ''])
                    linha_csv.append('')
                    linha_csv.extend(dados_por_tipo['Cliente (B)'][i] if i < len(dados_por_tipo['Cliente (B)']) else ['', ''])
                    linha_csv.append('')
                    linha_csv.extend(dados_por_tipo['Generico'][i] if i < len(dados_por_tipo['Generico']) else ['', ''])
                    writer.writerow(linha_csv)
                
                writer.writerow([])
                writer.writerow(['Estatisticas:'])
                writer.writerow(['Tipo', 'Quantidade', 'Media (g)', 'Minimo (g)', 'Maximo (g)'])
                
                for tipo in ['Padrao (A)', 'Cliente (B)', 'Generico']:
                    if dados_por_tipo[tipo]:
                        pesos = [float(m[0].replace(',', '.')) for m in dados_por_tipo[tipo]]
                        stats = [tipo, len(pesos), f"{sum(pesos) / len(pesos):.4f}".replace('.', ','), f"{min(pesos):.4f}".replace('.', ','), f"{max(pesos):.4f}".replace('.', ',')]
                        writer.writerow(stats)
            
            os.replace(arquivo_tmp, arquivo)

            self.log(f"Salvo: {self.ultimo_peso_valido:.4f}g como [{tipo_amostra}]")
            if "A" in tipo_amostra: self.flash_button(self.btn_A)
            elif "B" in tipo_amostra: self.flash_button(self.btn_B)

        except PermissionError:
            messagebox.showerror("Erro de Permissão", f"Não foi possível salvar o arquivo '{os.path.basename(arquivo)}'.\n\nVerifique se ele não está aberto em outro programa (como o Excel) e tente novamente.")
            if os.path.exists(arquivo + ".tmp"): os.remove(arquivo + ".tmp")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar", f"Ocorreu um erro inesperado ao salvar o arquivo:\n\n{e}")
            self.log(f"ERRO ao salvar medida: {e}")
            if os.path.exists(arquivo + ".tmp"): os.remove(arquivo + ".tmp")

    def flash_button(self, btn):
        c = btn._fg_color
        btn.configure(fg_color="white", text_color="black")
        self.after(100, lambda: btn.configure(fg_color=c, text_color="white"))

    def log(self, msg):
        # Insere novas mensagens ao final (ordem cronologica: cima -> baixo)
        try:
            self.textbox_log.insert("end", msg + "\n")
            self.textbox_log.see("end")
        except Exception:
            pass

    def handle_connection_loss(self):
        """Gerencia a perda de conexão, atualizando a UI para o estado desconectado."""
        self.monitorando = False
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
            except Exception as e:
                self.log(f"Erro ao fechar porta serial: {e}")
        self.ser = None
        
        self.btn_conexao.configure(text="Conectar", fg_color="green")
        self.combo_portas.configure(state="normal")
        self.entry_arquivo.configure(state="normal")
        self.btn_A.configure(state="disabled")
        self.btn_B.configure(state="disabled")
        self.btn_tarar.configure(state="disabled")
        
        self.lbl_status.configure(text="Conexão perdida. Reconecte.", text_color="orange")
        self.lbl_peso.configure(text="--- g")
        self.log("‼️ Conexão com a balança foi perdida.")
        messagebox.showwarning("Conexão Perdida", "A comunicação com a balança foi interrompida. Verifique o cabo e reconecte.")

    def on_closing(self):
        self.monitorando = False
        if self.ser and self.ser.is_open: self.ser.close()
        self.destroy()

if __name__ == "__main__":
    app = AppBalanca()
    app.mainloop()