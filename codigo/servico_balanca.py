import serial
import serial.tools.list_ports
import time
import re
import threading

class ServicoBalanca:
    def __init__(self, on_peso_update, on_status_update, on_log, on_connection_loss):
        """
        Inicializa o serviço da balança.
        :param on_peso_update: Callback para ser chamado quando um novo peso é lido. Ex: fn(peso_float)
        :param on_status_update: Callback para log de status. Ex: fn(mensagem, cor)
        :param on_log: Callback para log geral de eventos. Ex: fn(mensagem)
        :param on_connection_loss: Callback para quando a conexão é perdida. Ex: fn()
        """
        self.ser = None
        self.monitorando = False
        self.monitoramento_pausado = False
        self.ultimo_peso_valido = None
        self.leitura_estavel = False
        self._stop_event = threading.Event()
        self._monitor_thread = None
        
        # Callbacks para comunicação com a camada de aplicação
        self.on_peso_update = on_peso_update
        self.on_status_update = on_status_update
        self.on_log = on_log
        self.on_connection_loss = on_connection_loss

    @staticmethod
    def listar_portas_disponiveis():
        """Retorna uma lista de portas seriais disponíveis."""
        portas = serial.tools.list_ports.comports()
        return [p.device for p in portas] if portas else ["Nenhuma"]

    def conectar(self, porta):
        """Tenta conectar na porta serial especificada e inicia o monitoramento."""
        if porta in ["Nenhuma", "..."]:
            return False, "Nenhuma porta selecionada."

        self._stop_event.clear()
        self.monitoramento_pausado = False

        try:
            self.ser = serial.Serial(porta, 1200, bytesize=7, parity='O', stopbits=1, timeout=0.5)

            # Testa por erro 30 antes de prosseguir
            erro30, resp = self._verificar_erro_30()
            if erro30:
                msg = (f"ATENÇÃO: Possível 'Erro 30' detectado na balança.\n\n"
                       "Este erro impede a comunicação e a leitura de pesos.\n\n"
                       "Para resolver:\n"
                       "1. Pressione o botão 📄 PRINT (ou ESC) no painel da balança.\n"
                       "2. Se o erro persistir, consulte o manual do aplicativo para mais instruções.\n\n"
                       f"Detalhes: {resp}")
                self.on_log(f"Conexão cancelada: erro 30 detectado na {porta}.")
                if self.ser: self.ser.close()
                self.ser = None
                return False, msg

            # Inicia o monitoramento em uma thread
            self.monitorando = True
            self._monitor_thread = threading.Thread(target=self._thread_monitoramento, daemon=True)
            self._monitor_thread.start()
            
            self.on_status_update(f"Conectado em {porta}", "#00FF00")
            self.on_log("Conectado! Pode iniciar as leituras.")
            return True, "Conectado com sucesso."

        except Exception as e:
            msg = (f"Não foi possível conectar à porta {porta}.\n\n"
                   "Possíveis causas:\n"
                   "1. A balança não está conectada ao computador.\n"
                   "2. Outro software está usando a mesma porta.\n"
                   "3. Ocorreu um 'Erro 30' interno na balança.\n\n"
                   "Solução para 'Erro 30':\n"
                   "1. Aperte o botão 📄 PRINT (ou ESC) na balança.\n"
                   "2. Consulte o manual do aplicativo para mais detalhes.\n\n"
                   f"Erro técnico: {e}")
            self.on_log(f"Falha na conexão com {porta}: {e}")
            return False, msg

    def desconectar(self):
        """Encerra a conexão serial e o monitoramento."""
        self.monitorando = False
        self._stop_event.set()
        if self._monitor_thread and self._monitor_thread.is_alive():
            self._monitor_thread.join(timeout=1.0)
        self._monitor_thread = None

        if self.ser and self.ser.is_open:
            self.ser.close()
        self.ser = None
        self.monitoramento_pausado = False
        self.on_status_update("Desconectado", "gray")
        self.on_log("Desconectado.")

    def _thread_monitoramento(self):
        """Loop que roda em background para ler o peso da balança."""
        while not self._stop_event.is_set() and self.monitorando and self.ser and self.ser.is_open:
            if self.monitoramento_pausado:
                time.sleep(0.5)
                continue
            try:
                self.ser.reset_input_buffer()
                self.ser.write(b'\x1bP\r\n') 
                linha = self.ser.readline().decode('ascii', errors='ignore')

                if not self.monitorando or self._stop_event.is_set():
                    break

                # A balança Sartorius geralmente indica estabilidade com um espaço antes do sinal.
                # Uma leitura instável pode ter '?' ou outro caractere.
                # Regex para capturar um valor que parece estável (começa com espaços/sinal).
                match = re.search(r"([-+ ]\s*(\d+\.\d+))", linha)
                if match:
                    peso_str = match.group(2) # Captura apenas o número
                    self.ultimo_peso_valido = float(peso_str)
                    self.leitura_estavel = '?' not in linha # Simples verificação de estabilidade
                    # Notifica a aplicação principal sobre o novo peso
                    if self.on_peso_update:
                        self.on_peso_update(self.ultimo_peso_valido, self.leitura_estavel)
                else:
                    self.leitura_estavel = False
                    # Notifica a UI sobre a instabilidade, mesmo sem um novo peso válido.
                    if self.on_peso_update:
                        self.on_peso_update(self.ultimo_peso_valido, self.leitura_estavel)

                time.sleep(0.2)
            except (serial.SerialException, OSError) as e:
                self.on_log(f"ERRO: A porta serial foi desconectada ou falhou: {e}")
                if self.on_connection_loss:
                    self.on_connection_loss()
                break # Encerra o loop
            except Exception as e:
                self.on_log(f"ERRO inesperado no monitoramento: {e}")
                time.sleep(1)

    def get_leitura_instantanea(self):
        """
        Tenta obter uma leitura estável da balança por até 3 segundos, com intervalos de 0.3s.
        Retorna (peso, True) se conseguir, ou (None, False) caso contrário.
        """
        if self._stop_event.is_set() or not self.ser or not self.ser.is_open:
            return None, False

        end_time = time.time() + 3.0
        while time.time() < end_time:
            try:
                self.ser.reset_input_buffer()
                self.ser.write(b'\x1bP\r\n')
                linha = self.ser.readline().decode('ascii', errors='ignore')

                match = re.search(r"([-+ ]\s*(\d+\.\d+))", linha)
                if match:
                    peso_str = match.group(2)
                    peso = float(peso_str)
                    estavel = '?' not in linha
                    if estavel:
                        return peso, True # Sucesso! Retorna a leitura estável.
            except (serial.SerialException, OSError, Exception) as e:
                self.on_log(f"ERRO ao obter leitura instantânea: {e}")
                return None, False # Falha crítica na comunicação
            time.sleep(0.3) # Aguarda antes da próxima tentativa
        
        return None, False # Falha: Tempo esgotado sem leitura estável

    def _verificar_erro_30(self, timeout=3.0):
        """Verifica se a balança responde com erro 30. Lógica interna."""
        if self._stop_event.is_set() or not self.ser or not self.ser.is_open:
            return (False, '')

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
                    texto = raw.decode('ascii', errors='ignore').strip()
                    if texto:
                        linhas.append(texto)
                        # Regex corrigida para evitar "FutureWarning: Possible nested set".
                        # A regex agora procura 'err' ou 'error', seguido por caracteres separadores (espaço, :, -), e então '30'.
                        if re.search(r'(?i)(err(or)?\s*[:\-\s]*\s*30)', texto) or (texto.isdigit() and int(texto) == 30):
                            return (True, f"Código 30 detectado: {texto}")
                        if re.search(r"[-+]?\s*\d+\.\d+", texto):
                            peso_lido = True
                            return (False, '')
            except:
                break
        
        if not peso_lido:
            resposta = "\\n".join(linhas) if linhas else "(sem resposta da balança)"
            return (True, f"Balança não forneceu leitura válida. Possível erro 30. Respostas: {resposta}")
        
        return (False, '')

    def enviar_comando_tara(self):
        """Envia o comando para tarar/zerar a balança."""
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(b'\x1bf4_\r\n') 
                self.on_log("\nComando de TARA enviado.\n")
                return True
            except Exception as e:
                self.on_log(f"Erro ao enviar comando de tara: {e}")
                return False
        return False
        
    def pausar_monitoramento(self):
        """Pausa a leitura contínua da balança."""
        self.monitoramento_pausado = True

    def retomar_monitoramento(self):
        """Retoma a leitura contínua da balança."""
        self.monitoramento_pausado = False

    def is_connected(self):
        return self.ser is not None and self.ser.is_open and not self._stop_event.is_set()

    def get_ultimo_peso(self):
        return self.ultimo_peso_valido