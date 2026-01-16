import serial
import time
import re
import csv
import keyboard  # Biblioteca nova para detectar teclas
from datetime import datetime

# --- CONFIGURAÇÕES ---
PORTA = 'COM8'
BAUDRATE = 1200
BYTESIZE = serial.SEVENBITS
PARITY = serial.PARITY_ODD
STOPBITS = serial.STOPBITS_ONE
TIMEOUT = 1

ARQUIVO_SAIDA = 'dados_coletados.csv'

def conectar_sartorius():
    print("Conectando à balança...")
    try:
        ser = serial.Serial(
            port=PORTA,
            baudrate=BAUDRATE,
            bytesize=BYTESIZE,
            parity=PARITY,
            stopbits=STOPBITS,
            timeout=TIMEOUT,
            rtscts=False
        )
        print(f"SUCESSO: Conectado na {PORTA}")
        return ser
    except serial.SerialException as e:
        print(f"Erro de conexão: {e}")
        return None

def enviar_comando(ser, comando):
    comandos = {
        'TARAR': b'\x1bf4_',
        'PRINT': b'\x1bP',  # Comando para pedir o peso
    }
    if comando in comandos:
        ser.write(comandos[comando] + b'\r\n')
        # Pequena pausa para a balança processar
        time.sleep(0.2) 

def ler_peso_com_timeout(ser, tentativas=20):
    """
    Tenta ler o peso por X tentativas (aprox 2 segundos).
    Retorna o valor float ou None se falhar.
    """
    for _ in range(tentativas):
        if ser.in_waiting > 0:
            raw_data = ser.readline()
            try:
                texto = raw_data.decode('ascii', errors='ignore').strip()
                # Regex procura numero: sinal opcional, digitos, ponto, digitos
                match = re.search(r"[-+]?\s*\d+\.\d+", texto)
                if match:
                    peso_str = match.group().replace(" ", "")
                    return float(peso_str)
            except:
                pass
        time.sleep(0.1)
    return None

def inicializar_tabela():
    """Cria o arquivo CSV com o cabeçalho se ele não existir"""
    try:
        with open(ARQUIVO_SAIDA, 'x', newline='') as arquivo:
            escritor = csv.writer(arquivo, delimiter=';')
            escritor.writerow(['ID', 'Data', 'Hora', 'Peso (g)'])
    except FileExistsError:
        pass # Arquivo já existe, tudo bem

def salvar_dados(id_amostra, peso):
    """Escreve uma linha nova no arquivo"""
    agora = datetime.now()
    data = agora.strftime("%d/%m/%Y")
    hora = agora.strftime("%H:%M:%S")
    
    with open(ARQUIVO_SAIDA, 'a', newline='') as arquivo:
        escritor = csv.writer(arquivo, delimiter=';')
        escritor.writerow([id_amostra, data, hora, str(peso).replace('.', ',')])
    
    print(f"✅ SALVO: Amostra {id_amostra} | {peso} g")

# --- PROGRAMA PRINCIPAL ---
if __name__ == "__main__":
    balanca = conectar_sartorius()
    
    if balanca:
        inicializar_tabela()
        contador_amostras = 1
        
        print("\n" + "="*40)
        print(" SISTEMA DE AQUISIÇÃO DE DADOS (TIPO SOMA)")
        print(f" Arquivo de saída: {ARQUIVO_SAIDA}")
        print(" COMANDOS:")
        print("  [ESPAÇO] -> Capturar Peso")
        print("  [ESC]    -> Sair")
        print("="*40 + "\n")

        try:
            while True:
                # Se apertar ESC, sai do programa
                if keyboard.is_pressed('esc'):
                    print("Encerrando...")
                    break
                
                # Se apertar ESPAÇO, faz a leitura
                if keyboard.is_pressed('space'):
                    print("🔄 Solicitando leitura à balança...", end='\r')
                    
                    # 1. Limpa o buffer antigo para não pegar leitura velha
                    balanca.reset_input_buffer()
                    
                    # 2. Envia comando para a balança mandar o peso (Software Trigger)
                    enviar_comando(balanca, 'PRINT')
                    
                    # 3. Lê a resposta
                    peso = ler_peso_com_timeout(balanca)
                    
                    if peso is not None:
                        salvar_dados(contador_amostras, peso)
                        contador_amostras += 1
                    else:
                        print("❌ Erro: Balança não respondeu ou leitura instável.")
                    
                    # Espera a tecla ser solta para não registrar 10 vezes seguidas
                    while keyboard.is_pressed('space'):
                        time.sleep(0.1)

                time.sleep(0.05) # Evita uso excessivo da CPU

        except KeyboardInterrupt:
            print("Parando...")
        finally:
            balanca.close()