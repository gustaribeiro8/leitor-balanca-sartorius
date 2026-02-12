import csv
import os
from datetime import datetime

class ServicoCsv:
    def __init__(self, on_log):
        """
        Inicializa o serviço de CSV.
        :param on_log: Callback para log de eventos. Ex: fn(mensagem)
        """
        self.on_log = on_log

    def _get_caminho_arquivo(self, nome_ensaio):
        """Constrói o caminho completo do arquivo CSV."""
        pasta = "dados coletados"
        if not os.path.exists(pasta):
            os.makedirs(pasta)
        
        nome_base = nome_ensaio.strip() or f"ensaio_{datetime.now().strftime('%Y-%m-%d')}"
        if not nome_base.lower().endswith(".csv"):
            nome_base += ".csv"
            
        return os.path.join(pasta, nome_base)

    def abrir_no_explorer(self, nome_ensaio):
        """Abre o arquivo CSV no programa padrão (como Excel)."""
        arquivo = self._get_caminho_arquivo(nome_ensaio)
        if os.path.exists(arquivo):
            try:
                os.startfile(arquivo)
                return True, f"Abrindo {os.path.basename(arquivo)}..."
            except Exception as e:
                return False, f"Não foi possível abrir o arquivo: {e}"
        else:
            return False, "Arquivo ainda não existe. Salve uma medição primeiro."

    def salvar_medida(self, nome_ensaio, tipo_amostra, peso_valido):
        """
        Salva uma nova medida no arquivo CSV.
        Esta função lê o arquivo existente, adiciona a nova medida,
        recalcula as estatísticas e reescreve o arquivo de forma atômica.
        """
        if peso_valido is None:
            self.on_log("⚠️ Aguarde uma leitura estável da balança...")
            return None, None, "Leitura de peso inválida."

        arquivo = self._get_caminho_arquivo(nome_ensaio)
        agora = datetime.now()

        # Estrutura para armazenar dados lidos e novos
        dados_por_tipo = {'Padrao (A)': [], 'Cliente (B)': [], 'Generico': []}
        datas = []
        
        try:
            # 1. Leitura do arquivo existente (se houver)
            if os.path.exists(arquivo):
                with open(arquivo, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f, delimiter=';')
                    linhas = list(reader)
                    if len(linhas) > 2:
                        for linha in linhas[2:]:
                            # Ignora linhas de estatísticas ou separadores
                            if not linha or "Estatisticas" in linha[0]: break
                            if len(linha) >= 11:
                                if linha[0] and linha[0] != 'Data' and linha[0] not in datas: datas.append(linha[0])
                                if linha[3]: dados_por_tipo['Padrao (A)'].append([linha[3], linha[4]])
                                if linha[6]: dados_por_tipo['Cliente (B)'].append([linha[6], linha[7]])
                                if linha[9]: dados_por_tipo['Generico'].append([linha[9], linha[10]])

            # 2. Adição da nova medida à estrutura em memória
            data_str = agora.strftime("%d/%m/%Y")
            if data_str not in datas: datas.append(data_str)
            
            nova_medida = [str(peso_valido).replace('.', ','), agora.strftime("%H:%M:%S")]
            dados_por_tipo[tipo_amostra].append(nova_medida)

            # Contadores atualizados
            contagem_A = len(dados_por_tipo.get('Padrao (A)', []))
            contagem_B = len(dados_por_tipo.get('Cliente (B)', []))

            # 3. Escrita Atômica (em um arquivo temporário)
            arquivo_tmp = arquivo + ".tmp"
            with open(arquivo_tmp, 'w', newline='', encoding='utf--8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['Data', 'Hora', '', 'Padrao (A)', '', '', 'Cliente (B)', '', '', 'Generico', ''])
                writer.writerow(['', '', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora'])
                
                max_medidas = max(len(v) for v in dados_por_tipo.values()) if any(dados_por_tipo.values()) else 0
                
                for i in range(max_medidas):
                    if i > 0 and i % 4 == 0: writer.writerow([]) # Linha de separação
                    
                    data_col = datas[-1] if datas and i == 0 else ''
                    hora_col = agora.strftime("%H:%M:%S") if i == 0 else ''
                    linha_csv = [data_col, hora_col, '']
                    
                    linha_csv.extend(dados_por_tipo['Padrao (A)'][i] if i < len(dados_por_tipo['Padrao (A)']) else ['', ''])
                    linha_csv.append('')
                    linha_csv.extend(dados_por_tipo['Cliente (B)'][i] if i < len(dados_por_tipo['Cliente (B)']) else ['', ''])
                    linha_csv.append('')
                    linha_csv.extend(dados_por_tipo['Generico'][i] if i < len(dados_por_tipo['Generico']) else ['', ''])
                    writer.writerow(linha_csv)
                
                # 4. Cálculo e escrita das estatísticas
                writer.writerow([])
                writer.writerow(['Estatisticas:'])
                writer.writerow(['Tipo', 'Quantidade', 'Media (g)', 'Minimo (g)', 'Maximo (g)'])
                
                for tipo in ['Padrao (A)', 'Cliente (B)', 'Generico']:
                    if dados_por_tipo[tipo]:
                        pesos_float = [float(m[0].replace(',', '.')) for m in dados_por_tipo[tipo]]
                        stats = [
                            tipo, 
                            len(pesos_float), 
                            f"{sum(pesos_float) / len(pesos_float):.4f}".replace('.', ','), 
                            f"{min(pesos_float):.4f}".replace('.', ','), 
                            f"{max(pesos_float):.4f}".replace('.', ',')
                        ]
                        writer.writerow(stats)
            
            # 5. Substituição do arquivo original pelo temporário
            os.replace(arquivo_tmp, arquivo)

            self.on_log(f"Salvo: {peso_valido:.4f}g como [{tipo_amostra}]")
            return contagem_A, contagem_B, None # Sucesso

        except PermissionError:
            msg = f"Não foi possível salvar o arquivo '{os.path.basename(arquivo)}'.\n\nVerifique se ele não está aberto em outro programa (como o Excel) e tente novamente."
            if os.path.exists(arquivo_tmp): os.remove(arquivo_tmp)
            return None, None, msg
        except Exception as e:
            msg = f"Ocorreu um erro inesperado ao salvar o arquivo:\n\n{e}"
            self.on_log(f"ERRO ao salvar medida: {e}")
            if os.path.exists(arquivo_tmp): os.remove(arquivo_tmp)
            return None, None, msg