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

    def salvar_lote_medidas(self, nome_ensaio, pesos_A, pesos_B):
        """Salva um lote de medidas, adicionando uma linha em branco no final."""
        pesos_por_tipo = {
            'Padrao (A)': pesos_A,
            'Cliente (B)': pesos_B
        }
        return self._salvar_medidas(nome_ensaio, pesos_por_tipo, adicionar_linha_branca=True)

    def salvar_medida(self, nome_ensaio, tipo_amostra, peso_valido):
        """
        Salva uma nova medida no arquivo CSV.
        Esta função lê o arquivo existente, adiciona a nova medida,
        recalcula as estatísticas e reescreve o arquivo de forma atômica.
        """
        if peso_valido is None:
            return None, None, "Peso inválido fornecido."
        return self._salvar_medidas(nome_ensaio, {tipo_amostra: [peso_valido]})

    def _salvar_medidas(self, nome_ensaio, pesos_por_tipo, adicionar_linha_branca=False):
        """
        Lógica central para salvar uma ou mais medidas no arquivo CSV.
        :param pesos_por_tipo: Um dicionário {'Tipo Amostra': [lista de pesos]}.
        :param adicionar_linha_branca: Se True, adiciona uma linha vazia após as medições.
        """
        if not any(pesos_por_tipo.values()):
            return None, None, "Nenhum peso válido fornecido para salvar."

        arquivo = self._get_caminho_arquivo(nome_ensaio)
        agora = datetime.now()

        dados_por_tipo = {'Padrao (A)': [], 'Cliente (B)': [], 'Generico': []}
        cabecalho1, cabecalho2 = [], []
        
        try:
            # 1. Leitura do arquivo existente (se houver)
            if os.path.exists(arquivo):
                with open(arquivo, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f, delimiter=';')
                    linhas = list(reader)
                    if len(linhas) >= 2:
                        cabecalho1 = linhas[0]
                        cabecalho2 = linhas[1]
                        lendo_dados = True
                        for linha in linhas[2:]:
                            # Ignora linhas de estatísticas, separadores ou linhas em branco
                            if not linha or not any(linha):
                                continue
                            if "Estatisticas" in linha[0]:
                                lendo_dados = False # Para de ler dados ao encontrar a seção de estatísticas
                            if lendo_dados:
                                # Recria a estrutura de dados a partir das colunas
                                if len(linha) > 4 and linha[3]: dados_por_tipo['Padrao (A)'].append([linha[3], linha[4]])
                                if len(linha) > 7 and linha[6]: dados_por_tipo['Cliente (B)'].append([linha[6], linha[7]])
                                if len(linha) > 10 and linha[9]: dados_por_tipo['Generico'].append([linha[9], linha[10]])

            data_str = agora.strftime("%d/%m/%Y")
            for tipo_amostra, lista_pesos in pesos_por_tipo.items():
                for peso in lista_pesos:
                    nova_medida = [str(peso).replace('.', ','), agora.strftime("%H:%M:%S")]
                    dados_por_tipo[tipo_amostra].append(nova_medida)

            # Contadores atualizados
            contagem_A = len(dados_por_tipo.get('Padrao (A)', []))
            contagem_B = len(dados_por_tipo.get('Cliente (B)', []))

            # 3. Escrita Atômica (em um arquivo temporário)
            arquivo_tmp = arquivo + ".tmp"
            with open(arquivo_tmp, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(cabecalho1 or ['Data', 'Hora', '', 'Padrao (A)', '', '', 'Cliente (B)', '', '', 'Generico', ''])
                writer.writerow(cabecalho2 or ['', '', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora', '', 'Peso (g)', 'Hora'])

                max_medidas = max(len(v) for v in dados_por_tipo.values()) if any(dados_por_tipo.values()) else 0
                
                for i in range(max_medidas):
                    # Adiciona data e hora apenas na primeira linha para clareza
                    data_col = data_str if i == 0 and not cabecalho1 else ''
                    hora_col = agora.strftime("%H:%M:%S") if i == 0 else ''
                    linha_csv = [data_col, hora_col, '']
                    
                    # Preenche as colunas com os dados ou com strings vazias
                    linha_csv.extend(dados_por_tipo['Padrao (A)'][i] if i < len(dados_por_tipo['Padrao (A)']) else ['', ''])
                    linha_csv.append('')
                    linha_csv.extend(dados_por_tipo['Cliente (B)'][i] if i < len(dados_por_tipo['Cliente (B)']) else ['', ''])
                    linha_csv.append('')
                    linha_csv.extend(dados_por_tipo['Generico'][i] if i < len(dados_por_tipo['Generico']) else ['', ''])
                    writer.writerow(linha_csv)

                    # Insere uma linha em branco após cada lote completo de 4 medidas,
                    # para facilitar a visualização ao abrir no Excel/Calc.
                    if (i + 1) % 4 == 0 and i < max_medidas - 1:
                        writer.writerow([])

                if adicionar_linha_branca:
                    writer.writerow([]) # Adiciona linha em branco para separar lotes

                # 4. Cálculo e escrita das estatísticas
                # Adiciona uma linha em branco para separar os dados das estatísticas (ou para separar lotes)
                writer.writerow([])

                writer.writerow(['Estatisticas:'])
                writer.writerow(['Tipo', 'Quantidade', 'Media (g)', 'Minimo (g)', 'Maximo (g)'])
                
                for tipo in ['Padrao (A)', 'Cliente (B)', 'Generico']:
                    if dados_por_tipo[tipo]:
                        pesos_float = [float(m[0].replace(',', '.')) for m in dados_por_tipo[tipo] if m[0]]
                        stats = [
                            tipo, 
                            len(pesos_float), 
                            f"{sum(pesos_float) / len(pesos_float):.6f}".replace('.', ','), 
                            f"{min(pesos_float):.6f}".replace('.', ','), 
                            f"{max(pesos_float):.6f}".replace('.', ',')
                        ]
                        writer.writerow(stats)
            
            # 5. Substituição do arquivo original pelo temporário
            os.replace(arquivo_tmp, arquivo)

            total_medidas_salvas = sum(len(v) for v in pesos_por_tipo.values())
            self.on_log(f"Lote de {total_medidas_salvas} medida(s) salvo no arquivo {os.path.basename(arquivo)}")
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