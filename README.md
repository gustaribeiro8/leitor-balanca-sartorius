# SISAQUI - Automação de Balança

Este projeto contém uma aplicação de desktop desenvolvida para interfacear com balanças analíticas (especificamente modelos Sartorius que utilizam comunicação serial), automatizando a coleta e o registro de dados de pesagem.

A aplicação permite a captura de pesos em um formato de planilha livre, com colunas designadas para amostras "Padrão (A)", "Cliente (B)" e "Genérico", salvando os resultados em um arquivo `.csv` com cálculo automático de estatísticas básicas.

## ✨ Funcionalidades

- **Conexão Serial:** Conecta-se a balanças através de portas COM virtuais, com listagem automática das portas disponíveis.
- **Leitura em Tempo Real:** Exibe o peso atual lido da balança em uma interface gráfica moderna.
- **Tratamento de Erro Específico:** Detecta e informa o usuário sobre o "Erro 30", um estado comum em balanças Sartorius que impede a comunicação.
- **Captura de Dados:** Permite capturar o peso com atalhos de teclado (`a`, `b`, `espaço`) ou cliques de botão, organizando-os em colunas.
- **Geração de Planilha (CSV):** Salva os dados em um arquivo `.csv` bem formatado, incluindo:
    - Data e hora da medição.
    - Colunas separadas para diferentes tipos de amostra.
    - Cálculo e registro de estatísticas (quantidade, média, mínimo, máximo) para cada tipo de amostra.
- **Interface Gráfica:** Interface de usuário construída com CustomTkinter, com modos claro e escuro.

## 📂 Estrutura do Projeto

O código-fonte foi refatorado para seguir uma arquitetura modular, separando as responsabilidades para facilitar a manutenção e a escalabilidade.

```
automatizacao_balancas/
│
├── codigo/
│   ├── app_principal.py    # Ponto de entrada e Controller da aplicação
│   ├── app_ui.py           # Módulo da Interface Gráfica (View)
│   ├── servico_balanca.py  # Módulo para comunicação com a balança
│   ├── servico_csv.py      # Módulo para manipulação de arquivos CSV
│   └── icone_sartorius.ico # Ícone da aplicação
│
├── dados coletados/        # Pasta onde os arquivos .csv são salvos
│
├── documentação/
│
├── requirements.txt        # Dependências do projeto
│
└── README.md               # Este arquivo
```

## 🚀 Começando

Siga estas instruções para configurar e executar o projeto em seu ambiente de desenvolvimento.

### Pré-requisitos

- Python 3.8 ou superior
- `pip` (gerenciador de pacotes do Python)

### Instalação

1. Clone este repositório ou baixe os arquivos para o seu computador.
2. Abra um terminal na pasta raiz do projeto (`automatizacao_balancas`).
3. Instale as dependências necessárias executando o seguinte comando:

   ```shell
   pip install -r requirements.txt
   ```

### Executando a Aplicação

Para iniciar a aplicação a partir do código-fonte, execute o script principal:

```shell
python codigo/app_principal.py
```

## 📦 Compilando para Executável (Build)

É possível gerar um arquivo executável (`.exe`) que encapsula toda a aplicação, permitindo que ela seja executada em outros computadores Windows sem a necessidade de instalar Python ou as dependências.

Para isso, o `PyInstaller` é utilizado.

1.  Certifique-se de que o PyInstaller está instalado:
    ```shell
    pip install pyinstaller
    ```
2.  Navegue até a pasta `codigo` pelo terminal:
    ```shell
    cd codigo
    ```
3.  Execute o seguinte comando para iniciar o processo de compilação:

    ```shell
    python -m PyInstaller --noconsole --onefile --clean --icon=icone_sartorius.ico --add-data "icone_sartorius.ico;." app_principal.py
    ```
    - `--noconsole`: Impede que uma janela de terminal seja aberta ao executar o `.exe`.
    - `--onefile`: Agrupa tudo em um único arquivo executável.
    - `--icon`: Define o ícone da aplicação.
    - `--add-data`: Garante que o arquivo de ícone seja incluído no build para ser exibido na janela da aplicação.

4.  Após a conclusão, o executável estará na pasta `dist` dentro da pasta `codigo`.

---
*Documentação gerada automaticamente.*
