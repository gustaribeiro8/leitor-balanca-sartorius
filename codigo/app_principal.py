import sys
from app_ui import AppUI
from servico_balanca import ServicoBalanca
from servico_csv import ServicoCsv

class AppPrincipal:
    def __init__(self):
        # --- ESTADO DA APLICAÇÃO ---
        self.conectado = False
        self.ultimo_peso_valido = None
        self.contadores = {'A': 0, 'B': 0}

        # --- INICIALIZAÇÃO DOS MÓDULOS ---
        # O Controller passa suas próprias funções (callbacks) para os serviços e a UI.
        # self.log é uma função do controller que a UI e os serviços podem chamar.
        self.ui = AppUI(self)
        self.servico_csv = ServicoCsv(self.log)
        self.servico_balanca = ServicoBalanca(
            on_peso_update=self.on_peso_update,
            on_status_update=self.on_status_update,
            on_log=self.log,
            on_connection_loss=self.on_connection_loss
        )
        
        # Inicia a aplicação
        self.ui.show_info("Atenção", "Certifique-se de que a balança está ligada, estável e pronta para a conexão.")
        self.atualizar_lista_portas()
        
    def run(self):
        """Inicia o loop principal da interface gráfica."""
        self.ui.mainloop()

    # --- MÉTODOS CHAMADOS PELA UI (Ações do Usuário) ---
    def alternar_conexao(self):
        if self.conectado:
            self.servico_balanca.desconectar()
            self.conectado = False
            self.ui.set_estado_conectado(False)
        else:
            porta = self.ui.get_porta_selecionada()
            sucesso, mensagem = self.servico_balanca.conectar(porta)
            if sucesso:
                self.conectado = True
                self.ui.set_estado_conectado(True)
            else:
                self.ui.show_error("Erro de Conexão", mensagem)
                self.conectado = False
                self.ui.set_estado_conectado(False)

    def tarar_balanca(self):
        if self.conectado:
            self.servico_balanca.enviar_comando_tara()

    def capturar_coluna(self, coluna_letra):
        if not self.conectado:
            self.log("Conecte à balança para capturar.")
            return

        tipo_map = {'A': 'Padrao (A)', 'B': 'Cliente (B)', 'G': 'Generico'}
        tipo_amostra = tipo_map.get(coluna_letra, 'Generico')
        
        self.log(f"Capturando para coluna {coluna_letra} -> {tipo_amostra}...")
        
        nome_ensaio = self.ui.get_nome_ensaio()
        
        # Salva a medida usando o serviço de CSV
        count_A, count_B, erro = self.servico_csv.salvar_medida(
            nome_ensaio, tipo_amostra, self.ultimo_peso_valido
        )
        
        if erro:
            self.ui.show_error("Erro ao Salvar", erro)
        else:
            # Atualiza contadores e UI
            self.contadores['A'] = count_A
            self.contadores['B'] = count_B
            self.ui.atualizar_contadores(count_A, count_B)
            if coluna_letra in ["A", "B"]:
                self.ui.flash_button(coluna_letra)

    def abrir_arquivo(self):
        nome_ensaio = self.ui.get_nome_ensaio()
        sucesso, mensagem = self.servico_csv.abrir_no_explorer(nome_ensaio)
        if not sucesso:
            self.ui.show_warning("Aviso", mensagem)
        self.log(mensagem)
        
    def atualizar_lista_portas(self):
        portas = self.servico_balanca.listar_portas_disponiveis()
        self.ui.atualizar_lista_portas(portas)
        self.log("Lista de portas atualizada.")

    def on_closing(self):
        """Chamado quando a janela é fechada."""
        self.log("Fechando aplicação...")
        self.servico_balanca.desconectar()
        self.ui.destroy()
        sys.exit()

    # --- MÉTODOS CHAMADOS PELOS SERVIÇOS (Callbacks) ---
    def on_peso_update(self, peso):
        """Callback: Chamado pelo ServicoBalanca quando há novo peso."""
        self.ultimo_peso_valido = peso
        self.ui.atualizar_peso_display(peso)

    def on_status_update(self, texto, cor):
        """Callback: Chamado pelo ServicoBalanca para atualizar o status."""
        self.ui.atualizar_status(texto, cor)

    def on_connection_loss(self):
        """Callback: Chamado pelo ServicoBalanca quando a conexão cai."""
        self.conectado = False
        self.ui.after(0, self._handle_connection_loss_ui)

    def _handle_connection_loss_ui(self):
        """Garante que a atualização da UI ocorra na thread principal."""
        self.ui.set_estado_conectado(False)
        self.ui.atualizar_status("Conexão perdida. Reconecte.", "orange")
        self.log("‼️ Conexão com a balança foi perdida.")
        self.ui.show_warning("Conexão Perdida", "A comunicação com a balança foi interrompida. Verifique o cabo e reconecte.")

    def log(self, mensagem):
        """Ponto central para logging. Chamado por todos os componentes."""
        # O self.ui.after garante que a chamada para a UI seja da thread principal
        self.ui.after(0, self.ui.log, mensagem)

if __name__ == "__main__":
    # Define o modo de aparência antes de instanciar qualquer widget
    import customtkinter as ctk
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    
    app = AppPrincipal()
    app.run()
