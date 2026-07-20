import sys
from app_ui import AppUI
import threading
from servico_balanca import ServicoBalanca
from servico_csv import ServicoCsv

class AppPrincipal:
    def __init__(self):
        # --- ESTADO DA APLICAÇÃO ---
        self.conectado = False
        self.leitura_estavel = False
        self.contadores_totais = {'A': 0, 'B': 0}
        self._capturando = False # Novo estado para evitar capturas simultâneas
        self._encerrando = False

        # Estado para controle de lotes mistos (ex: ABBA)
        self.lote_em_andamento_A = []
        self.lote_em_andamento_B = []

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
            # Pausar temporariamente caso já esteja em um estado inconsistente
            self.servico_balanca.pausar_monitoramento()
            try:
                sucesso, mensagem = self.servico_balanca.conectar(porta)
                if sucesso:
                    self.conectado = True
                    self.ui.set_estado_conectado(True)
                else:
                    self.ui.show_error("Erro de Conexão", mensagem)
                    self.conectado = False
                    self.ui.set_estado_conectado(False)
            finally:
                # Retoma o monitoramento se a conexão foi bem-sucedida
                if self.conectado:
                    self.servico_balanca.retomar_monitoramento()

    def tarar_balanca(self):
        if self.conectado:
            self.servico_balanca.enviar_comando_tara()

    def capturar_coluna(self, coluna_letra):
        """Inicia o processo de captura de uma nova medida."""
        if not self.conectado:
            self.log("Conecte à balança para capturar.")
            return
        
        if self._capturando:
            self.log("Aguarde, captura anterior em andamento...")
            return

        # Inicia o processo de captura em uma nova thread para não travar a UI
        threading.Thread(target=self._thread_captura, args=(coluna_letra,), daemon=True).start()

    def _thread_captura(self, coluna_letra):
        """
        Executa em uma thread separada para obter a leitura da balança
        sem congelar a interface do usuário.
        """
        self._capturando = True
        self._safe_schedule_ui(self.ui.set_estado_capturando, True)

        try:
            # Solicita uma leitura instantânea no momento da captura
            peso_capturado, estavel = self.servico_balanca.get_leitura_instantanea()

            if not estavel or peso_capturado is None:
                # Pausa o monitoramento antes de mostrar o alerta modal
                self.servico_balanca.pausar_monitoramento()
                try:
                    self.log("⚠️ Captura cancelada: leitura instável.")
                    # A chamada para a UI (show_warning) deve ser agendada na thread principal
                    self._safe_schedule_ui(
                        self.ui.show_warning, 
                        "Aviso", 
                        "A leitura da balança não está estável. Aguarde a estabilização para capturar o peso."
                    )
                finally:
                    self.servico_balanca.retomar_monitoramento()
                return # Finaliza a thread de captura

            # A lógica de negócio pode continuar aqui, mas as atualizações da UI
            # devem ser agendadas para a thread principal.
            self._safe_schedule_ui(self._processar_captura_bem_sucedida, coluna_letra, peso_capturado)

        finally:
            self._capturando = False
            self._safe_schedule_ui(self.ui.set_estado_capturando, False)

    def _processar_captura_bem_sucedida(self, coluna_letra, peso_capturado):
        """Processa a captura após uma leitura bem-sucedida (executa na thread da UI)."""
        if coluna_letra == 'A':
            self.lote_em_andamento_A.append(peso_capturado)
            self.log(f"Adicionado à Coluna A: {peso_capturado:.6f}g (Total A: {len(self.lote_em_andamento_A)})")
        elif coluna_letra == 'B':
            self.lote_em_andamento_B.append(peso_capturado)
            self.log(f"Adicionado à Coluna B: {peso_capturado:.6f}g (Total B: {len(self.lote_em_andamento_B)})")
        else: # Coluna Genérica (G) - salva imediatamente
            self._salvar_medida_unica(peso_capturado)
            return
        
        # Atualiza a UI para mostrar o progresso do lote
        self.ui.atualizar_contadores_lote(len(self.lote_em_andamento_A), len(self.lote_em_andamento_B))
        self.ui.flash_button(coluna_letra)

        # Nova lógica para verificar se um lote está completo
        count_A = len(self.lote_em_andamento_A)
        count_B = len(self.lote_em_andamento_B)

        # Um lote está completo se:
        # 1. Apenas um tipo foi coletado e atingiu 4 medições.
        # 2. Ambos os tipos foram coletados e ambos atingiram 4 medições.
        lote_completo = (count_A == 4 and count_B == 0) or \
                        (count_B == 4 and count_A == 0) or \
                        (count_A == 4 and count_B == 4)

        if lote_completo:
            self._confirmar_e_salvar_lote() # Esta função já pausa/retoma o monitoramento

    def _confirmar_e_salvar_lote(self):
        """Mostra um popup de confirmação e salva o lote se o usuário concordar."""
        leituras_A_str = "\n".join([f"  - {peso:.6f} g" for peso in self.lote_em_andamento_A])
        leituras_B_str = "\n".join([f"  - {peso:.6f} g" for peso in self.lote_em_andamento_B])
        
        mensagem = "Lote de medições completo.\n\n"
        if leituras_A_str:
            mensagem += f"Leituras para A:\n{leituras_A_str}\n"
        if leituras_B_str:
            mensagem += f"Leituras para B:\n{leituras_B_str}\n"
        mensagem += "\nDeseja salvar este lote?"
        
        # Pausa o monitoramento para evitar que a UI seja inundada com eventos
        # enquanto a caixa de diálogo modal está aberta.
        self.servico_balanca.pausar_monitoramento()
        try:
            # Usando askyesno: "Sim" para salvar, "Não" para refazer
            if self.ui.show_confirmation("Confirmar Lote", mensagem):
                self._salvar_lote_atual()
            else:
                self.log(f"Lote descartado pelo usuário. Por favor, refaça as 4 leituras.")
                self.ui.show_info("Lote Descartado", "O lote foi descartado. Você pode iniciar a coleta das 4 amostras novamente.")
        finally:
            # Garante que o monitoramento seja retomado e o estado resetado,
            # independentemente da escolha do usuário.
            self.servico_balanca.retomar_monitoramento()
            self.lote_em_andamento_A = []
            self.lote_em_andamento_B = []
            self.ui.atualizar_contadores(self.contadores_totais['A'], self.contadores_totais['B']) # Volta a exibir totais

    def _salvar_lote_atual(self):
        """Envia o lote para o serviço de CSV e atualiza os contadores."""
        nome_ensaio = self.ui.get_nome_ensaio()

        # Salva o lote de medidas usando o serviço de CSV
        count_A, count_B, erro = self.servico_csv.salvar_lote_medidas(
            nome_ensaio, 
            self.lote_em_andamento_A, 
            self.lote_em_andamento_B
        )

        if erro:
            self.ui.show_error("Erro ao Salvar", erro)
            # Não há necessidade de pausar aqui, pois o erro já ocorreu e a UI não está bloqueada
        else:
            self.contadores_totais['A'] = count_A
            self.contadores_totais['B'] = count_B
            self.ui.atualizar_contadores(count_A, count_B)

    def _salvar_medida_unica(self, peso):
        """Salva uma única medida na coluna 'Genérico'."""
        nome_ensaio = self.ui.get_nome_ensaio()
        # Pausa o monitoramento antes de salvar, caso ocorra um erro de permissão
        # que exiba uma caixa de diálogo.
        self.servico_balanca.pausar_monitoramento()
        try:
            _, _, erro = self.servico_csv.salvar_medida(nome_ensaio, 'Generico', peso)
            if erro:
                self.ui.show_error("Erro ao Salvar", erro)
        finally:
            self.servico_balanca.retomar_monitoramento()

    def abrir_arquivo(self):
        nome_ensaio = self.ui.get_nome_ensaio()
        # Pausa para o caso de o arquivo não existir e um warning ser exibido
        self.servico_balanca.pausar_monitoramento()
        try:
            sucesso, mensagem = self.servico_csv.abrir_no_explorer(nome_ensaio)
            if not sucesso:
                self.ui.show_warning("Aviso", mensagem)
            self.log(mensagem)
        finally:
            self.servico_balanca.retomar_monitoramento()
        
    def atualizar_lista_portas(self):
        portas = self.servico_balanca.listar_portas_disponiveis()
        self.ui.atualizar_lista_portas(portas)
        self.log("Lista de portas atualizada.")

    def on_closing(self):
        """Chamado quando a janela é fechada."""
        if self._encerrando:
            return

        self._encerrando = True
        self.log("Fechando aplicação...")
        try:
            self.servico_balanca.desconectar()
        finally:
            if self.ui and self.ui.winfo_exists():
                self.ui.destroy()
            sys.exit(0)

    # --- MÉTODOS CHAMADOS PELOS SERVIÇOS (Callbacks) ---
    def on_peso_update(self, peso, estavel):
        """Callback: Chamado pelo ServicoBalanca quando há novo peso."""
        if self._encerrando:
            return

        self.leitura_estavel = estavel
        self._safe_schedule_ui(self._handle_peso_update_ui, peso, estavel)

    def _handle_peso_update_ui(self, peso, estavel):
        """Garante que a atualização da UI do peso ocorra na thread principal."""
        self.ui.atualizar_peso_display(peso, estavel)
        self.ui.atualizar_status("Leitura estável" if estavel else "Balança instável, aguarde", "white" if estavel else "orange")
        
    def on_status_update(self, texto, cor):
        """Callback: Chamado pelo ServicoBalanca para atualizar o status."""
        if self._encerrando:
            return
        self._safe_ui_call(lambda: self.ui.atualizar_status(texto, cor))

    def on_connection_loss(self):
        """Callback: Chamado pelo ServicoBalanca quando a conexão cai."""
        if self._encerrando:
            return

        self.conectado = False
        self._safe_schedule_ui(self._handle_connection_loss_ui)

    def _handle_connection_loss_ui(self):
        """Garante que a atualização da UI ocorra na thread principal."""
        self.ui.set_estado_conectado(False)
        self.ui.atualizar_status("Conexão perdida. Reconecte.", "orange")
        self.log("‼️ Conexão com a balança foi perdida.")
        # Não é necessário pausar aqui, pois a thread de monitoramento já parou.
        # Apenas exibimos o aviso.
        self.ui.show_warning(
            "Conexão Perdida", 
            "A comunicação com a balança foi interrompida. Verifique o cabo e reconecte."
        )

    def _safe_ui_call(self, func):
        """Executa uma chamada de UI apenas se a janela ainda existir."""
        if self._encerrando:
            return
        if self.ui and self.ui.winfo_exists():
            func()

    def _safe_schedule_ui(self, func, *args):
        """Agenda uma atualização da UI na thread principal, se a janela ainda existir."""
        if self._encerrando:
            return
        if self.ui and self.ui.winfo_exists():
            self.ui.after(0, func, *args)

    def log(self, mensagem):
        """Ponto central para logging. Chamado por todos os componentes."""
        self._safe_schedule_ui(self.ui.log, mensagem)

if __name__ == "__main__":
    # Define o modo de aparência antes de instanciar qualquer widget
    import customtkinter as ctk
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    
    app = AppPrincipal()
    app.run()
