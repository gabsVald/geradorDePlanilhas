"""
================================================================================
ui/interface.py — Interface Gráfica Principal (CustomTkinter)
================================================================================
Define a janela principal do aplicativo com:
  - Cabeçalho vermelho com título
  - Botão de processamento (Colar e Gerar Planilhas)
  - Label de status
  - Barra de progresso
  - Modo de desenvolvedor oculto (ativado digitando "dev")

Threading:
  O processamento é executado em uma thread daemon separada para evitar
  que a interface trave durante operações de rede ou geração de arquivos.
  Atualizações da UI são feitas via self.after(0, callback) — thread-safe.

Modo de Teste:
  Ativado/desativado digitando "d" → "e" → "v" na janela.
  No modo teste: pasta de destino é Desktop/TESTES_GERADOR, sem cache de rede.
================================================================================
"""

import os
import threading

import customtkinter as ctk       # Framework de UI moderno baseado em tkinter
from tkinter import messagebox    # Diálogos nativos do sistema operacional

from utils.config import REGRAS               # Configurações globais (versão, etc.)
from utils.updater import verificar_atualizacao  # Verificação de atualização ao iniciar
from core.processador import processar_clipboard  # Lógica principal de processamento


# ===== CONSTANTES DE COR =====
COR_PRINCIPAL = "#d32732"  # Vermelho Ingecon (cabeçalho e botão normal)
COR_HOVER     = "#a81f28"  # Vermelho mais escuro para hover
COR_TESTE     = "#e67e22"  # Laranja — indica que o modo de teste está ativo

ctk.set_appearance_mode("Light")  # Tema claro para consistência com o ambiente corporativo


class AppIngecon(ctk.CTk):
    """
    Janela principal do aplicativo Ingecon - Gerador de Planilhas.

    Herda de ctk.CTk (janela raiz do CustomTkinter).
    Inicializa a interface, verifica atualizações e aguarda interação do usuário.
    """

    def __init__(self):
        super().__init__()
        verificar_atualizacao()  # Verifica versão no servidor antes de montar a UI
        self.versao = REGRAS.get("versao_atual", "Desconhecida")
        self.setup_window()
        self.setup_variables()
        self.create_widgets()
        # Captura todas as teclas digitadas para detectar o código secreto "dev"
        self.bind("<Key>", self.verificar_codigo_secreto)

    # ===== CONFIGURAÇÃO INICIAL =====

    def setup_window(self):
        """Configura título, tamanho e cor de fundo da janela principal."""
        self.title(f"Ingecon - Gerador de Planilhas V{self.versao}")
        self.geometry("480x420")
        self.configure(fg_color="#f5f5f5")  # Fundo cinza claro
        self.grid_columnconfigure(0, weight=1)  # Coluna única expande com a janela

    def setup_variables(self):
        """Inicializa variáveis de estado do aplicativo."""
        self.modo_teste_ativo = False  # False = produção, True = modo desenvolvedor
        self.buffer_teclas    = ""    # Acumula últimas teclas para detectar "dev"
        self.SECRET_CODE      = "dev" # Sequência de ativação do modo de teste

    def create_widgets(self):
        """Cria e posiciona todos os widgets da interface."""

        # --- Cabeçalho vermelho ---
        self.header_frame = ctk.CTkFrame(self, fg_color=COR_PRINCIPAL, height=70, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))
        self.header_frame.grid_columnconfigure(0, weight=1)

        self.label_titulo = ctk.CTkLabel(
            self.header_frame,
            text="GERADOR DE PLANILHAS",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="white"
        )
        self.label_titulo.grid(row=0, column=0, pady=20)

        # --- Botão principal ---
        self.btn_processar = ctk.CTkButton(
            self,
            text="Colar - Gerar Planilhas",
            command=self.iniciar_processamento,
            fg_color=COR_PRINCIPAL,
            hover_color=COR_HOVER,
            height=50,
            corner_radius=8,
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.btn_processar.grid(row=1, column=0, padx=40, pady=40)

        # --- Label de status (feedback textual) ---
        self.status_label = ctk.CTkLabel(self, text="Aguardando comando...", text_color="gray")
        self.status_label.grid(row=2, column=0, pady=5)

        # --- Barra de progresso (oculta até o processamento começar) ---
        self.progress = ctk.CTkProgressBar(
            self,
            orientation="horizontal",
            progress_color=COR_PRINCIPAL,
            width=300
        )
        self.progress.set(0)
        # Nota: a barra é adicionada ao grid apenas quando o processamento inicia

    # ===== MODO DE TESTE (CÓDIGO SECRETO) =====

    def verificar_codigo_secreto(self, event):
        """
        Detecta a digitação do código secreto "dev" para ativar/desativar o modo de teste.

        Mantém um buffer das últimas N teclas (N = len(SECRET_CODE)) e compara
        com o código. Funciona em qualquer campo da janela sem interferir no uso normal.

        Parâmetros:
            event: Evento de teclado do tkinter.
        """
        self.buffer_teclas += event.char.lower()
        # Mantém o buffer no tamanho exato do código secreto (janela deslizante)
        if len(self.buffer_teclas) > len(self.SECRET_CODE):
            self.buffer_teclas = self.buffer_teclas[-len(self.SECRET_CODE):]

        if self.buffer_teclas == self.SECRET_CODE:
            self.modo_teste_ativo = not self.modo_teste_ativo  # Alterna o estado
            self.buffer_teclas    = ""
            self.atualizar_visual_teste()

    def atualizar_visual_teste(self):
        """
        Atualiza a aparência do botão para refletir o modo atual (teste ou produção).

        Em modo de teste: botão laranja com texto "MODO TESTE ATIVO".
        Em modo de produção: botão vermelho com texto padrão.
        """
        if self.modo_teste_ativo:
            self.btn_processar.configure(
                fg_color=COR_TESTE,
                hover_color="#d35400",
                text="MODO TESTE ATIVO"
            )
        else:
            self.btn_processar.configure(
                fg_color=COR_PRINCIPAL,
                hover_color=COR_HOVER,
                text="Colar - Gerar Planilhas"
            )

    # ===== CICLO DE PROCESSAMENTO =====

    def iniciar_processamento(self):
        """
        Inicia o processamento em uma thread separada para não travar a UI.

        Desabilita o botão e exibe a barra de progresso antes de iniciar a thread.
        """
        self.btn_processar.configure(state="disabled")  # Evita cliques duplos
        self.status_label.configure(text="Processando...", text_color="#3498db")
        self.progress.grid(row=3, column=0, padx=20, pady=10)
        self.progress.start()  # Animação indeterminada (spinner contínuo)

        # Thread daemon: encerra automaticamente com o processo principal
        threading.Thread(target=self.executar_processo, daemon=True).start()

    def executar_processo(self):
        """
        Executa o processamento principal em background.

        Chamada em thread separada. Usa self.after(0, callback) para
        atualizar a UI de forma thread-safe ao concluir.
        """
        try:
            resultado = processar_clipboard(self.modo_teste_ativo)
            # Agenda a atualização da UI na thread principal
            self.after(0, self.sucesso_final, resultado)
        except Exception as e:
            self.after(0, self.erro_final, str(e))

    def sucesso_final(self, resultado):
        """
        Chamado na thread principal após processamento bem-sucedido.

        Exibe os popups na ordem correta e abre a pasta de destino:
          1. Aviso de "nada gerado" (se aplicável)
          2. Popup de peças duplicadas (RN-041)
          3. Popup de planilhas migradas (RN-040)
          4. Abre a pasta no Explorer e exibe "Concluído!"

        Parâmetros:
            resultado (dict): Retorno de processar_clipboard().
        """
        self.progress.stop()
        self.progress.grid_forget()  # Esconde a barra de progresso
        self.btn_processar.configure(state="normal")

        # Caso especial: nenhum arquivo foi gerado e não há duplicatas
        aviso = resultado.get("aviso")
        if aviso:
            self.status_label.configure(text="Aviso: Nada gerado.", text_color="#e67e22")
            messagebox.showwarning("Atenção", aviso)
            return

        self.status_label.configure(text="Concluído!", text_color="#2ecc71")

        # RN-041: Popup de duplicados — peças que já existem em PLANOS DE CORTE 2026
        if resultado.get("bloqueados"):
            msg = "\n".join(sorted(set(resultado["bloqueados"])))
            messagebox.showwarning(
                "Peças Repetidas Identificadas",
                f"⚠️ ATENÇÃO: As peças abaixo não foram geradas pois já existem na rede:\n\n"
                f"{msg}\n\n"
                f"Copie estes arquivos manualmente para a pasta atual se necessário."
            )

        # RN-040: Popup informativo de planilhas migradas do formato antigo
        if resultado.get("migrados"):
            msg = "\n".join(resultado["migrados"])
            messagebox.showinfo(
                "Planilhas Migradas",
                f"✅ As planilhas abaixo foram migradas do formato antigo para o novo formato "
                f"e arquivadas em ANTIGOS - NÃO USAR:\n\n{msg}"
            )

        # Abre a pasta no Windows Explorer para o usuário ver os arquivos gerados
        os.startfile(resultado["pasta"])
        messagebox.showinfo("Ingecon", "Concluído!")

    def erro_final(self, mensagem):
        """
        Chamado na thread principal quando ocorre uma exceção no processamento.

        Restaura a UI ao estado inicial e exibe o erro para o usuário.

        Parâmetros:
            mensagem (str): Mensagem de erro capturada da exceção.
        """
        self.progress.stop()
        self.progress.grid_forget()
        self.btn_processar.configure(state="normal")
        self.status_label.configure(text="Erro ocorrido.", text_color="#e74c3c")
        messagebox.showerror("Erro", str(mensagem))


# ===== PONTO DE ENTRADA (execução direta do arquivo) =====
if __name__ == "__main__":
    app = AppIngecon()
    app.mainloop()