import os
import threading
import customtkinter as ctk
from tkinter import messagebox

from utils.config import REGRAS
from utils.updater import verificar_atualizacao
from core.processador import processar_clipboard

# Identidade Visual (Cores)
COR_PRINCIPAL = "#d32732"
COR_HOVER = "#a81f28"
COR_TESTE = "#e67e22" 

ctk.set_appearance_mode("Light") 

class AppIngecon(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        verificar_atualizacao()
        
        self.versao = REGRAS.get("versao_atual", "Desconhecida")
        self.setup_window()
        self.setup_variables()
        self.create_widgets()
        
        self.bind("<Key>", self.verificar_codigo_secreto)

    def setup_window(self):
        self.title(f"Ingecon - Gerador de Planilhas V{self.versao}")
        self.geometry("480x420") 
        self.configure(fg_color="#f5f5f5") 
        self.grid_columnconfigure(0, weight=1)

    def setup_variables(self):
        self.modo_teste_ativo = False
        self.buffer_teclas = ""
        self.SECRET_CODE = "dev"

    def create_widgets(self):
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

        self.status_label = ctk.CTkLabel(self, text="Aguardando comando...", text_color="gray")
        self.status_label.grid(row=2, column=0, pady=5)

        self.progress = ctk.CTkProgressBar(
            self, 
            orientation="horizontal", 
            progress_color=COR_PRINCIPAL, 
            width=300
        )
        self.progress.set(0)

    def verificar_codigo_secreto(self, event):
        self.buffer_teclas += event.char.lower()
        if len(self.buffer_teclas) > len(self.SECRET_CODE):
            self.buffer_teclas = self.buffer_teclas[-len(self.SECRET_CODE):]
            
        if self.buffer_teclas == self.SECRET_CODE:
            self.modo_teste_ativo = not self.modo_teste_ativo
            self.buffer_teclas = ""
            self.atualizar_visual_teste()

    def atualizar_visual_teste(self):
        if self.modo_teste_ativo:
            self.btn_processar.configure(fg_color=COR_TESTE, hover_color="#d35400", text="MODO TESTE ATIVO")
        else:
            self.btn_processar.configure(fg_color=COR_PRINCIPAL, hover_color=COR_HOVER, text="Colar - Gerar Planilhas")

    def iniciar_processamento(self):
        self.btn_processar.configure(state="disabled")
        self.status_label.configure(text="Processando...", text_color="#3498db")
        self.progress.grid(row=3, column=0, padx=20, pady=10)
        self.progress.start()
        threading.Thread(target=self.executar_processo, daemon=True).start()

    def executar_processo(self):
        try:
            resultado = processar_clipboard(self.modo_teste_ativo)
            self.after(0, self.sucesso_final, resultado)
        except Exception as e: 
            self.after(0, self.erro_final, str(e))

    def sucesso_final(self, resultado): 
        self.progress.stop()
        self.progress.grid_forget()
        self.btn_processar.configure(state="normal")
        
        # BUG 3: Verifica se houve aviso de "nada gerado"
        aviso = resultado.get("aviso")
        if aviso:
            self.status_label.configure(text="Aviso: Nada gerado.", text_color="#e67e22")
            messagebox.showwarning("Atenção", aviso)
            return

        self.status_label.configure(text="Concluído!", text_color="#2ecc71")
        
        if resultado.get("migrados"):
            msg = "\n".join(sorted(set(resultado["migrados"])))
            messagebox.showinfo("Migração Detectada", f"Projetos migrados:\n\n{msg}")
            
        if resultado.get("repetidos"):
            msg = ", ".join(sorted(set(resultado["repetidos"])))
            messagebox.showwarning("Repetidos", f"Projetos {msg} já existem e foram ignorados.")
            
        os.startfile(resultado["pasta"])
        messagebox.showinfo("Ingecon", "Concluído!")

    def erro_final(self, mensagem): 
        self.progress.stop()
        self.progress.grid_forget()
        self.btn_processar.configure(state="normal")
        self.status_label.configure(text="Erro ocorrido.", text_color="#e74c3c")
        messagebox.showerror("Erro", str(mensagem))

if __name__ == "__main__":
    app = AppIngecon()
    app.mainloop()