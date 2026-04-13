import os
import sys
from pathlib import Path
from tkinter import messagebox
from utils.config import REGRAS

def verificar_atualizacao():
    """Checa se existe uma versão mais recente no servidor."""
    try:
        # Puxamos os caminhos direto do nosso regras.json
        dir_sistema = Path(REGRAS["diretorios"]["raiz"]) / REGRAS["diretorios"]["nome_pasta_sistema"]
        arq_versao = dir_sistema / "version.txt"
        exe_servidor = dir_sistema / "Gerador_Planilhas_Ingecon.exe"
        
        versao_atual = REGRAS["versao_atual"]

        if arq_versao.exists():
            with open(arq_versao, "r", encoding="utf-8") as f: 
                v_serv = f.read().strip()
            
            if v_serv != versao_atual:
                if messagebox.askyesno("Atualização", f"Nova versão {v_serv} disponível. Atualizar?"):
                    executar_patch(exe_servidor)
                    os._exit(0)
    except Exception as e: 
        print(f"[UPDATE WARN] {e}")

def executar_patch(exe_servidor):
    """Gera e executa um script batch para substituir o executável atual."""
    c_exe = sys.executable
    n_exe = os.path.basename(c_exe)
    
    bat = (
        f'@echo off\n:loop\ntaskkill /f /im "{n_exe}" >nul 2>&1\n'
        f'del /f /q "{c_exe}" >nul 2>&1\n'
        f'if exist "{c_exe}" (timeout /t 1 >nul\ngoto loop)\n'
        f'copy /y "{exe_servidor}" "{c_exe}"\nstart "" "{c_exe}"\nexit'
    )
    p_bat = Path(os.environ["TEMP"]) / f"patch_ingecon_{os.getpid()}.bat"
    
    with open(p_bat, "w", encoding="utf-8") as f: 
        f.write(bat)
        
    os.startfile(p_bat)