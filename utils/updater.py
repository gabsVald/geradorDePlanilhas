"""
================================================================================
utils/updater.py — Sistema de Atualização Automática
================================================================================
Verifica ao iniciar o aplicativo se existe uma versão mais recente disponível
no servidor de rede configurado em regras.json.

Mecanismo:
  1. Lê o arquivo version.txt no servidor
  2. Compara com a versão atual (REGRAS['versao_atual'])
  3. Se diferente, exibe diálogo de confirmação
  4. Se aceito, gera um .bat que substitui o .exe e reinicia o aplicativo

O .bat é necessário porque o processo não pode sobrescrever a si mesmo
enquanto está em execução — o .bat mata o processo primeiro.
================================================================================
"""

import os
import sys
from pathlib import Path
from tkinter import messagebox

from utils.config import REGRAS  # Configurações globais (versão atual, caminhos)


def verificar_atualizacao():
    """
    Verifica se existe uma versão mais recente do aplicativo disponível na rede.

    Fluxo:
      1. Localiza version.txt na pasta do sistema na rede (regras.json → diretorios)
      2. Compara a versão do servidor com REGRAS['versao_atual'] por igualdade de string
         (qualquer diferença, incluindo capitalização, aciona a atualização)
      3. Exibe diálogo sim/não para o usuário
      4. Se aceito: gera e executa o patch via _executar_patch()

    Falhas de leitura (servidor offline, arquivo ausente, sem permissão)
    são silenciosas — o sistema continua normalmente sem atualizar.
    """
    try:
        dir_sistema  = Path(REGRAS["diretorios"]["raiz"]) / REGRAS["diretorios"]["nome_pasta_sistema"]
        arq_versao   = dir_sistema / "version.txt"    # Arquivo de versão no servidor
        exe_servidor = dir_sistema / "Gerador_Planilhas_Ingecon.exe"  # .exe atualizado no servidor

        versao_atual = REGRAS["versao_atual"]  # Ex: "2.2.12"

        if arq_versao.exists():
            with open(arq_versao, "r", encoding="utf-8") as f:
                v_serv = f.read().strip()  # Ex: "2.2.13"

            if v_serv != versao_atual:
                # Exibe diálogo de confirmação antes de prosseguir
                if messagebox.askyesno("Atualização", f"Nova versão {v_serv} disponível. Atualizar?"):
                    executar_patch(exe_servidor)
                    os._exit(0)  # Encerra imediatamente para o .bat assumir o controle

    except Exception as e:
        # Falha silenciosa: log no console mas não bloqueia a inicialização
        print(f"[UPDATE WARN] {e}")


def executar_patch(exe_servidor):
    """
    Gera e executa um script .bat para substituir o executável atual pela versão do servidor.

    O .bat é necessário porque um processo Windows não pode sobrescrever seu próprio .exe
    enquanto está em execução. O script:
      1. Encerra o processo atual (taskkill)
      2. Aguarda até o arquivo ser liberado (loop com timeout)
      3. Copia o novo .exe do servidor para o local do atual
      4. Reinicia o aplicativo com o novo .exe

    O .bat é gerado em %TEMP% com nome único (PID do processo) para evitar conflitos
    em caso de múltiplas instâncias simultâneas.

    Parâmetros:
        exe_servidor (Path): Caminho completo do .exe atualizado disponível no servidor.
    """
    c_exe = sys.executable                    # Caminho do .exe atual sendo executado
    n_exe = os.path.basename(c_exe)           # Apenas o nome do arquivo (ex: "Gerador_Planilhas_Ingecon.exe")

    bat = (
        f'@echo off\n'
        f':loop\n'
        f'taskkill /f /im "{n_exe}" >nul 2>&1\n'       # Força encerramento do processo atual
        f'del /f /q "{c_exe}" >nul 2>&1\n'              # Tenta deletar o .exe antigo
        f'if exist "{c_exe}" ('
        f'timeout /t 1 >nul\ngoto loop)\n'              # Aguarda até o arquivo ser liberado
        f'copy /y "{exe_servidor}" "{c_exe}"\n'         # Copia o novo .exe do servidor
        f'start "" "{c_exe}"\n'                          # Reinicia com o novo .exe
        f'exit'
    )

    # Salva o .bat no diretório temporário do sistema com nome único por PID
    p_bat = Path(os.environ["TEMP"]) / f"patch_ingecon_{os.getpid()}.bat"

    with open(p_bat, "w", encoding="utf-8") as f:
        f.write(bat)

    os.startfile(p_bat)  # Executa o .bat de forma assíncrona; o processo atual encerra logo após