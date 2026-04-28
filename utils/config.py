"""
================================================================================
utils/config.py — Carregamento de Configurações Globais
================================================================================
Carrega o arquivo regras.json e expõe a variável global REGRAS.

O sistema distingue dois modos de execução:
  - Desenvolvimento: regras.json está na raiz do projeto (../utils/)
  - Executável (.exe): regras.json está na mesma pasta do .exe (sys.executable)

A variável REGRAS é importada por todos os módulos que precisam de configuração.
================================================================================
"""

import json
import sys
from pathlib import Path


def obter_caminho_base():
    """
    Determina o diretório raiz onde o regras.json deve ser encontrado.

    Em modo executável (.exe compilado com PyInstaller), o .json deve
    estar junto ao executável para ser editável pelo usuário.
    Em modo desenvolvimento, está na raiz do projeto.

    Retorna:
        Path: Diretório base para localizar o regras.json.
    """
    if getattr(sys, 'frozen', False):
        # Executável PyInstaller: mesmo diretório do .exe
        return Path(sys.executable).parent
    # Desenvolvimento: sobe um nível a partir de utils/ para chegar na raiz
    return Path(__file__).parent.parent


def carregar_regras():
    """
    Lê e parseia o arquivo regras.json.

    Em caso de falha (arquivo não encontrado, JSON inválido, permissão negada),
    imprime um aviso e retorna um dicionário vazio. O sistema continuará
    com comportamento degradado em vez de travar completamente.

    Retorna:
        dict: Dicionário com todas as regras do sistema, ou {} em caso de erro.
    """
    caminho_json = obter_caminho_base() / "regras.json"
    try:
        with open(caminho_json, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"[ERRO CRÍTICO] Não foi possível carregar regras.json: {e}")
        return {}


# Variável global importada por todos os módulos do sistema.
# Carregada uma única vez ao iniciar o aplicativo.
REGRAS = carregar_regras()