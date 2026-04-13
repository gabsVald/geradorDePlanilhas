import json
import sys
from pathlib import Path

def obter_caminho_base():
    """Garante que encontra o JSON rodando no código ou no executável compilado."""
    if getattr(sys, 'frozen', False):
        # Se for .exe, o json deve estar na mesma pasta do .exe
        return Path(sys.executable).parent
    # Se for código (main.py), o json está uma pasta acima de 'utils'
    return Path(__file__).parent.parent

def carregar_regras():
    caminho_json = obter_caminho_base() / "regras.json"
    try:
        with open(caminho_json, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"[ERRO CRÍTICO] Não foi possível carregar regras.json: {e}")
        return {}

# Variável global que o resto do sistema vai importar
REGRAS = carregar_regras()