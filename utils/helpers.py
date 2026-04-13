import os
import sys
import re
from pathlib import Path

def resource_path(relative_path):
    """Gerencia caminhos de arquivos internos do executável (PyInstaller)."""
    if hasattr(sys, '_MEIPASS'):
        p_interno = Path(sys._MEIPASS) / relative_path
        if p_interno.exists(): 
            return str(p_interno)
    return str(Path(os.getcwd()) / relative_path)

def limpar(val):
    """Limpa strings de dados vindos do Excel/Clipboard."""
    if val is None: 
        return ""
    v = str(val).strip()
    if v.endswith('.0'): 
        v = v[:-2]
    return v if v.lower() not in ['nan', 'none', 'null', ''] else ""

def converter_para_numero(valor):
    """Converte strings para inteiros arredondados."""
    limpo = limpar(valor)
    if not limpo or limpo in ["-", "="]: 
        return limpo
    try:
        v_aj = limpo.replace(',', '.')
        val_float = float(v_aj)
        return int(val_float + 0.5) if val_float >= 0 else int(val_float - 0.5)
    except Exception: 
        return limpo

def limpar_material_rigoroso(texto):
    """Remove termos técnicos e dimensões da descrição do material."""
    if not texto: 
        return ""
    t = re.sub(r'\b(ORIG|ESS)\b', '', str(texto), flags=re.IGNORECASE).replace('=', '')
    t = re.sub(r'\s*\b\d+(?:[\.,]\d+)?\s*[xX].*$', '', t, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', t).strip(' -')