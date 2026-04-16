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

def converter_para_numero(valor, retornar_marcador=False):
    """Converte strings para inteiros arredondados. Evita erros de str > int."""
    limpo = limpar(valor)
    if not limpo: 
        return "" if retornar_marcador else None
    if limpo in ["-", "="]: 
        return limpo if retornar_marcador else None
    try:
        v_aj = limpo.replace(',', '.')
        val_float = float(v_aj)
        return int(val_float + 0.5) if val_float >= 0 else int(val_float - 0.5)
    except Exception: 
        return limpo if retornar_marcador else None

def limpar_material_rigoroso(texto):
    """
    Limpa a string de material para exibição na planilha.
    Remove: ORIG, ESS (word boundary), '=', e dimensões no padrão NxN ou NxNxN.
    NÃO remove nomes de materiais (MDF, CRU, etc.) — esses já são filtrados
    pelo f_valido() antes de chegar aqui.
    """
    t = str(texto).upper()
    t = t.replace('=', '')
    t = re.sub(r'\bORIG\b', '', t)
    t = re.sub(r'\bESS\b', '', t)
    t = re.sub(r'\b\d+([,\.]\d+)?(X\d+([,\.]\d+)?){1,2}(MM)?\b', '', t)
    t = re.sub(r'\b\d+([,\.]\d+)?\s*MM\b', '', t)
    return re.sub(r'\s+', ' ', t).strip()