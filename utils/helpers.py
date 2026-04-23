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
    
    # Limpeza de unidades comuns que vem do PDM ou planilhas antigas
    limpo = re.sub(r'(?i)\s*(pç|pc|un|und|unid|unidade)\.*', '', limpo)
    
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
    Remove: ORIG, ESS, '=', dimensões completas, dimensões truncadas e traços.
    """
    if not texto: return ""
    t = str(texto).upper()
    
    # 1. Remove marcações específicas que vêm sujas do PDM
    t = re.sub(r'\b(ORIG|ESS)\b', '', t)
    t = t.replace('=', '')
    
    # 2. Remove medidas completas, inclusive com hifens inseridos por erro (ex: 440X-260X15 ou 15,5X20,5)
    # Permite casas decimais, espaços variáveis e hifens opcionais logo após o X
    t = re.sub(r'\b\d+(?:,\d+)?\s*[xX]\s*[-]?\s*\d+(?:,\d+)?(?:\s*[xX]\s*[-]?\s*\d+(?:,\d+)?)?\b', '', t)
    
    # 3. Remove pedaços de medida truncados (ex: "440X " ou "15X")
    t = re.sub(r'\b\d+\s*[xX]\b', '', t)
    
    # 4. Remove sufixos de espessura (ex: 15MM ou 15 MM)
    t = re.sub(r'\b\d+\s*MM\b', '', t)
    
    # 5. Remove hifens seguidos de números no final da string (ex: " - 15")
    t = re.sub(r'-\s*\d+\s*$', '', t)
    
    # 6. Remove hifens soltos, no início ou no final da string
    t = re.sub(r'^\s*-\s*', '', t)
    t = re.sub(r'\s*-\s*$', '', t)
    
    # 7. Remove hifens isolados no meio do texto que ficaram vazios de ambos os lados
    t = re.sub(r'\s+-\s+', ' ', t)
    
    # 8. Remove espaços múltiplos para evitar colunas dessincronizadas visualmente
    t = re.sub(r'\s+', ' ', t)
    
    return t.strip().strip('-').strip()