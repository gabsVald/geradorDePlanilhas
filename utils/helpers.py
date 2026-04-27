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
    """
    Limpa dados vindos do Excel/Clipboard de forma segura.
    Trata None, Booleans e Strings.
    """
    if val is None: 
        return ""
    
    # Se for Booleano (True/False), converte para string antes de qualquer coisa
    if isinstance(val, bool):
        val = str(val)
        
    v = str(val).strip()
    if v.endswith('.0'): 
        v = v[:-2]
        
    return v if v.lower() not in ['nan', 'none', 'null', ''] else ""

def converter_para_numero(valor, retornar_marcador=False, remover_unidades=True):
    """Converte strings para inteiros arredondados. Evita erros de str > int."""
    limpo = limpar(valor)
    if not limpo: 
        return "" if retornar_marcador else None
    
    # Limpeza de unidades comuns; pode ser desativado para Códigos de item alfanuméricos
    if remover_unidades:
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
    Remove: ORIG, ESS, '=', dimensões completas, dimensões truncadas,
    sufixos, parênteses vazios e traços residuais. Preserva medidas de chapa.
    """
    if not texto: return ""
    t = str(texto).upper()
    
    # 1. Remove marcações específicas que vêm sujas do PDM
    t = re.sub(r'\b(ORIG|ESS)\b', '', t)
    t = t.replace('=', '')
    
    # 2. Remove medidas completas, aceitando ponto ou vírgula
    patrao_medida = r'\b\d+(?:[.,]\d+)?\s*[xX]\s*[-]?\s*\d+(?:[.,]\d+)?(?:\s*[xX]\s*[-]?\s*\d+(?:[.,]\d+)?)?\b'
    t = re.sub(patrao_medida, '', t)
    
    # 3. Remove pedaços de medida truncados (ex: "254.7X")
    t = re.sub(r'\b\d+(?:[.,]\d+)?\s*[xX]\b', '', t)
    
    # 4. NOVO: Preserva números de medidas em parênteses, removendo apenas "MM" (ex: (3100MM) -> (3100))
    t = re.sub(r'\(\s*(\d+(?:[.,]\d+)?)\s*MM\s*\)', r'(\1)', t)
    
    # 5. Remove sufixos de espessura/comprimento órfãos (ex: 18MM) que não estão em parênteses
    t = re.sub(r'\b\d+(?:[.,]\d+)?\s*MM\b', '', t)
    
    # 6. Remove parênteses vazios ou contendo apenas espaços
    t = re.sub(r'\(\s*\)', '', t)
    
    # 7. Remove hifens seguidos de números no final da string (ex: " - 15")
    t = re.sub(r'-\s*\d+(?:[.,]\d+)?\s*$', '', t)
    
    # 8. Remove hifens soltos, no início ou no final da string
    t = re.sub(r'^\s*-\s*', '', t)
    t = re.sub(r'\s*-\s*$', '', t)
    
    # 9. Remove hifens isolados no meio do texto que ficaram vazios de ambos os lados
    t = re.sub(r'\s+-\s+', ' ', t)
    
    # 10. Limpeza final de espaços múltiplos
    t = re.sub(r'\s+', ' ', t)
    
    return t.strip().strip('-').strip()