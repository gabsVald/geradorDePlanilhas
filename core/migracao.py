import os
import re
import json
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero

# Carrega mapeamento MP → material atualizado
_MAP_MP_PATH = Path(__file__).parent.parent / "mapeamento.json"
try:
    with open(_MAP_MP_PATH, encoding='utf-8') as _f:
        _MAP_MP = json.load(_f)
except Exception:
    _MAP_MP = {}

def _material_por_mp(mp_raw: str, fallback: str) -> str:
    """Retorna o nome atualizado do material a partir do código MP.
    Se não encontrar no mapeamento, usa o fallback (nome original do arquivo)."""
    mp = str(mp_raw).strip()
    return _MAP_MP.get(mp, fallback)

def is_prensado_migracao(cod, desc, col_a=""):
    d_up = str(desc).upper()
    c_str = str(cod).strip()
    a_up = str(col_a).upper()
    
    gatilhos = REGRAS["prensados"]["descricoes_gatilho"]
    codigos = REGRAS["prensados"]["codigos_gatilho"]
    
    return any(g in d_up for g in gatilhos) or \
           any(g in a_up for g in gatilhos) or \
           c_str in codigos

def extrair_dados_migracao(caminho):
    try:
        termos_ignorar = ["PROGRAMAÇÃO", "DATA", "UN", "MEDIDA", "MATERIAL", "PROCESSO", "DESCRIÇÃO", "CÓDIGO", "QNT", "PROG."]
#        
        if str(caminho).lower().endswith('.ods'):
            df_old = pd.read_excel(caminho, engine='odf', header=None).fillna('')
            while df_old.shape[1] < 15: df_old[df_old.shape[1]] = ''

            # A3 está na col 1 da linha 2 (não col 0)
            # A3 pode estar em col 0 (sem coluna QNT) ou col 1 (com coluna QNT)
            _a3_raw = df_old.iloc[2, 0] if str(df_old.iloc[2, 0]).strip() not in ['', 'nan', 'QNT'] else df_old.iloc[2, 1]
            a3_valor = float(converter_para_numero(_a3_raw) or 1.0)
            a3_valor = 1.0 if a3_valor == 0 else a3_valor

            # Código da peça está no título: row 1, col 2, antes do " - "
            titulo_raw = limpar(df_old.iloc[1, 2])
            cod_titulo = titulo_raw.split(' - ', 1)[0].strip() if ' - ' in titulo_raw else ''

            blocos = []
            bloco_atual = {'tipo': 'normal', 'itens': []}

            for r in range(5, len(df_old)):
                c0 = limpar(df_old.iloc[r, 0])
                # col 11 = código do material (MP), não o código da peça
                # col 9  = descrição do material
                cod  = cod_titulo                     # código da peça vem do título
                desc = limpar(df_old.iloc[r, 9])      # material/descrição
                
                # Filtro de termos: c0 com todos; desc/cod só com termos longos (>=4 chars)
                # Evita 'UN' dentro de 'FUNDO' mas captura 'PROGRAMAÇÃO', 'MATERIAL' etc.
                _termos_longos = [t for t in termos_ignorar if len(t) >= 4]
                if any(t in str(c0).upper() for t in termos_ignorar) or \
                   any(t in str(desc).upper() for t in _termos_longos) or \
                   any(t in str(cod).upper() for t in _termos_longos):
                    continue

                if is_prensado_migracao(cod, desc, c0):
                    if bloco_atual['itens']: blocos.append(bloco_atual)
                    f_cod, f_desc = cod, desc
                    if not cod and " - " in c0:
                        partes = c0.split(" - ", 1)
                        f_cod, f_desc = partes[0].strip(), partes[1].strip()
                    elif not cod: f_desc = c0
                    bloco_atual = {'tipo': 'prensado', 'prensado_info': {1: f_cod, 3: f_desc}, 'itens': []}
                    continue

                comp = limpar(df_old.iloc[r, 3])   # col 3 no ODS
                larg = limpar(df_old.iloc[r, 6])   # col 6 no ODS
                # Linha sem dados úteis: pula se não tem dimensões nem descrição
                if not comp and not larg and not desc: continue

                # col0 comportamento depende do layout:
                # - Com UN (col1 numérico): col0 é qtd absoluta → dividir por a3
                # - Sem UN: col0 já é o fator unitário direto
                try: f_b_raw = float(converter_para_numero(df_old.iloc[r, 0]) or 0)
                except: f_b_raw = 0.0

                # Layout ODS varia: col1 com UN (número) → dims em 3/6/8
                # col1 vazio, '-' ou '=' = sem UN → dims em 2/5/7
                _col1_raw = str(df_old.iloc[r, 1]).strip()
                _tem_un = bool(converter_para_numero(_col1_raw)) if _col1_raw not in ['', 'nan'] else False
                _dc, _dl, _da, _dm, _df, _dd = (3,6,8,9,10,11) if _tem_un else (2,5,7,8,10,11)
                # Fita de borda lateral: col1 com '-' ou '=' no ODS
                _fita_lat = _col1_raw if _col1_raw in ['-', '='] else None
                # Fita de borda topo: col4 com '-' ou '=' no ODS
                _fita_top = str(df_old.iloc[r, 4]).strip() if str(df_old.iloc[r, 4]).strip() in ['-', '='] else None
                item = {
                    # Regra: se estava vazio no ODS, deve ficar vazio no novo arquivo
                    1: limpar(df_old.iloc[r, 12]),   # código do item (col 12) — pode ser vazio
                    8:  limpar(df_old.iloc[r, _dc]),  # Comprimento
                    10: limpar(df_old.iloc[r, _dl]),  # Largura
                    12: limpar(df_old.iloc[r, _da]),  # Altura/espessura
                    'mat_orig':  limpar(df_old.iloc[r, _dm]),
                    'veio_orig': None,                             # ODS não tem coluna de veio
                    'fita_orig': limpar(df_old.iloc[r, _df]),      # processo/fita
                    'fita_lat':  _fita_lat,                        # '-' ou '=' da col1 (borda lateral)
                    'fita_top':  _fita_top,                        # '-' ou '=' da col4 (borda topo)
                    'desc_orig': limpar(df_old.iloc[r, _dd]),      # descrição/código MP
                    # Com UN (col1 numérico): col0 é fator unitário → NÃO dividir
                    # Sem UN (col1 vazio):    col0 é qtd absoluta  → dividir por a3
                    'q_unitaria_fatorada': (f_b_raw if _tem_un else (f_b_raw / a3_valor if a3_valor > 0 else f_b_raw)),
                    'is_migrado': True
                }
                bloco_atual['itens'].append(item)
            
            if bloco_atual['itens']: blocos.append(bloco_atual)
            return blocos, a3_valor

        else:
            ws_d = load_workbook(caminho, data_only=True).active
            a3_valor = float(converter_para_numero(ws_d['A3'].value) or 1.0)
            a3_valor = 1.0 if a3_valor == 0 else a3_valor
            
            blocos = []
            bloco_atual = {'tipo': 'normal', 'itens': []}
            
            for r in range(1, 500):
                c0_raw = ws_d.cell(row=r, column=1).value
                c0 = limpar(c0_raw)
                cod = limpar(ws_d.cell(row=r, column=13).value)
                desc = limpar(ws_d.cell(row=r, column=12).value)

                # Filtro de termos: c0 com todos; desc/cod só com termos longos (>=4 chars)
                # Evita 'UN' dentro de 'FUNDO' mas captura 'PROGRAMAÇÃO', 'MATERIAL' etc.
                _termos_longos = [t for t in termos_ignorar if len(t) >= 4]
                if any(t in str(c0).upper() for t in termos_ignorar) or \
                   any(t in str(desc).upper() for t in _termos_longos) or \
                   any(t in str(cod).upper() for t in _termos_longos):
                    continue

                if is_prensado_migracao(cod, desc, c0):
                    if bloco_atual['itens']: blocos.append(bloco_atual)
                    f_cod, f_desc = cod, desc
                    if not cod and " - " in c0:
                        partes = c0.split(" - ", 1)
                        f_cod, f_desc = partes[0].strip(), partes[1].strip()
                    elif not cod: f_desc = c0
                    bloco_atual = {'tipo': 'prensado', 'prensado_info': {1: f_cod, 3: f_desc}, 'itens': []}
                    continue

                comp = limpar(ws_d.cell(row=r, column=3).value)
                larg = limpar(ws_d.cell(row=r, column=6).value)
                if not cod and not desc and not comp and not larg: continue

                try: f_b = float(converter_para_numero(c0_raw) or 0)
                except: f_b = 0.0
                
                # Fita de borda XLSX: col 2 = lateral (B), col 5 = topo (E)
                _xlsx_fita_lat = ws_d.cell(row=r, column=2).value
                _xlsx_fita_top = ws_d.cell(row=r, column=5).value
                _fita_lat_xlsx = str(_xlsx_fita_lat).strip() if str(_xlsx_fita_lat).strip() in ['-', '='] else None
                _fita_top_xlsx = str(_xlsx_fita_top).strip() if str(_xlsx_fita_top).strip() in ['-', '='] else None
                # col 9=VEIO, col 10=PROCESSO, col 11=DESCRIÇÃO
                _veio_xlsx   = ws_d.cell(row=r, column=10).value
                _fita_xlsx   = limpar(ws_d.cell(row=r, column=11).value) or ''
                item = {
                    1: cod, 8: comp, 10: larg,
                    12: limpar(ws_d.cell(row=r, column=8).value),
                    'mat_orig':  limpar(ws_d.cell(row=r, column=9).value),
                    'veio_orig': _veio_xlsx,
                    'fita_orig': _fita_xlsx,
                    'fita_lat':  _fita_lat_xlsx,
                    'fita_top':  _fita_top_xlsx,
                    'desc_orig': desc,
                    'q_unitaria_fatorada': f_b / a3_valor if a3_valor > 0 else f_b,
                    'is_migrado': True
                }
                bloco_atual['itens'].append(item)
            
            if bloco_atual['itens']: blocos.append(bloco_atual)
            return blocos, a3_valor
            
    except Exception as e: 
        print(f"[MIGRATION ERROR] {e}")
        return [], 1.0

def mapear_rede_cache():
    cache = [] # Segurança: Lista em vez de dicionário para evitar sobreposição de nomes idênticos
    pasta_base = Path(REGRAS["diretorios"]["raiz"]).parent
    if pasta_base.exists():
        for root, _, files in os.walk(pasta_base):
            for f in files:
                if f.lower().endswith(('.xlsx', '.ods', '.xlsm')): 
                    cache.append((f, os.path.join(root, f)))
    return cache

def verificar_duplicidade_em_rede(codigo, cache_rede):
    c = str(codigo).strip()
    if not c: return None
    # Segurança: Regex exige que o código seja seguido de um separador válido ou fim de string
    padrao = re.compile(rf"^{re.escape(c)}(\s|-|_|\.|$)")
    for f, caminho in cache_rede:
        if padrao.match(f): return caminho
    return None