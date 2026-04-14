import os
import re
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero

def is_prensado_migracao(cod, desc):
    """Verifica se a linha atual da planilha antiga é um título de prensado."""
    d_up = str(desc).upper()
    c_str = str(cod).strip()
    return any(g in d_up for g in REGRAS["prensados"]["descricoes_gatilho"]) or \
           c_str in REGRAS["prensados"]["codigos_gatilho"]

def extrair_dados_migracao(caminho):
    """Lê planilhas antigas e organiza itens em blocos (Normais e Prensados)."""
    try:
        if str(caminho).lower().endswith('.ods'):
            df_old = pd.read_excel(caminho, engine='odf', header=None).fillna('')
            while df_old.shape[1] < 15: df_old[df_old.shape[1]] = ''
            a3_valor = float(converter_para_numero(df_old.iloc[2, 0]) or 1.0)
            a3_valor = 1.0 if a3_valor == 0 else a3_valor
            
            blocos = []
            bloco_atual = {'tipo': 'normal', 'itens': []}
            
            for r in range(5, len(df_old)):
                # Captura dados das colunas do formato ODS antigo
                fator_bruto_raw = df_old.iloc[r, 0]
                cod = limpar(df_old.iloc[r, 12])
                desc = limpar(df_old.iloc[r, 11])
                comp = limpar(df_old.iloc[r, 2])
                larg = limpar(df_old.iloc[r, 5])
                esp = limpar(df_old.iloc[r, 7])

                # 1. VERIFICA SE É TÍTULO DE PRENSADO (Antes de filtrar vazios)
                if is_prensado_migracao(cod, desc):
                    if bloco_atual['itens']:
                        blocos.append(bloco_atual)
                    bloco_atual = {
                        'tipo': 'prensado', 
                        'prensado_info': {1: cod, 3: desc}, 
                        'itens': []
                    }
                    continue

                # 2. FILTRA LINHAS REALMENTE VAZIAS
                if not cod and not desc and not comp and not larg:
                    continue

                try:
                    f_b = float(converter_para_numero(fator_bruto_raw) or 0)
                except: f_b = 0.0
                
                item = {
                    1: cod, 15: df_old.iloc[r, 1], 8: comp,
                    16: df_old.iloc[r, 4], 10: larg,
                    12: esp, 'mat_orig': df_old.iloc[r, 8],
                    'veio_orig': df_old.iloc[r, 9], 'fita_orig': df_old.iloc[r, 10],
                    'desc_orig': desc, 'q_unitaria_fatorada': f_b / a3_valor if a3_valor > 0 else f_b,
                    'is_migrado': True
                }
                bloco_atual['itens'].append(item)
            
            if bloco_atual['itens']: blocos.append(bloco_atual)
            return blocos, a3_valor

        else: # XLSX / XLSM (Formato Novo mas precisa de migração)
            ws_d = load_workbook(caminho, data_only=True).active
            a3_valor = float(converter_para_numero(ws_d['A3'].value) or 1.0)
            a3_valor = 1.0 if a3_valor == 0 else a3_valor
            
            blocos = []
            bloco_atual = {'tipo': 'normal', 'itens': []}
            
            for r in range(6, 500):
                fator_bruto_raw = ws_d.cell(row=r, column=1).value
                cod = limpar(ws_d.cell(row=r, column=13).value)
                desc = limpar(ws_d.cell(row=r, column=12).value)
                comp = limpar(ws_d.cell(row=r, column=3).value)
                larg = limpar(ws_d.cell(row=r, column=6).value)
                esp = limpar(ws_d.cell(row=r, column=8).value)

                if is_prensado_migracao(cod, desc):
                    if bloco_atual['itens']: blocos.append(bloco_atual)
                    bloco_atual = {'tipo': 'prensado', 'prensado_info': {1: cod, 3: desc}, 'itens': []}
                    continue

                if not cod and not desc and not comp and not larg: continue

                try:
                    f_b = float(str(fator_bruto_raw or 0).replace(',', '.'))
                except: f_b = 0.0
                
                item = {
                    1: cod, 15: ws_d.cell(row=r, column=2).value, 8: comp,
                    16: ws_d.cell(row=r, column=5).value, 10: larg,
                    12: esp, 'mat_orig': ws_d.cell(row=r, column=9).value,
                    'veio_orig': ws_d.cell(row=r, column=10).value, 'fita_orig': ws_d.cell(row=r, column=11).value,
                    'desc_orig': desc, 'q_unitaria_fatorada': f_b / a3_valor if a3_valor > 0 else f_b,
                    'is_migrado': True
                }
                bloco_atual['itens'].append(item)
            
            if bloco_atual['itens']: blocos.append(bloco_atual)
            return blocos, a3_valor
            
    except Exception as e: 
        print(f"[MIGRATION ERROR] {e}")
        return [], 1.0

def mapear_rede_cache():
    cache = {}
    pasta_base = Path(REGRAS["diretorios"]["raiz"]).parent
    if pasta_base.exists():
        for root, _, files in os.walk(pasta_base):
            for f in files:
                if f.lower().endswith(('.xlsx', '.ods', '.xlsm')): cache[f] = os.path.join(root, f)
    return cache

def verificar_duplicidade_em_rede(codigo, cache_rede):
    c = str(codigo).strip()
    if not c: return None
    padrao = re.compile(rf"^{re.escape(c)}(\D|$)")
    for f, caminho in cache_rede.items():
        if padrao.match(f): return caminho
    return None