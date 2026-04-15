import os
import re
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero

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
            a3_valor = float(converter_para_numero(df_old.iloc[2, 0]) or 1.0)
            a3_valor = 1.0 if a3_valor == 0 else a3_valor
            
            blocos = []
            bloco_atual = {'tipo': 'normal', 'itens': []}
            
            for r in range(5, len(df_old)):
                c0 = limpar(df_old.iloc[r, 0])
                cod = limpar(df_old.iloc[r, 12])
                desc = limpar(df_old.iloc[r, 11])
                
                if any(t in str(c0).upper() for t in termos_ignorar) or \
                   any(t in str(cod).upper() for t in termos_ignorar) or \
                   any(t in str(desc).upper() for t in termos_ignorar):
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

                comp = limpar(df_old.iloc[r, 2])
                larg = limpar(df_old.iloc[r, 5])
                if not cod and not desc and not comp and not larg: continue

                try: f_b = float(converter_para_numero(df_old.iloc[r, 0]) or 0)
                except: f_b = 0.0
                
                item = {
                    1: cod, 15: df_old.iloc[r, 1], 8: comp, 16: df_old.iloc[r, 4], 10: larg,
                    12: limpar(df_old.iloc[r, 7]), 'mat_orig': limpar(df_old.iloc[r, 8]), 
                    'veio_orig': df_old.iloc[r, 9], 'fita_orig': df_old.iloc[r, 10], 
                    'desc_orig': desc, 'q_unitaria_fatorada': f_b / a3_valor if a3_valor > 0 else f_b, 
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

                if any(t in str(c0).upper() for t in termos_ignorar) or \
                   any(t in str(cod).upper() for t in termos_ignorar) or \
                   any(t in str(desc).upper() for t in termos_ignorar):
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
                
                item = {
                    1: cod, 15: ws_d.cell(row=r, column=2).value, 8: comp, 16: ws_d.cell(row=r, column=5).value, 
                    10: larg, 12: limpar(ws_d.cell(row=r, column=8).value), 
                    'mat_orig': limpar(ws_d.cell(row=r, column=9).value), 
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