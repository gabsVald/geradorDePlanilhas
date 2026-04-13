import os
import re
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

# Importamos as nossas regras e ferramentas
from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero

def mapear_rede_cache():
    """Cria um dicionário de ficheiros existentes na rede para busca rápida."""
    cache = {}
    
    # A pasta de verificação original era a raiz do setor (06 - PLANOS DE CORTE ATUALIZADOS)
    # Pegamos no diretório de 2026 do JSON e recuamos uma pasta (parent) para ter o caminho base.
    pasta_base_verificacao = Path(REGRAS["diretorios"]["raiz"]).parent
    pastas_verificacao = [pasta_base_verificacao]
    
    for p in pastas_verificacao:
        if not p.exists(): 
            continue
        for root, _, files in os.walk(p):
            for f in files:
                if f.lower().endswith(('.xlsx', '.ods', '.xlsm')): 
                    cache[f] = os.path.join(root, f)
    return cache

def verificar_duplicidade_em_rede(codigo, cache_rede):
    """Verifica se o código do projeto já possui uma planilha gerada na rede."""
    c = str(codigo).strip()
    if not c: 
        return None
    padrao = re.compile(rf"^{re.escape(c)}(\D|$)")
    
    for f, caminho in cache_rede.items():
        if padrao.match(f): 
            return caminho
    return None

def extrair_dados_migracao(caminho):
    """Extrai dados de planilhas antigas (ODS ou XLSX) para o novo formato."""
    try:
        if str(caminho).lower().endswith('.ods'):
            df_old = pd.read_excel(caminho, engine='odf', header=None).fillna('')
            while df_old.shape[1] < 15: 
                df_old[df_old.shape[1]] = ''
                
            try: 
                a3_valor = float(converter_para_numero(df_old.iloc[2, 0]) or 1.0)
            except Exception: 
                a3_valor = 1.0
                
            if a3_valor == 0: 
                a3_valor = 1.0
            
            itens = []
            for r in range(5, len(df_old)):
                cod = limpar(df_old.iloc[r, 12])
                desc = limpar(df_old.iloc[r, 11])
                comp = limpar(df_old.iloc[r, 2])
                larg = limpar(df_old.iloc[r, 5])
                
                if not cod and not desc and not comp and not larg: 
                    continue
                if str(cod).upper() == "X" and not desc and not comp and not larg: 
                    continue
                
                try: 
                    fator_bruto = float(converter_para_numero(df_old.iloc[r, 0]) or 0)
                except Exception: 
                    fator_bruto = 0.0
                    
                fator_unitario = fator_bruto / a3_valor if a3_valor > 0 else fator_bruto
                
                item = {
                    1: cod, 15: df_old.iloc[r, 1], 8: df_old.iloc[r, 2],
                    16: df_old.iloc[r, 4], 10: df_old.iloc[r, 5],
                    12: df_old.iloc[r, 7], 'mat_orig': df_old.iloc[r, 8],
                    'veio_orig': df_old.iloc[r, 9], 'fita_orig': df_old.iloc[r, 10],
                    'desc_orig': desc, 'q_unitaria_fatorada': fator_unitario, 'is_migrado': True
                }
                itens.append(item)
            return itens, a3_valor
            
        else: # Trata formatos Excel padrão (.xlsx, .xlsm)
            wb_data = load_workbook(caminho, data_only=True)
            ws_d = wb_data.active
            
            try: 
                a3_valor = float(converter_para_numero(ws_d['A3'].value) or 1.0)
            except Exception: 
                a3_valor = 1.0
                
            if a3_valor == 0: 
                a3_valor = 1.0
            
            itens = []
            for r in range(6, 500):
                cod = limpar(ws_d.cell(row=r, column=13).value)
                desc = limpar(ws_d.cell(row=r, column=12).value)
                comp = limpar(ws_d.cell(row=r, column=3).value)
                larg = limpar(ws_d.cell(row=r, column=6).value)
                
                if not cod and not desc and not comp and not larg: 
                    continue
                if str(cod).upper() == "X" and not desc and not comp and not larg: 
                    continue
                
                try: 
                    fator_bruto = float(str(ws_d.cell(row=r, column=1).value or 0).replace(',', '.'))
                except Exception: 
                    fator_bruto = 0.0
                    
                fator_unitario = fator_bruto / a3_valor if a3_valor > 0 else fator_bruto
                
                item = {
                    1: cod, 15: ws_d.cell(row=r, column=2).value, 8: ws_d.cell(row=r, column=3).value,
                    16: ws_d.cell(row=r, column=5).value, 10: ws_d.cell(row=r, column=6).value,
                    12: ws_d.cell(row=r, column=8).value, 'mat_orig': ws_d.cell(row=r, column=9).value,
                    'veio_orig': ws_d.cell(row=r, column=10).value, 'fita_orig': ws_d.cell(row=r, column=11).value,
                    'desc_orig': desc, 'q_unitaria_fatorada': fator_unitario, 'is_migrado': True
                }
                itens.append(item)
            return itens, a3_valor
            
    except Exception as e: 
        print(f"[MIGRATION WARN] Erro ao extrair de {caminho}: {e}")
        return [], 1.0