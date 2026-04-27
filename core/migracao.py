import os
import re
import json
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from utils.config import REGRAS
from utils.helpers import limpar

_MAP_MP_PATH = Path(__file__).parent.parent / "mapeamento.json"
try:
    with open(_MAP_MP_PATH, encoding='utf-8') as _f:
        _MAP_MP = json.load(_f)
except Exception:
    _MAP_MP = {}

def is_pdm_code(s):
    """Filtro rígido para garantir que o Código não é confundido com outros textos."""
    s = str(s).strip().upper()
    if not s: return False
    # Aceita PÇ 1, PÇ, PC 12
    if re.match(r'^P[CÇ]\s*\d*$', s): return True
    # Aceita códigos Ingecon começando por 11, 15, 0
    if re.match(r'^(11|15|0)\d{4,}[A-Z]?$', s): return True
    # Se for um número muito alto, provavelmente é o código numérico cru
    try:
        val = float(s.replace(',', '.'))
        if val > 50000: return True 
    except: pass
    return False

def extrair_dados_linha_inteligente(linha_raw, a3_valor, cod_fallback=''):
    """
    Sistema Estruturado por Fila: Mapeia as colunas I, J, K, L, M de forma 
    exata, mas resiste a colunas mescladas ou deletadas pelo projetista.
    """
    linha = [limpar(x) for x in linha_raw]
    
    # Preenche a linha com vazios para garantir que as colunas M e N existam
    while len(linha) < 25: linha.append("")
    
    # 1. Encontrar o Comprimento (A âncora C)
    comp_idx = -1
    for i in range(5):
        val = re.sub(r'(?i)\s*(pç|pc|un|und|unid|unidade)\.*', '', str(linha[i])).strip()
        try:
            if float(val.replace(',', '.')) >= 50:
                comp_idx = i
                break
        except: pass
        
    if comp_idx == -1: return None
    
    # Offset lida com planilhas que apagaram a Coluna A
    offset = comp_idx - 2
    
    # 2. Extração QTD e Medidas
    qtd_idx = 0 + offset
    if qtd_idx >= 0:
        qtd_str = re.sub(r'(?i)\s*(pç|pc|un|und|unid|unidade)\.*', '', str(linha[qtd_idx])).strip()
        try: q_val = float(qtd_str.replace(',', '.'))
        except: q_val = 0
    else:
        q_val = 0
        
    q_unit = q_val / a3_valor if a3_valor > 0 else q_val
    
    if 1 + offset >= 0:
        q1_str = re.sub(r'(?i)\s*(pç|pc|un|und|unid|unidade)\.*', '', str(linha[1 + offset])).strip()
        try: 
            q1_val = float(q1_str.replace(',', '.'))
            if q1_val > 0 and q1_val < 50: q_unit = q1_val
        except: pass

    comp = linha[2 + offset]
    larg = linha[5 + offset]
    esp  = linha[7 + offset]
    fita_lat = linha[1 + offset] if linha[1 + offset] in ['-', '='] else None
    fita_top = linha[4 + offset] if linha[4 + offset] in ['-', '='] else None
    
    # ---------------------------------------------------------------------
    # 3. Extração Geográfica Exata (I, J, K, L, M)
    # ---------------------------------------------------------------------
    mat = str(linha[8 + offset]).strip()
    veio_val = str(linha[9 + offset]).strip().upper()
    
    termos_proc = ['SEC', 'SEC-LAM', 'SERRA', 'SERRA-LAM', 'LAM', 'FITA', 'BORDA', 'BORD', 'USI', 'USINAGEM', 'CRU']
    is_proc = lambda x: any(t in x.upper() for t in termos_proc)

    fita = ""
    desc = ""
    cod = ""

    # Sistema de Fila (Lida com o arrastamento de colunas se algo foi apagado)
    if is_proc(veio_val):
        veio = None
        fita = veio_val
        # A fila começa no índice 10 (K) porque o Processo comeu a coluna do Veio
        fila = [str(linha[i + offset]).strip() for i in range(10, 14)]
    else:
        veio = 1 if veio_val in ['1', '1.0', 'S', 'SIM', 'X'] else None
        # A fila começa no índice 10 (K) e vai até à N
        fila = [str(linha[i + offset]).strip() for i in range(10, 15)]
        
        if is_proc(fila[0]):
            fita = fila.pop(0)
        elif fila[0] == "":
            fita = fila.pop(0) 
        else:
            pass # Processo foi deletado e os dados andaram para a esquerda
            
    # O que sobrou na Fila é, por ordem: Coluna L (Descrição) e Coluna M (Código)
    desc = fila.pop(0)
    c1 = fila.pop(0)
    c2 = fila.pop(0)
    
    # 4. Avaliação L e M para evitar junção do Código
    if is_pdm_code(c2):
        cod = c2
        desc = f"{desc} {c1}".strip()
    elif is_pdm_code(c1):
        cod = c1
        if c2: desc = f"{desc} {c2}".strip() # c2 seria apenas um texto extra perdido
    else:
        cod = c1
        if c2: desc = f"{desc} {c2}".strip()

    # 5. Fallback para Coluna Mesclada (Se o Código estiver dentro da L junto com a Descrição)
    if " - " in desc and not is_pdm_code(cod):
        parts = desc.rsplit(" - ", 1)
        if is_pdm_code(parts[1]):
            desc = parts[0].strip()
            cod = parts[1].strip()
        elif is_pdm_code(parts[0]):
            desc = parts[1].strip()
            cod = parts[0].strip()

    # Prevenção Final
    if is_pdm_code(desc) and not is_pdm_code(cod):
        cod = desc
        desc = ""
        
    if not cod: cod = cod_fallback

    return {
        1: cod, 
        8: comp, 
        10: larg, 
        12: esp,
        'mat_orig': mat, 
        'veio_orig': veio,
        'fita_orig': fita, 
        'fita_lat': fita_lat, 
        'fita_top': fita_top,
        'desc_orig': desc, 
        'q_unitaria_fatorada': q_unit, 
        'is_migrado': True
    }

def extrair_dados_migracao(caminho):
    try:
        termos_ignorar = ["PROGRAMAÇÃO", "DATA", "UN", "MEDIDA", "MATERIAL", "PROCESSO", "DESCRIÇÃO", "CÓDIGO", "QNT", "PROG."]
        
        blocos = []
        bloco_atual = {'tipo': 'normal', 'itens': []}
        linhas_raw = []
        a3_valor = 1.0
        cod_titulo = ''

        # Leitura da Planilha
        if str(caminho).lower().endswith('.ods'):
            df_old = pd.read_excel(caminho, engine='odf', header=None).fillna('')
            while df_old.shape[1] < 16: df_old[df_old.shape[1]] = ''
            
            linha_a3 = [limpar(df_old.iloc[2, c]) for c in range(df_old.shape[1])]
            for cell in linha_a3:
                try:
                    num = float(str(cell).replace(',', '.'))
                    if num > 0: a3_valor = num; break
                except: pass
            
            linha_titulo = [limpar(df_old.iloc[1, c]) for c in range(df_old.shape[1])]
            for cell in linha_titulo:
                if " - " in cell: cod_titulo = cell.split(' - ')[0].strip(); break
            
            for r in range(5, len(df_old)):
                linhas_raw.append([df_old.iloc[r, c] for c in range(df_old.shape[1])])
                
        else:
            ws_d = load_workbook(caminho, data_only=True).active
            try: a3_valor = float(str(ws_d['A3'].value).replace(',', '.'))
            except: a3_valor = 0
            if a3_valor == 0:
                for c in range(1, 10):
                    try:
                        num = float(str(ws_d.cell(row=3, column=c).value).replace(',', '.'))
                        if num > 0: a3_valor = num; break
                    except: pass
            a3_valor = 1.0 if a3_valor == 0 else a3_valor
            
            for r in range(1, min(500, ws_d.max_row + 1)):
                linhas_raw.append([ws_d.cell(row=r, column=c).value for c in range(1, 20)])

        gatilhos_pren = REGRAS["prensados"]["descricoes_gatilho"]
        codigos_pren = REGRAS["prensados"]["codigos_gatilho"]

        for linha_bruta in linhas_raw:
            linha = [limpar(x) for x in linha_bruta]
            while len(linha) < 16: linha.append('')
            
            texto_linha = " ".join([str(x) for x in linha if x]).upper()
            if not texto_linha: continue
            
            _termos_longos = [t for t in termos_ignorar if len(t) >= 4]
            ignorar = False
            for cell in linha:
                cell_up = str(cell).upper()
                if cell_up in ["UN", "QNT", "QTD"] or any(t in cell_up for t in _termos_longos):
                    ignorar = True; break
            if ignorar: continue
            
            if any(g in texto_linha for g in gatilhos_pren) or any(c in texto_linha for c in codigos_pren):
                if bloco_atual['itens']: blocos.append(bloco_atual)
                
                texto_prensado = ""
                for cell in linha:
                    cell_up = str(cell).upper()
                    if any(g in cell_up for g in gatilhos_pren) or any(c in cell_up for c in codigos_pren):
                        texto_prensado = str(cell); break
                if not texto_prensado:
                    for cell in linha:
                        if cell: texto_prensado = str(cell); break
                            
                f_cod, f_desc = "", texto_prensado
                if " - " in texto_prensado:
                    partes = texto_prensado.split(" - ", 1)
                    f_cod, f_desc = partes[0].strip(), partes[1].strip()
                    
                bloco_atual = {'tipo': 'prensado', 'prensado_info': {1: f_cod, 3: f_desc}, 'itens': []}
                continue

            item = extrair_dados_linha_inteligente(linha, a3_valor, cod_titulo)
            if item:
                bloco_atual['itens'].append(item)
                
        if bloco_atual['itens']: 
            blocos.append(bloco_atual)
            
        return blocos, a3_valor

    except Exception as e:
        print(f"[MIGRATION ERROR] {e}")
        return [], 1.0

def mapear_rede_cache():
    cache = []
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
    padrao = re.compile(rf"^{re.escape(c)}(\s|-|_|\.|$)")
    for f, caminho in cache_rede:
        if padrao.match(f): return caminho
    return None