import os
import re
import pandas as pd
from pathlib import Path

from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero
from core.migracao import mapear_rede_cache, verificar_duplicidade_em_rede, extrair_dados_migracao
from core.excel import gerar_arquivo_excel

def f_valido(f):
    c = str(limpar(f.get(1, "")))
    a = str(f.get(2, "")).upper()
    d = str(f.get(3, "")).upper()
    mc = str(limpar(f.get(14, "")))
    filtros = REGRAS["filtros"]
    if any(x in a for x in filtros["descricoes_ignoradas"]) or \
       any(x in d for x in filtros["materiais_ignorados"]) or \
       any(x in mc for x in filtros["materiais_ignorados"]): return False
    if '*' in a and not c.startswith(tuple(filtros["prefixos_validos"])): return False
    if mc.startswith(tuple(filtros["mp_iniciais_ignoradas"])): 
        is_especial = any(m in d for m in REGRAS["especiais"]["materiais_plus_5mm"]) or re.search(r'\bTS\b', d)
        return c.startswith(tuple(filtros["prefixos_validos"])) if is_especial else False
    return c.startswith(tuple(filtros["prefixos_validos"]))
#
def is_prensado(r):
    desc = str(r.get(3, "")).upper()
    acab = str(r.get(2, "")).upper()
    cod = str(limpar(r.get(1, "")))
    return any(g in desc for g in REGRAS["prensados"]["descricoes_gatilho"]) or \
           cod in REGRAS["prensados"]["codigos_gatilho"] or \
           any(g in acab for g in REGRAS["prensados"]["acabamentos_gatilho"])

def processar_clipboard(is_teste=False):
    df = pd.read_clipboard(sep='\t', header=None, dtype=str).fillna('')
    dir_sistema = Path(REGRAS["diretorios"]["raiz"]) / REGRAS["diretorios"]["nome_pasta_sistema"]
    molde = dir_sistema / "planilha_molde.xlsm"
    
    if not molde.exists() and not is_teste: raise Exception("Molde não encontrado.")
    if df.shape[0] < 2 or df.shape[1] < 6: raise Exception("Dados insuficientes no clipboard.")

    niveis_encontrados = [str(x).count('.') for x in df[0] if re.match(r'^\d+(\.\d+)*$', str(x).strip())]
    if not niveis_encontrados: raise Exception("Estrutura de níveis não identificada.")
        
    id_p_raw = str(df.iloc[1, 1]).strip().upper()
    
    # --- TRAVA 1: PADRÃO 3 LETRAS + NÚMEROS (ZAR, REN, etc) ---
    if not re.match(r'^[A-Z]{3}\d+', id_p_raw):
        raise Exception(f"Não foi encontrado código pai do projeto. Verifique se a opção \"Incluir Selecionado\" esta ativa no PDM")

    prefixos_validos = tuple(REGRAS["filtros"]["prefixos_validos"])
    if not any(str(limpar(c)).startswith(prefixos_validos) for c in df[1]):
        raise Exception("Nenhum código primordial válido encontrado na seleção.")
    
    id_p = re.sub(r'[\\/*?:"<>|]', '-', id_p_raw)
    desktop_path = Path(os.path.join(os.path.expanduser("~"), "Desktop"))
    mapeamento = REGRAS["diretorios"]["mapeamento_pastas"]
    pasta_marca = next((v for k, v in mapeamento.items() if k in id_p), "Outros")
    pasta = (desktop_path / "TESTES_GERADOR" / id_p) if is_teste else (Path(REGRAS["diretorios"]["raiz"]) / pasta_marca / id_p)
    
    niv_pai = min(niveis_encontrados)
    if not limpar(df.iloc[1, 1]).startswith(tuple(REGRAS["filtros"]["prefixos_validos"])): niv_pai += 1

    cache_rede = {} if is_teste else mapear_rede_cache()
    cons = {}
    for _, r in df.iterrows():
        nv, cod = limpar(r[0]), limpar(r[1])
        if cod.startswith(tuple(REGRAS["filtros"]["prefixos_validos"])) and nv.count('.') == niv_pai:
            if nv not in cons: 
                cons[nv] = {'pai': r, 'blocos': [], 'qtd_p_total': 0, 'excluir_prefixos': [], 'niv_base': niv_pai}
            cons[nv]['qtd_p_total'] += float(converter_para_numero(r[5]) or 0)
    
    for _, r in df.iterrows():
        nv, cod = limpar(r[0]), limpar(r[1])
        acab = limpar(r[2])
        if nv.count('.') == niv_pai + 1:
            nv_pai_str = nv.rsplit('.', 1)[0]
            if nv_pai_str in cons:
                pai_r = cons[nv_pai_str]['pai']
                if limpar(pai_r[1]).startswith('15') and not limpar(pai_r[2]) and acab:
                    if nv not in cons:
                        qtd_p = cons[nv_pai_str]['qtd_p_total']
                        qtd_f = float(converter_para_numero(r[5]) or 0)
                        cons[nv] = {'pai': r, 'blocos': [], 'qtd_p_total': qtd_p * qtd_f, 'excluir_prefixos': [], 'niv_base': niv_pai + 1}
                        cons[nv_pai_str]['excluir_prefixos'].append(nv)

    arquivos_bloqueados = []
    processar_list = []
    for nv_p, info in cons.items():
        cod_p = limpar(info['pai'][1])
        cam_net = None if is_teste else verificar_duplicidade_em_rede(cod_p, cache_rede)
        if cam_net:
            pasta_origem = os.path.basename(os.path.dirname(cam_net))
            arquivos_bloqueados.append(f"• Peça {cod_p} (Já existe em: {pasta_origem})")
        else: processar_list.append((nv_p, info))

    arquivos_gerados_count = 0
    def garantir_pasta():
        if not os.path.exists(pasta): os.makedirs(pasta)

    for nv_p, info in processar_list:
        if not info['blocos']:
            c_p = limpar(info['pai'][1])
            niv_base = info.get('niv_base', niv_pai)
            mask = df[0].str.startswith(nv_p + ".")
            for excl in info.get('excluir_prefixos', []):
                mask = mask & ~(df[0] == excl) & ~df[0].str.startswith(excl + ".")
            desc_df = df[mask].copy()
            p_is_p = is_prensado(info['pai'])
            
            b_roots = {}
            for _, r in desc_df.iterrows():
                nv, cod = limpar(r[0]), limpar(r[1])
                if (c_p.startswith('15') and cod.startswith('15') and nv.count('.') > niv_base) or is_prensado(r):
                    pref = [p for p in b_roots.keys() if nv.startswith(p + ".")]
                    parent_qf = b_roots[max(pref, key=len)]['qf'] if pref else 1.0
                    b_roots[nv] = {'tipo': 'prensado', 'prensado_info': r, 'itens': [], 'qf': float(converter_para_numero(r[5]) or 1) * parent_qf}
            
            bloco_a = {'tipo': 'normal', 'itens': []}
            for _, r in desc_df.iterrows():
                nv = limpar(r[0])
                pref = [p for p in b_roots.keys() if nv.startswith(p + ".")]
                parent = b_roots[max(pref, key=len)] if pref else None
                if not (nv in b_roots) and f_valido(r):
                    ic = r.copy().to_dict()
                    if parent: 
                        ic['q_unitaria_fatorada'] = float(converter_para_numero(r[5]) or 0) * parent['qf']
                        parent['itens'].append(ic)
                    elif nv.count('.') == niv_base + 1: 
                        ic['q_unitaria_fatorada'] = float(converter_para_numero(r[5]) or 0)
                        bloco_a['itens'].append(ic)
            
            if bloco_a['itens']: info['blocos'].append(bloco_a)
            for br in b_roots.values(): 
                if br['itens']: info['blocos'].append(br)

            if any(len(b['itens']) > 0 for b in info['blocos']) or f_valido(info['pai']):
                garantir_pasta()
                if not info['blocos']: info['blocos'] = [{'tipo': 'normal', 'itens': [{'q_unitaria_fatorada': 1.0, **info['pai'].to_dict()}]}]
                gerar_arquivo_excel(info['pai'], info['blocos'], id_p, info['qtd_p_total'], molde, pasta, p_is_p)
                arquivos_gerados_count += 1

    return {"pasta": str(pasta), "bloqueados": arquivos_bloqueados, "aviso": "Nada gerado." if arquivos_gerados_count == 0 and not arquivos_bloqueados else None}