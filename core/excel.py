import os
import re
from pathlib import Path
from datetime import datetime
from copy import copy

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, Alignment, PatternFill

from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero, limpar_material_rigoroso, resource_path

def buscar_valor_valido(item, indices):
    """
    Percorre uma lista de índices e retorna o primeiro valor numérico válido (> 0).
    Ignora NaNs, zeros e strings vazias.
    """
    for idx in indices:
        val = item.get(idx)
        if val is not None:
            texto = str(val).strip().lower()
            if texto not in ["", "nan", "none", "0", "0,0", "0.0"]:
                num = converter_para_numero(val)
                if num and num > 0:
                    return num
    return 0

def escrever_seguro(ws, coord, valor, alinhamento=None):
    try:
        cell = ws[coord]
        if cell.__class__.__name__ == 'MergedCell':
            for r in ws.merged_cells.ranges:
                if coord in r:
                    m = ws.cell(row=r.min_row, column=r.min_col)
                    m.value = valor
                    if alinhamento: 
                        m.alignment = alinhamento
                    return
        else:
            cell.value = valor
            if alinhamento: 
                cell.alignment = alinhamento
    except Exception as e: 
        print(f"[EXCEL WARN] Falha ao escrever na célula {coord}: {e}")

def tratar_cabecalho_a1(ws, id_projeto):
    marcas_texto = REGRAS["especiais"]["marcas_texto"]
    logos_imagem = REGRAS["especiais"]["marcas_imagem"]
    id_up = str(id_projeto).upper()
    ws['A1'].value = None

    for sigla, nome in marcas_texto.items():
        if sigla in id_up:
            escrever_seguro(ws, 'A1', nome, Alignment(horizontal='center', vertical='center'))
            ws['A1'].font = Font(name='Arial', size=22, bold=True)
            return

    for sigla, arq in logos_imagem.items():
        if sigla in id_up:
            path = resource_path(f"logos/{arq}.png")
            if Path(path).exists():
                img = OpenpyxlImage(path)
                img.width, img.height = 152, 42
                ws.row_dimensions[1].height = 33
                ws.add_image(img, 'A1')
                return

    path_ing = resource_path("logos/ingecon.png")
    if Path(path_ing).exists():
        img = OpenpyxlImage(path_ing)
        img.width, img.height = 152, 42
        ws.row_dimensions[1].height = 33
        ws.add_image(img, 'A1')

def ajustar_molde_elastico(ws, num_itens):
    for r in range(1, 50):
        ws.row_dimensions[r].hidden = False

    padrao, l_rodape, quadro = 3, 9, None
    
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= l_rodape:
            if m.min_row == l_rodape: 
                quadro = {'min_col': m.min_col, 'max_col': m.max_col, 'max_row': m.max_row}
            try: ws.unmerge_cells(str(m))
            except Exception: pass
            
    if quadro and quadro['max_row'] > l_rodape: 
        ws.delete_rows(l_rodape + 1, quadro['max_row'] - l_rodape)
        
    diff = num_itens - padrao
    if diff > 0: ws.insert_rows(l_rodape, diff)
    elif diff < 0: ws.delete_rows(l_rodape + diff, abs(diff))
    
    font_pecas = Font(name='Arial', size=10)
    
    for r in range(6, 6 + num_itens):
        ws.row_dimensions[r].height = 25.5
        ws.cell(row=r, column=4).value = ws.cell(row=r, column=7).value = "X"
        
        for c in range(1, 16):
            ws.cell(row=r, column=c).font = font_pecas
            if r > 6:
                src, tgt = ws.cell(row=6, column=c), ws.cell(row=r, column=c)
                if src.has_style: tgt._style = copy(src._style)
                
    n_inicio = 6 + num_itens
    if quadro: 
        ws.merge_cells(start_row=n_inicio, start_column=quadro['min_col'], 
                       end_row=n_inicio, end_column=quadro['max_col'])
    return n_inicio


def _sufixo_dimensional(itens, fonte_espessura):
    """
    Extrai as dimensões consolidadas dos itens de um bloco prensado.
    Comprimento e Largura = maior valor entre todos os itens - 10.
    Espessura = vem do título do prensado (prensado_info ou pai), não dos itens.
    Retorna uma string " CxLxE" ou "" se faltar alguma medida.
    """
    if not itens:
        return ""
    max_c = max((buscar_valor_valido(i, [8, 15]) for i in itens), default=0)
    max_l = max((buscar_valor_valido(i, [10, 16]) for i in itens), default=0)
    v_e = buscar_valor_valido(fonte_espessura, [12, 13])
    v_c = max_c - 10 if max_c else 0
    v_l = max_l - 10 if max_l else 0
    if v_c > 0 and v_l > 0 and v_e:
        return f" {int(v_c)}X{int(v_l)}X{int(v_e)}"
    return ""

def gerar_arquivo_excel(pai, blocos, id_proj, qtd_tot, molde, pasta, pai_is_prensado):
    wb = load_workbook(molde, keep_vba=True)
    ws = wb.active
    total_itens = sum(len(b['itens']) + (1 if b['tipo'] == 'prensado' else 0) for b in blocos)
    l_obs = ajustar_molde_elastico(ws, total_itens)
    
    fill_botao = PatternFill(start_color='0078D7', end_color='0078D7', fill_type='solid')
    fill_erro = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    font_botao = Font(name='Arial', color='FFFFFF', bold=True, size=10)
    font_veio = Font(name='Arial', color='FFFFFF', bold=True, size=8)
    font_arial_12 = Font(name='Arial', size=12)
    font_arial_12_bold = Font(name='Arial', size=12, bold=True)
    
    font_material = Font(name='Arial', size=8)
    font_processo = Font(name='Arial', size=9)
    
    align_botao = Alignment(horizontal='center', vertical='center')
    align_centro = Alignment(horizontal='center', vertical='center') # ALINHAMENTO AO CENTRO

    cod_p, acab_p, desc_p = limpar(pai[1]), limpar(pai[2]), limpar(pai[3])
    tit_base = f"{cod_p}_{acab_p.strip(' _-')} - {desc_p}" if acab_p.strip(' _-') else f"{cod_p} - {desc_p}"
    tit = tit_base
    
    tratar_cabecalho_a1(ws, id_proj)
    escrever_seguro(ws, 'B3', tit)
    ws['B3'].font = font_arial_12_bold
    
    try: ws['A3'].value = float(str(qtd_tot).replace(',', '.'))
    except: ws['A3'].value = qtd_tot
    ws['A3'].font = font_arial_12
    
    ws['M2'].value = datetime.now().strftime('%d/%m/%Y')
    ws['M2'].font = Font(name='Arial1', size=12)
    
    mat_veio = REGRAS["especiais"]["materiais_com_veio"]
    mat_esp = REGRAS["especiais"]["materiais_plus_5mm"]

    row_idx = 6
    for b in blocos:
        bloco_e_prensado = (b['tipo'] == 'prensado' or pai_is_prensado)
        
        # --- ORDENAÇÃO POR ESPESSURA (Apenas itens normais, Crescente) ---
        if b['tipo'] == 'normal':
            b['itens'].sort(key=lambda i: buscar_valor_valido(i, [12, 13]))
        # -----------------------------------------------------------------

        if b['tipo'] == 'prensado':
            ws.row_dimensions[row_idx].height = 25.5
            ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=13)
            cell_h = ws.cell(row=row_idx, column=2)
            cell_h.value = f"{limpar(b['prensado_info'][1])} - {limpar(b['prensado_info'][3])}"
            cell_h.font = Font(name='Arial', bold=True, size=11)
            cell_h.alignment = align_botao
            row_idx += 1
        
        for item in b['itens']:
            r, is_mig = row_idx, item.get('is_migrado', False)
            desc_f = str(limpar(item.get(3, "")))
            if is_mig:
                d_l, m_l = item['desc_orig'], item['mat_orig']
                fita, veio = str(item['fita_orig']), item['veio_orig']
                fita_b = item.get('fita_lat')  # '-' ou '=' ou None
                fita_e = item.get('fita_top')  # '-' ou '=' ou None
                txt = f"{d_l} {m_l}".upper()
            else:
                txt = f"{str(item.get(2, ''))} {desc_f} {str(item.get(14, ''))}".upper()
                d_l, m_l = (desc_f.split(" - ", 1) if " - " in desc_f else ("-", desc_f))
                if any(m in txt for m in mat_esp): d_l = str(limpar(item.get(14, "")))
                fita_col_b = str(limpar(item.get(15, "")))  
                fita_col_e = str(limpar(item.get(16, "")))  
                fita_b = fita_col_b if fita_col_b in ['-', '='] else None
                fita_e = fita_col_e if fita_col_e in ['-', '='] else None
                _madeira_bruta = any(m in txt for m in ["MADEIRA BRUTA PINUS", "MADEIRA BRUTA TAUARI"])
                if _madeira_bruta and (fita_b or fita_e):
                    fita = "SERRA-LAM"
                elif _madeira_bruta:
                    fita = "SERRA"
                elif (fita_b or fita_e):
                    fita = "SEC-LAM"
                else:
                    fita = "SEC"
                veio = None

            plus = 5 if (any(m in txt for m in mat_esp) and not is_mig and not bloco_e_prensado) else 0
            val_fat = float(item.get('q_unitaria_fatorada', 0))
            
            v_c = buscar_valor_valido(item, [15, 8, 9])
            v_l = buscar_valor_valido(item, [16, 10, 11])
            v_a = buscar_valor_valido(item, [12, 13])

            if v_c > 0: v_c += plus
            if v_l > 0: v_l += plus
            
            ws.cell(row=r, column=1).value = f"={val_fat}*A3"
            
            font_fita = Font(name="Arial", size=12, bold=True)
            if fita_b:
                c = ws.cell(row=r, column=2)
                c.value, c.font = fita_b, font_fita
                c.data_type = 's'  
            if fita_e:
                c = ws.cell(row=r, column=5)
                c.value, c.font = fita_e, font_fita
                c.data_type = 's'  
            ws.cell(row=r, column=3).value = v_c
            ws.cell(row=r, column=6).value = v_l
            ws.cell(row=r, column=8).value = v_a

            if bloco_e_prensado and not is_mig:
                _suf_itens = _sufixo_dimensional(b['itens'], b.get('prensado_info', pai))
                d_l_final = _suf_itens.strip() if _suf_itens else d_l
            else:
                d_l_final = d_l
            ws.cell(row=r, column=12).value = d_l_final
            
            # ATRIBUIÇÃO DA FONTE 8 E ALINHAMENTO AO CENTRO PARA MATERIAL (COLUNA 9)
            ws.cell(row=r, column=9).value = limpar_material_rigoroso(m_l)
            ws.cell(row=r, column=9).font = font_material
            ws.cell(row=r, column=9).alignment = align_centro

            if is_mig: ws.cell(row=r, column=10).value = veio
            else:
                tem_v = any(m in txt for m in mat_veio)
                if "KRION" in txt and str(converter_para_numero(item.get(12, ""))) == "3": tem_v = True
                if tem_v: ws.cell(row=r, column=10).value = 1
            
            # ATRIBUIÇÃO DA FONTE 9 E ALINHAMENTO AO CENTRO PARA PROCESSO (COLUNA 11)
            ws.cell(row=r, column=11).value = fita
            ws.cell(row=r, column=11).font = font_processo
            ws.cell(row=r, column=11).alignment = align_centro
            
            _cod_raw = item.get(1, "")
            _cod_num = converter_para_numero(_cod_raw)
            ws.cell(row=r, column=13).value = _cod_num if _cod_num is not None else limpar(_cod_raw) or None
            
            if ws.cell(row=r, column=10).value == 1:
                cv = ws.cell(row=r, column=15)
                cv.value, cv.fill, cv.font, cv.alignment = "⇄", fill_botao, font_veio, align_botao
            
            if not is_mig and not bloco_e_prensado and not any(m in txt for m in mat_esp):
                cn = ws.cell(row=r, column=14)
                cn.value, cn.fill, cn.font, cn.alignment = "+5", fill_botao, font_botao, align_botao
            
            row_idx += 1
    
    escrever_seguro(ws, f"A{l_obs}", f"PROJETO DE REFERÊNCIA: {id_proj}", Alignment(horizontal='left'))
    ws[f"A{l_obs}"].font = Font(name='Arial Black', size=14, bold= True)
    
    nome_base_limpo = re.sub(r'[\\/*?:\u0022<>|]', '', tit).strip()[:120]
    caminho = os.path.join(pasta, f"{nome_base_limpo}.xlsm")
    
    caminho_tmp = caminho + ".tmp"
    try:
        wb.save(caminho_tmp)
        if os.path.exists(caminho): os.remove(caminho)
        os.rename(caminho_tmp, caminho)
    except Exception as e:
        if os.path.exists(caminho_tmp): os.remove(caminho_tmp)
        raise e