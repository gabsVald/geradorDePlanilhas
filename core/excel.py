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
#
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
    
    font_pecas = Font(name='Arial', size=14)
    
    for r in range(6, 6 + num_itens):
        ws.row_dimensions[r].height = 35.25
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

def gerar_arquivo_excel(pai, blocos, id_proj, qtd_tot, molde, pasta, pai_is_prensado):
    wb = load_workbook(molde, keep_vba=True)
    ws = wb.active
    total_itens = sum(len(b['itens']) + (1 if b['tipo'] == 'prensado' else 0) for b in blocos)
    l_obs = ajustar_molde_elastico(ws, total_itens)
    
    fill_botao = PatternFill(start_color='0078D7', end_color='0078D7', fill_type='solid')
    fill_erro = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    font_botao = Font(name='Arial', color='FFFFFF', bold=True, size=10)
    font_veio = Font(name='Arial', color='FFFFFF', bold=True, size=8)
    font_arial_14 = Font(name='Arial', size=14)
    font_arial_14_bold = Font(name='Arial', size=14, bold=True)
    align_botao = Alignment(horizontal='center', vertical='center')

    cod_p, acab_p, desc_p = limpar(pai[1]), limpar(pai[2]), limpar(pai[3])
    tit = f"{cod_p}_{acab_p.strip(' _-')} - {desc_p}" if acab_p.strip(' _-') else f"{cod_p} - {desc_p}"
    
    tratar_cabecalho_a1(ws, id_proj)
    escrever_seguro(ws, 'B3', tit)
    ws['B3'].font = font_arial_14_bold
    
    try: ws['A3'].value = float(str(qtd_tot).replace(',', '.'))
    except: ws['A3'].value = qtd_tot
    ws['A3'].font = font_arial_14
    
    ws['M2'].value = datetime.now().strftime('%d/%m/%Y')
    ws['M2'].font = Font(name='Arial', size=22)
    
    mat_veio = REGRAS["especiais"]["materiais_com_veio"]
    mat_esp = REGRAS["especiais"]["materiais_plus_5mm"]

    row_idx = 6
    for b in blocos:
        bloco_e_prensado = (b['tipo'] == 'prensado' or pai_is_prensado)
        
        if b['tipo'] == 'prensado':
            ws.row_dimensions[row_idx].height = 15.75
            ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=13)
            cell_h = ws.cell(row=row_idx, column=2)
            cell_h.value = f"{limpar(b['prensado_info'][1])} - {limpar(b['prensado_info'][3])}"
            cell_h.font = Font(name='Arial', bold=True, size=15)
            cell_h.alignment = align_botao
            row_idx += 1
        
        for item in b['itens']:
            r, is_mig = row_idx, item.get('is_migrado', False)
            desc_f = str(limpar(item.get(3, "")))
            if is_mig:
                d_l, m_l = item['desc_orig'], item['mat_orig']
                fita, veio = str(item['fita_orig']), item['veio_orig']
                fita_b, fita_e = None, None  # migrados não reescrevem colunas B/E
                txt = f"{d_l} {m_l}".upper()
            else:
                txt = f"{str(item.get(2, ''))} {desc_f} {str(item.get(14, ''))}".upper()
                d_l, m_l = (desc_f.split(" - ", 1) if " - " in desc_f else ("-", desc_f))
                if any(m in txt for m in mat_esp): d_l = str(limpar(item.get(14, "")))
                # Lê os valores brutos das colunas P (15) e Q (16) do CSV
                fita_col_b = str(limpar(item.get(15, "")))  # col P → coluna B da planilha
                fita_col_e = str(limpar(item.get(16, "")))  # col Q → coluna E da planilha
                fita_b = fita_col_b if fita_col_b in ['-', '='] else None
                fita_e = fita_col_e if fita_col_e in ['-', '='] else None
                fita = "SEC-LAM" if (fita_b or fita_e) else "SEC"
                veio = None

            plus = 5 if (any(m in txt for m in mat_esp) and not is_mig and not bloco_e_prensado) else 0
            val_fat = float(item.get('q_unitaria_fatorada', 0))
            
            # --- LÓGICA MULTI-COLUNA ROBUSTA (Correção do Acrílico) ---
            # Comprimento: Prioridade Col 15 (FB) -> Col 8 (I) -> Col 9 (J)
            v_c = buscar_valor_valido(item, [15, 8, 9])
            # Largura: Prioridade Col 16 (FB) -> Col 10 (K) -> Col 11 (L)
            v_l = buscar_valor_valido(item, [16, 10, 11])
            # Altura: Prioridade Col 12 (M) -> Col 13 (N)
            v_a = buscar_valor_valido(item, [12, 13])

            if v_c > 0: v_c += plus
            if v_l > 0: v_l += plus
            
            ws.cell(row=r, column=1).value = f"={val_fat}*A3"
            # Fitas de borda: col P (15) → coluna B (2) | col Q (16) → coluna E (5)
            font_fita = Font(name="Arial", size=28, bold=True)
            if not is_mig:
                if fita_b:
                    c = ws.cell(row=r, column=2)
                    c.value, c.font = fita_b, font_fita
                if fita_e:
                    c = ws.cell(row=r, column=5)
                    c.value, c.font = fita_e, font_fita
            ws.cell(row=r, column=3).value = v_c
            ws.cell(row=r, column=6).value = v_l
            ws.cell(row=r, column=8).value = v_a
            # ---------------------------------------------------------

            ws.cell(row=r, column=12).value = d_l
            ws.cell(row=r, column=9).value = limpar_material_rigoroso(m_l)

            if is_mig: ws.cell(row=r, column=10).value = veio
            else:
                tem_v = any(m in txt for m in mat_veio)
                if "KRION" in txt and str(converter_para_numero(item.get(12, ""))) == "3": tem_v = True
                if tem_v: ws.cell(row=r, column=10).value = 1
            
            ws.cell(row=r, column=11).value = fita
            # col 13: código do item — pode ser numérico ou texto (ex: 'PÇ 1')
            _cod_raw = item.get(1, "")
            _cod_num = converter_para_numero(_cod_raw)
            ws.cell(row=r, column=13).value = _cod_num if _cod_num is not None else limpar(_cod_raw) or None
            
            if ws.cell(row=r, column=10).value == 1:
                cv = ws.cell(row=r, column=15)
                cv.value, cv.fill, cv.font, cv.alignment = "⇄", fill_botao, font_veio, align_botao
            
            if not bloco_e_prensado and not any(m in txt for m in mat_esp):
                cn = ws.cell(row=r, column=14)
                cn.value, cn.fill, cn.font, cn.alignment = "+5", fill_botao, font_botao, align_botao
            
            row_idx += 1
    
    escrever_seguro(ws, f"A{l_obs}", f"PROJETO DE REFERÊNCIA: {id_proj}", Alignment(horizontal='left'))
    ws[f"A{l_obs}"].font = Font(name='Arial Black', size=22, bold= True)
    
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